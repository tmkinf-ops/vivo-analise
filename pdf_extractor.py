import pdfplumber
import re
import os
from typing import List, Dict, Optional


class PDFExtractor:
    """Extrator inteligente de dados de PDFs de contratos e contas telefônicas."""

    # Padrões para números de telefone brasileiros — ordem: mais específico primeiro
    PHONE_PATTERNS = [
        # (XX) X XXXX-XXXX  — celular com 9 dígito e espaço
        r'\((\d{2})\)\s*9\s*(\d{4})[-.\s]?(\d{4})',
        # (XX) XXXXX-XXXX  — celular
        r'\((\d{2})\)\s*(\d{5})[-.\s](\d{4})',
        # (XX) XXXX-XXXX   — fixo
        r'\((\d{2})\)\s*(\d{4})[-.\s](\d{4})',
        # XX-XXXXX-XXXX / XX-XXXX-XXXX — formato Vivo com traço entre DDD e número
        r'\b(\d{2})-(\d{4,5})-(\d{4})\b',
        # XX 9 XXXX-XXXX   — sem parênteses, celular
        r'\b(\d{2})\s+9\s*(\d{4})[-.\s](\d{4})\b',
        # XX XXXXX-XXXX    — sem parênteses
        r'\b(\d{2})\s+(\d{5})[-.\s](\d{4})\b',
        # XX XXXX-XXXX     — fixo sem parênteses
        r'\b(\d{2})\s+(\d{4})[-.\s](\d{4})\b',
        # Sequência bruta 10-11 dígitos isolada (último recurso)
        r'\b([1-9]\d{9,10})\b',
    ]

    # Padrões para valores monetários brasileiros
    MONEY_PATTERNS = [
        r'R\$\s*([\d]{1,3}(?:\.[\d]{3})*,[\d]{2})',        # R$ 1.234,56
        r'R\$\s*([\d]+,[\d]{2})',                           # R$ 123,56
        r'\b([\d]{1,3}(?:\.[\d]{3})*,[\d]{2})\b',          # 1.234,56
    ]

    OPERADORAS = ['Vivo', 'Claro', 'TIM', 'Oi', 'Nextel', 'Embratel', 'Algar', 'Sercomtel']

    MESES_PT = {
        'janeiro': '01', 'fevereiro': '02', 'março': '03', 'marco': '03',
        'abril': '04', 'maio': '05', 'junho': '06', 'julho': '07',
        'agosto': '08', 'setembro': '09', 'outubro': '10',
        'novembro': '11', 'dezembro': '12'
    }

    # Máximo de páginas a processar (evita PDFs lentos com cláusulas repetidas)
    MAX_PAGES = 20

    def extract_text_and_tables(self, pdf_path: str):
        """Extrai texto e tabelas de um PDF em uma única passagem. Retorna (text, tables).
        Limita a MAX_PAGES e pula páginas com conteúdo idêntico à anterior (cláusulas repetidas).
        """
        text = ''
        tables = []
        prev_page_len = -1
        repeated_count = 0
        try:
            with pdfplumber.open(pdf_path) as pdf:
                for i, page in enumerate(pdf.pages):
                    if i >= self.MAX_PAGES:
                        break
                    page_text = page.extract_text() or ''
                    # Pula páginas repetidas (mesmo comprimento que a anterior — cláusulas padrão)
                    if len(page_text) == prev_page_len and len(page_text) > 5000:
                        repeated_count += 1
                        continue
                    prev_page_len = len(page_text)
                    repeated_count = 0
                    if page_text:
                        text += page_text + '\n'
                    try:
                        page_tables = page.extract_tables()
                        if page_tables:
                            tables.extend(page_tables)
                    except Exception:
                        pass
        except Exception as e:
            raise Exception(f'Erro ao extrair texto do PDF: {str(e)}')
        return text, tables

    # Mantidos por compatibilidade
    def extract_text(self, pdf_path: str) -> str:
        text, _ = self.extract_text_and_tables(pdf_path)
        return text

    def extract_tables(self, pdf_path: str) -> List[List]:
        _, tables = self.extract_text_and_tables(pdf_path)
        return tables

    def normalize_phone(self, phone: str) -> str:
        """Normaliza número de telefone para apenas dígitos."""
        digits = re.sub(r'\D', '', phone)
        # Remove leading 0 from area codes if present (ex: 011 -> 11)
        if len(digits) == 12 and digits.startswith('0'):
            digits = digits[1:]
        return digits

    def parse_currency(self, value_str: str) -> Optional[float]:
        """Converte string de moeda brasileira para float."""
        if not value_str:
            return None
        # Remove R$, espaços
        cleaned = re.sub(r'[R$\s]', '', str(value_str))
        # Remove pontos de milhar
        cleaned = cleaned.replace('.', '')
        # Substitui vírgula decimal por ponto
        cleaned = cleaned.replace(',', '.')
        try:
            val = float(cleaned)
            return val if val >= 0 else None
        except ValueError:
            return None

    def find_phones_in_text(self, text: str) -> List[str]:
        """Encontra todos os números de telefone no texto."""
        phones_ordered = []  # mantém ordem de aparição
        seen = set()

        for pattern in self.PHONE_PATTERNS:
            for match in re.finditer(pattern, text):
                raw = match.group()
                digits = re.sub(r'\D', '', raw)

                # Normaliza: remove 0 inicial de DDD (011 → 11)
                if len(digits) == 12 and digits.startswith('0'):
                    digits = digits[1:]

                # Remove código de país 55
                if len(digits) == 13 and digits.startswith('55'):
                    digits = digits[2:]

                # Aceita apenas 10 ou 11 dígitos e DDD válido (11–99)
                if len(digits) in (10, 11):
                    ddd = int(digits[:2])
                    if 11 <= ddd <= 99 and digits not in seen:
                        seen.add(digits)
                        phones_ordered.append(digits)

        return phones_ordered

    def find_money_values(self, text: str) -> List[float]:
        """Encontra todos os valores monetários no texto."""
        values = []
        seen = set()
        for pattern in self.MONEY_PATTERNS:
            for match in re.finditer(pattern, text):
                val = self.parse_currency(match.group(1) if match.lastindex else match.group())
                if val is not None and val > 0 and val not in seen:
                    seen.add(val)
                    values.append(val)
        return values

    def extract_contract_number(self, text: str) -> Optional[str]:
        """Extrai número do contrato do texto."""
        patterns = [
            # VMN na mesma linha: VMN: 617648
            r'[Vv][Mm][Nn]\s*[:\s]+(\d{4,12})',
            # NºdaVMN: com número 1-2 linhas depois (formato Vivo)
            r'[Nn][º°]?\s*da?\s*[Vv][Mm][Nn]\s*[:\s]*(?:[^\n]*\n){0,2}\s*(\d{4,12})',
            # Número do contrato com separador
            r'[Cc]ontrato\s*[Nn][º°\.:\s]+(\d[\d\-/]+)',
            r'[Nn][º°\.]\s*[Cc]ontrato\s*[:\s]*(\d[\d\-/]+)',
            r'[Cc]od(?:igo)?\s+[Cc]ontrato\s*[:\s]*([\w][\w\-/]+)',
        ]
        for pattern in patterns:
            match = re.search(pattern, text, re.IGNORECASE)
            if match:
                val = match.group(1).strip()
                # Rejeita se parece título (letras maiúsculas concatenadas)
                if re.match(r'^[A-Z]{5,}', val):
                    continue
                return val
        return None

    def extract_client_name(self, text: str) -> Optional[str]:
        """Tenta extrair nome do cliente do texto."""
        # Padrões com separador (: ou -)
        patterns_with_sep = [
            r'[Rr]az[\xe3a]o\s+[Ss]ocial\s*[:\-]\s*(.+)',
            r'[Cc]liente\s*[:\-]\s*(.+)',
            r'[Ee]mpresa\s*[:\-]\s*(.+)',
            r'[Nn]ome\s*[:\-]\s*(.+)',
        ]
        # Padrão para Razão Social sem separador (ex: "RazãoSocial CNPJ" e valor na linha seguinte)
        pattern_razao_social_next_line = r'[Rr]az[\xe3a]o\s*[Ss]ocial.*?\n(.+)'

        candidates = []
        for pattern in patterns_with_sep:
            for m in re.finditer(pattern, text):
                candidates.append(m.group(1).split('\n')[0].strip())
        m = re.search(pattern_razao_social_next_line, text)
        if m:
            candidates.append(m.group(1).split('\n')[0].strip())

        # Validação: rejeita emails, texto muito longo/curto ou com padrões de cláusula
        bad_words = ('represent', 'contrato', 'presente', 'adesão', 'cláusula', 'objeto')
        for name in candidates:
            name = re.sub(r'\s+', ' ', name).strip()
            if not name or len(name) < 3 or len(name) > 120:
                continue
            if '@' in name:
                continue
            if any(w in name.lower() for w in bad_words):
                continue
            return name
        return None

    def extract_operator(self, text: str) -> Optional[str]:
        """Identifica a operadora de telecomunicações."""
        text_upper = text.upper()
        for op in self.OPERADORAS:
            if op.upper() in text_upper:
                return op
        return None

    def extract_competencia(self, text: str) -> Optional[str]:
        """Extrai a competência (período) da fatura."""
        # Formato MM/AAAA
        match = re.search(r'\b(0[1-9]|1[0-2])[/](20\d{2})\b', text)
        if match:
            return f"{match.group(2)}-{match.group(1)}"

        # Formato AAAA-MM
        match = re.search(r'\b(20\d{2})[-](0[1-9]|1[0-2])\b', text)
        if match:
            return f"{match.group(1)}-{match.group(2)}"

        # Formato por extenso: "Janeiro de 2025" ou "Janeiro/2025"
        for mes_nome, mes_num in self.MESES_PT.items():
            pattern = rf'\b{mes_nome}\s+(?:de\s+)?(20\d{{2}})\b'
            match = re.search(pattern, text, re.IGNORECASE)
            if match:
                return f"{match.group(1)}-{mes_num}"

        return None

    def extract_invoice_number(self, text: str) -> Optional[str]:
        """Extrai o número da fatura."""
        patterns = [
            r'[Ff]atura\s+[Nn][°º\.:\s]*(\d[\d\-/]+)',
            r'[Nn][°º\.]\s+(?:da\s+)?[Ff]atura\s*[:\s]*(\d[\d\-/]+)',
            r'[Nn]ota\s+[Ff]iscal\s+[Nn][°º\.:\s]*(\d[\d\-/]+)',
            r'[Cc][ó]digo\s+(?:da\s+)?[Ff]atura\s*[:\s]*(\d[\d\-/]+)',
            r'[Ii]nvoice\s+[Nn][°º\.:\s]*(\d[\d\-/]+)',
        ]
        for pattern in patterns:
            match = re.search(pattern, text)
            if match:
                return match.group(1).strip()
        return None

    def _associate_phone_value(self, phone: str, lines: List[str]) -> Optional[float]:
        """Tenta associar um valor monetário a um número de telefone no texto."""
        _, value = self._associate_phone_plan_value(phone, lines)
        return value

    def _associate_phone_plan_value(self, phone: str, lines: List[str]) -> tuple:
        """Retorna (plano, valor) associados a um número de telefone no texto."""
        phone_digits = re.sub(r'\D', '', phone)
        suffix = phone_digits[-8:]

        # Coleta todas as linhas onde o telefone aparece
        found_indices = []
        for i, line in enumerate(lines):
            line_digits = re.sub(r'\D', '', line)
            if suffix in line_digits:
                found_indices.append(i)

        # Prioridade 1: valor na mesma linha (busca texto entre telefone e valor)
        for i in found_indices:
            line = lines[i]
            for pattern in self.PHONE_PATTERNS:
                for m in re.finditer(pattern, line):
                    m_digits = re.sub(r'\D', '', m.group())
                    if len(m_digits) >= 10 and suffix in m_digits:
                        after = line[m.end():]
                        # Extrai plano: texto entre o telefone e o próximo valor ou telefone
                        money = self.find_money_values(after)
                        if money:
                            # O plano é o texto entre o telefone e o valor monetário
                            money_match = re.search(
                                r'([\d]{1,3}(?:\.[\d]{3})*,[\d]{2})', after)
                            if money_match:
                                plano = after[:money_match.start()].strip()
                                # Limpa prefixos/sufixos comuns
                                plano = re.sub(r'^[\s\-:]+|[\s\-:]+$', '', plano)
                                return (plano or '', money[0])
                            return ('', money[0])
                        break

        # Prioridade 2: valor nas próximas 3 linhas
        for i in found_indices:
            for j in range(i + 1, min(len(lines), i + 4)):
                money = self.find_money_values(lines[j])
                if money:
                    return ('', money[-1])

        return ('', None)

    def _extract_from_tables(self, tables: List, contract_number: Optional[str],
                              client_name: Optional[str], operator: Optional[str],
                              filename: str, processed: set, is_invoice: bool,
                              competencia: Optional[str] = None,
                              numero_fatura: Optional[str] = None) -> List[Dict]:
        """Extrai dados de telefone/valor a partir das tabelas do PDF."""
        results = []
        for table in (tables or []):
            if not table:
                continue
            for row in table:
                if not row:
                    continue
                row_text = ' '.join(str(c) for c in row if c)
                row_phones = self.find_phones_in_text(row_text)
                row_money  = self.find_money_values(row_text)
                for p in row_phones:
                    if p not in processed:
                        processed.add(p)
                        if is_invoice:
                            results.append({
                                'linha_telefone': p,
                                'valor_fatura': row_money[-1] if row_money else None,
                                'competencia': competencia or '',
                                'operadora': operator or '',
                                'numero_fatura': numero_fatura or '',
                                'arquivo_pdf_origem': filename,
                            })
                        else:
                            results.append({
                                'numero_contrato': contract_number or '',
                                'linha_telefone': p,
                                'valor_contratado': row_money[-1] if row_money else None,
                                'cliente': client_name or '',
                                'operadora': operator or '',
                                'arquivo_pdf_origem': filename,
                            })
        return results

    def get_raw_text(self, pdf_path: str) -> str:
        """Retorna o texto bruto extraído do PDF (para diagnóstico)."""
        text, _ = self.extract_text_and_tables(pdf_path)
        return text

    def extract_from_contract_pdf(self, pdf_path: str) -> List[Dict]:
        """
        Extrai dados de contrato telefônico de um PDF.
        Retorna lista de dicts com os dados extraídos.
        """
        filename = os.path.basename(pdf_path)
        try:
            text, tables = self.extract_text_and_tables(pdf_path)
        except Exception as e:
            return [{'erro': str(e), 'arquivo_pdf_origem': filename}]
        contract_number = self.extract_contract_number(text)
        client_name     = self.extract_client_name(text)
        operator        = self.extract_operator(text)
        lines           = text.split('\n')
        processed       = set()
        results         = []

        # --- Estratégia 1: linha a linha (texto corrido) ---
        phones = self.find_phones_in_text(text)
        for phone in phones:
            if phone in processed:
                continue
            processed.add(phone)
            value = self._associate_phone_value(phone, lines)
            results.append({
                'numero_contrato': contract_number or '',
                'linha_telefone':  phone,
                'valor_contratado': value,
                'cliente':   client_name or '',
                'operadora': operator    or '',
                'arquivo_pdf_origem': filename,
            })

        # --- Estratégia 2: tabelas ---
        table_results = self._extract_from_tables(
            tables, contract_number, client_name, operator,
            filename, processed, is_invoice=False,
        )
        results.extend(table_results)

        # --- Estratégia 3: janela deslizante de 3 linhas ---
        # Varre pares de linhas procurando (telefone próximo de valor)
        if not results:
            window = 4
            for i, line in enumerate(lines):
                chunk = '\n'.join(lines[i: i + window])
                chunk_phones = self.find_phones_in_text(chunk)
                chunk_money  = self.find_money_values(chunk)
                for p in chunk_phones:
                    if p not in processed:
                        processed.add(p)
                        results.append({
                            'numero_contrato': contract_number or '',
                            'linha_telefone':  p,
                            'valor_contratado': chunk_money[-1] if chunk_money else None,
                            'cliente':   client_name or '',
                            'operadora': operator    or '',
                            'arquivo_pdf_origem': filename,
                        })

        return results

    def extract_from_invoice_pdf(self, pdf_path: str) -> List[Dict]:
        """
        Extrai dados de conta/fatura telefônica de um PDF.
        Retorna lista de dicts com os dados extraídos.
        """
        filename = os.path.basename(pdf_path)
        try:
            text, tables = self.extract_text_and_tables(pdf_path)
        except Exception as e:
            return [{'erro': str(e), 'arquivo_pdf_origem': filename}]
        operator      = self.extract_operator(text)
        competencia   = self.extract_competencia(text)
        numero_fatura = self.extract_invoice_number(text)
        lines         = text.split('\n')
        processed     = set()
        results       = []

        # --- Estratégia 1: linha a linha ---
        phones = self.find_phones_in_text(text)
        for phone in phones:
            if phone in processed:
                continue
            processed.add(phone)
            plano, value = self._associate_phone_plan_value(phone, lines)
            results.append({
                'linha_telefone': phone,
                'valor_fatura':   value,
                'plano':          plano or '',
                'competencia':    competencia   or '',
                'operadora':      operator      or '',
                'numero_fatura':  numero_fatura or '',
                'arquivo_pdf_origem': filename,
            })

        # --- Estratégia 2: tabelas ---
        table_results = self._extract_from_tables(
            tables, None, None, operator,
            filename, processed, is_invoice=True,
            competencia=competencia, numero_fatura=numero_fatura,
        )
        results.extend(table_results)

        # --- Estratégia 3: janela deslizante ---
        if not results:
            window = 4
            for i, line in enumerate(lines):
                chunk = '\n'.join(lines[i: i + window])
                chunk_phones = self.find_phones_in_text(chunk)
                chunk_money  = self.find_money_values(chunk)
                for p in chunk_phones:
                    if p not in processed:
                        processed.add(p)
                        results.append({
                            'linha_telefone': p,
                            'valor_fatura':   chunk_money[-1] if chunk_money else None,
                            'plano':          '',
                            'competencia':    competencia   or '',
                            'operadora':      operator      or '',
                            'numero_fatura':  numero_fatura or '',
                            'arquivo_pdf_origem': filename,
                        })

        return results

    # ------------------------------------------------------------------
    # EXTRATOR ESPECIALIZADO: Fatura Vivo (Número / Plano / Valor Total)
    # ------------------------------------------------------------------
    def extract_fatura_vivo(self, pdf_path: str) -> dict:
        """
        Lê uma fatura Vivo e extrai linhas no formato:
          Número Vivo | Plano | Valor Total R$
        Retorna dict com keys: 'linhas', 'competencia', 'operadora', 'numero_fatura', 'raw_text'
        """
        filename = os.path.basename(pdf_path)
        try:
            text, tables = self.extract_text_and_tables(pdf_path)
        except Exception as e:
            return {'linhas': [], 'erro': str(e), 'raw_text': ''}

        operadora   = self.extract_operator(text) or 'Vivo'
        competencia = self.extract_competencia(text) or ''
        nf          = self.extract_invoice_number(text) or ''
        lines       = [l for l in text.split('\n') if l.strip()]
        processed   = set()
        results     = []

        # --- Estratégia 1: tabelas (mais fiel ao layout real do PDF) ---
        for table in (tables or []):
            if not table:
                continue
            # Detecta cabeçalho da tabela (busca linha com "Número" e "Plano" e "Valor")
            header_row = -1
            for ri, row in enumerate(table):
                if not row:
                    continue
                row_str = ' '.join(str(c or '') for c in row).lower()
                if ('número' in row_str or 'numero' in row_str or 'vivo' in row_str) and \
                   ('plano' in row_str) and ('valor' in row_str):
                    header_row = ri
                    break

            # Mapeia colunas pelo cabeçalho
            if header_row >= 0:
                hdr = [str(c or '').lower() for c in table[header_row]]
                def col_idx(*keywords):
                    for k in keywords:
                        for i, h in enumerate(hdr):
                            if k in h:
                                return i
                    return -1
                ci_phone = col_idx('número', 'numero', 'vivo', 'linha', 'tel')
                ci_plano = col_idx('plano', 'serviço', 'servico', 'produto')
                ci_valor = col_idx('valor', 'total', 'r$')

                for row in table[header_row + 1:]:
                    if not row:
                        continue
                    phone_raw = str(row[ci_phone] or '') if ci_phone >= 0 else ''
                    plano_raw = str(row[ci_plano] or '') if ci_plano >= 0 else ''
                    valor_raw = str(row[ci_valor] or '') if ci_valor >= 0 else ''

                    phone_digits = None
                    if phone_raw.strip():
                        found = self.find_phones_in_text(phone_raw)
                        if found:
                            phone_digits = found[0]

                    if phone_digits and phone_digits not in processed:
                        processed.add(phone_digits)
                        valor = self.parse_currency(valor_raw) or self.find_money_values(valor_raw + ' ' + plano_raw)
                        if isinstance(valor, list):
                            valor = valor[-1] if valor else None
                        results.append({
                            'numero_vivo':  phone_digits,
                            'plano':        plano_raw.strip(),
                            'valor_total':  valor,
                        })
            else:
                # Tabela sem cabeçalho reconhecido: varre células procurando telefone
                for row in table:
                    if not row:
                        continue
                    row_str = ' '.join(str(c or '') for c in row)
                    phones = self.find_phones_in_text(row_str)
                    money  = self.find_money_values(row_str)
                    # Tenta identificar plano (texto longo entre maiúsculas)
                    plano_match = re.search(r'[A-Z][A-Z\s]{10,}', row_str)
                    plano = plano_match.group().strip() if plano_match else ''
                    for p in phones:
                        if p not in processed:
                            processed.add(p)
                            results.append({
                                'numero_vivo': p,
                                'plano':       plano,
                                'valor_total': money[-1] if money else None,
                            })

        # --- Estratégia 2: texto linha a linha ---
        # Padrão: linha contendo telefone no formato XX-XXXXX-XXXX seguido de plano e valor
        phone_line_re = re.compile(
            r'(\d{2}-\d{4,5}-\d{4}|\(\d{2}\)\s*\d{4,5}[-\s]\d{4})'
            r'(.*?)'
            r'([\d]{1,3}(?:\.[\d]{3})*,[\d]{2})'
        )
        for line in lines:
            m = phone_line_re.search(line)
            if m:
                phones = self.find_phones_in_text(m.group(1))
                valor  = self.parse_currency(m.group(3))
                plano  = m.group(2).strip()
                if not plano:
                    plano_match = re.search(r'[A-Z][A-Z\s]{10,}', line)
                    plano = plano_match.group().strip() if plano_match else ''
                for p in phones:
                    if p not in processed:
                        processed.add(p)
                        results.append({
                            'numero_vivo': p,
                            'plano':       plano,
                            'valor_total': valor,
                        })

        # --- Estratégia 3: janela deslizante de 3 linhas ---
        if not results:
            for i, line in enumerate(lines):
                phones = self.find_phones_in_text(line)
                if not phones:
                    continue
                chunk = '\n'.join(lines[i:i+4])
                money = self.find_money_values(chunk)
                plano_m = re.search(r'[A-Z][A-Z\s]{10,}', chunk)
                plano = plano_m.group().strip() if plano_m else ''
                for p in phones:
                    if p not in processed:
                        processed.add(p)
                        results.append({
                            'numero_vivo': p,
                            'plano':       plano,
                            'valor_total': money[-1] if money else None,
                        })

        return {
            'linhas':       results,
            'competencia':  competencia,
            'operadora':    operadora,
            'numero_fatura': nf,
            'filename':     filename,
            'raw_text':     text[:3000],
        }

