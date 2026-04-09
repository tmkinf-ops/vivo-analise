import re
from datetime import datetime
from typing import List, Dict, Optional, Tuple

from models import db, Contrato, Conta, Comparacao, Configuracao


class Comparator:
    """Motor de comparação entre contratos e contas telefônicas."""

    def get_tolerance(self) -> Tuple[str, float]:
        """Obtém configurações de tolerância do banco de dados."""
        tipo = Configuracao.query.filter_by(chave='tolerancia_tipo').first()
        valor = Configuracao.query.filter_by(chave='tolerancia_valor').first()
        tipo_val = tipo.valor if tipo else 'percentual'
        try:
            valor_val = float(valor.valor) if valor else 5.0
        except (ValueError, TypeError):
            valor_val = 5.0
        return tipo_val, valor_val

    def normalize_phone(self, phone: str) -> str:
        """Remove não-dígitos e normaliza comprimento."""
        digits = re.sub(r'\D', '', str(phone))
        if len(digits) == 12 and digits.startswith('0'):
            digits = digits[1:]
        return digits

    def phones_match(self, phone_a: str, phone_b: str) -> bool:
        """Verifica se dois números de telefone correspondem."""
        a = self.normalize_phone(phone_a)
        b = self.normalize_phone(phone_b)
        if not a or not b:
            return False
        # Correspondência exata
        if a == b:
            return True
        # Correspondência pelos últimos 9 dígitos (móvel sem o 9 extra)
        if len(a) >= 9 and len(b) >= 9 and a[-9:] == b[-9:]:
            return True
        # Correspondência pelos últimos 8 dígitos (fixo sem DDD)
        if len(a) >= 8 and len(b) >= 8 and a[-8:] == b[-8:]:
            return True
        return False

    def is_within_tolerance(self, contratado: float, fatura: float,
                            tipo: str, tolerancia: float) -> bool:
        """Verifica se o valor da fatura está dentro da tolerância."""
        diff = abs(fatura - contratado)
        if tipo == 'fixo':
            return diff <= tolerancia
        else:  # percentual
            if contratado == 0:
                return fatura == 0
            pct = (diff / contratado) * 100
            return pct <= tolerancia

    def determine_status(self, contratado: Optional[float], fatura: float,
                         tipo: str, tolerancia: float) -> Tuple[str, str]:
        """
        Determina o status da comparação.
        Retorna (status, observação).
        """
        if contratado is None:
            return 'sem_contrato', 'Linha não encontrada na base contratual'

        diff = fatura - contratado
        diff_abs = abs(diff)

        if diff_abs < 0.005:  # praticamente igual
            return 'ok', 'Valor idêntico ao contratado'

        if self.is_within_tolerance(contratado, fatura, tipo, tolerancia):
            sinal = '+' if diff > 0 else ''
            return 'aproximado', f'Dentro da tolerância (diferença: R$ {sinal}{diff:.2f})'

        sinal = '+' if diff > 0 else ''
        if diff > 0:
            return 'divergente', f'Valor cobrado ACIMA do contratado (R$ {sinal}{diff:.2f})'
        else:
            return 'divergente', f'Valor cobrado ABAIXO do contratado (R$ {sinal}{diff:.2f})'

    def run_comparison(self, competencia: Optional[str] = None,
                       importacao_ids: Optional[List[int]] = None) -> Dict:
        """
        Executa o motor de comparação.
        Retorna dict com resultados e totais.
        """
        tipo_tol, valor_tol = self.get_tolerance()

        # Buscar contas a comparar
        query = Conta.query
        if competencia:
            query = query.filter_by(competencia=competencia)
        if importacao_ids:
            query = query.filter(Conta.importacao_id.in_(importacao_ids))
        contas = query.all()

        # Carregar contratos ativos em memória para comparação eficiente
        contratos_ativos = Contrato.query.filter_by(ativo=True).all()

        results = []
        contas_comparadas = set()

        # Processar cada conta
        for conta in contas:
            matching = []
            for contrato in contratos_ativos:
                if self.phones_match(conta.linha_telefone, contrato.linha_telefone):
                    matching.append(contrato)

            if len(matching) == 0:
                # Sem contrato correspondente
                comp = self._save_comparison(
                    conta=conta,
                    contrato=None,
                    status='sem_contrato',
                    observacao='Linha não encontrada na base contratual',
                    valor_contratado=None,
                    competencia=competencia or conta.competencia
                )
            elif len(matching) > 1:
                # Ambiguidade: múltiplos contratos
                comp = self._save_comparison(
                    conta=conta,
                    contrato=matching[0],
                    status='ambiguo',
                    observacao=f'Múltiplos contratos para esta linha ({len(matching)} registros)',
                    valor_contratado=matching[0].valor_contratado,
                    competencia=competencia or conta.competencia
                )
            else:
                contrato = matching[0]
                status, observacao = self.determine_status(
                    contrato.valor_contratado, conta.valor_fatura, tipo_tol, valor_tol
                )
                comp = self._save_comparison(
                    conta=conta,
                    contrato=contrato,
                    status=status,
                    observacao=observacao,
                    valor_contratado=contrato.valor_contratado,
                    competencia=competencia or conta.competencia
                )

            results.append(comp)
            contas_comparadas.add(self.normalize_phone(conta.linha_telefone))

        # Contratos sem fatura correspondente
        for contrato in contratos_ativos:
            phone_norm = self.normalize_phone(contrato.linha_telefone)
            found = any(
                self.phones_match(contrato.linha_telefone, p)
                for p in contas_comparadas
            )
            if not found:
                comp = self._save_comparison(
                    conta=None,
                    contrato=contrato,
                    status='sem_fatura',
                    observacao='Contrato sem fatura correspondente no período',
                    valor_contratado=contrato.valor_contratado,
                    competencia=competencia
                )
                results.append(comp)

        db.session.commit()

        # Totais
        totais = {
            'ok': sum(1 for r in results if r['status'] == 'ok'),
            'aproximado': sum(1 for r in results if r['status'] == 'aproximado'),
            'divergente': sum(1 for r in results if r['status'] == 'divergente'),
            'sem_contrato': sum(1 for r in results if r['status'] == 'sem_contrato'),
            'sem_fatura': sum(1 for r in results if r['status'] == 'sem_fatura'),
            'ambiguo': sum(1 for r in results if r['status'] == 'ambiguo'),
        }

        return {'results': results, 'totais': totais, 'total': len(results)}

    def _save_comparison(self, conta: Optional[Conta], contrato: Optional[Contrato],
                         status: str, observacao: str, valor_contratado: Optional[float],
                         competencia: Optional[str]) -> Dict:
        """Salva resultado da comparação no banco e retorna dict."""

        diferenca_valor = None
        diferenca_percentual = None
        valor_fatura = conta.valor_fatura if conta else None

        if valor_contratado is not None and valor_fatura is not None:
            diferenca_valor = round(valor_fatura - valor_contratado, 4)
            if valor_contratado != 0:
                diferenca_percentual = round((diferenca_valor / valor_contratado) * 100, 2)
            else:
                diferenca_percentual = 100.0 if valor_fatura > 0 else 0.0

        linha = (conta.linha_telefone if conta else
                 (contrato.linha_telefone if contrato else None))
        numero_contrato = contrato.numero_contrato if contrato else None
        operadora = (conta.operadora if conta else
                     (contrato.operadora if contrato else None))

        comp = Comparacao(
            contrato_id=contrato.id if contrato else None,
            conta_id=conta.id if conta else None,
            linha_telefone=linha,
            valor_contratado=valor_contratado,
            valor_fatura=valor_fatura,
            diferenca_valor=diferenca_valor,
            diferenca_percentual=diferenca_percentual,
            status_comparacao=status,
            observacao=observacao,
            competencia=competencia,
            numero_contrato=numero_contrato,
            operadora=operadora,
            data_processamento=datetime.utcnow()
        )
        db.session.add(comp)
        db.session.flush()  # Obter ID sem commit

        return {
            'id': comp.id,
            'status': status,
            'linha_telefone': linha,
            'numero_contrato': numero_contrato,
            'valor_contratado': valor_contratado,
            'valor_fatura': valor_fatura,
            'diferenca_valor': diferenca_valor,
            'diferenca_percentual': diferenca_percentual,
            'observacao': observacao,
            'competencia': competencia,
            'operadora': operadora,
            'contrato_id': contrato.id if contrato else None,
            'conta_id': conta.id if conta else None,
        }
