from flask_sqlalchemy import SQLAlchemy
from datetime import datetime

db = SQLAlchemy()


class Configuracao(db.Model):
    __tablename__ = 'configuracoes'

    id = db.Column(db.Integer, primary_key=True)
    chave = db.Column(db.String(50), unique=True, nullable=False)
    valor = db.Column(db.String(500), nullable=False)
    descricao = db.Column(db.String(300))

    def to_dict(self):
        return {
            'id': self.id,
            'chave': self.chave,
            'valor': self.valor,
            'descricao': self.descricao
        }


class PlanoPreco(db.Model):
    """Tabela de preços dos planos. Ex: 'SMART EMPRESAS 8GB TE' → R$ 37,49"""
    __tablename__ = 'planos_precos'

    id = db.Column(db.Integer, primary_key=True, autoincrement=True)
    nome_plano = db.Column(db.String(300), nullable=False, unique=True, index=True)
    valor_contrato = db.Column(db.Float, nullable=False)
    data_criacao = db.Column(db.DateTime, default=datetime.utcnow)

    def to_dict(self):
        return {
            'id': self.id,
            'nome_plano': self.nome_plano,
            'valor_contrato': self.valor_contrato,
            'data_criacao': self.data_criacao.strftime('%d/%m/%Y %H:%M') if self.data_criacao else None,
        }


class FaturaLinha(db.Model):
    """Linha importada de uma fatura PDF. Vinculada ao plano para comparação."""
    __tablename__ = 'fatura_linhas'

    id = db.Column(db.Integer, primary_key=True, autoincrement=True)
    numero_vivo = db.Column(db.String(20), nullable=False, index=True)
    plano = db.Column(db.String(300))
    valor_fatura = db.Column(db.Float)
    valor_contrato = db.Column(db.Float)       # preenchido automaticamente via PlanoPreco
    diferenca = db.Column(db.Float)             # valor_fatura - valor_contrato
    status = db.Column(db.String(20))           # ok | divergente
    arquivo_origem = db.Column(db.String(400))
    competencia = db.Column(db.String(7), index=True)
    data_importacao = db.Column(db.DateTime, default=datetime.utcnow)

    def to_dict(self):
        return {
            'id': self.id,
            'numero_vivo': self.numero_vivo,
            'plano': self.plano,
            'valor_fatura': self.valor_fatura,
            'valor_contrato': self.valor_contrato,
            'diferenca': self.diferenca,
            'status': self.status,
            'arquivo_origem': self.arquivo_origem,
            'competencia': self.competencia,
            'data_importacao': self.data_importacao.strftime('%d/%m/%Y %H:%M') if self.data_importacao else None,
        }


class Importacao(db.Model):
    __tablename__ = 'importacoes'

    id = db.Column(db.Integer, primary_key=True, autoincrement=True)
    tipo = db.Column(db.String(20), nullable=False)  # 'contrato' ou 'conta'
    arquivo_nome = db.Column(db.String(300))
    data_importacao = db.Column(db.DateTime, default=datetime.utcnow)
    status = db.Column(db.String(20), default='concluido')
    total_registros = db.Column(db.Integer, default=0)
    observacoes = db.Column(db.Text)

    contratos = db.relationship('Contrato', backref='importacao', lazy=True)
    contas = db.relationship('Conta', backref='importacao', lazy=True)

    def to_dict(self):
        return {
            'id': self.id,
            'tipo': self.tipo,
            'arquivo_nome': self.arquivo_nome,
            'data_importacao': self.data_importacao.strftime('%d/%m/%Y %H:%M') if self.data_importacao else None,
            'status': self.status,
            'total_registros': self.total_registros,
            'observacoes': self.observacoes
        }


class Contrato(db.Model):
    __tablename__ = 'contratos'

    id = db.Column(db.Integer, primary_key=True, autoincrement=True)
    numero_contrato = db.Column(db.String(100), index=True)
    linha_telefone = db.Column(db.String(20), nullable=False, index=True)
    valor_contratado = db.Column(db.Float, nullable=False)
    cliente = db.Column(db.String(300))
    operadora = db.Column(db.String(100))
    vigencia_inicio = db.Column(db.Date)
    vigencia_fim = db.Column(db.Date)
    arquivo_pdf_origem = db.Column(db.String(400))
    data_importacao = db.Column(db.DateTime, default=datetime.utcnow)
    observacoes = db.Column(db.Text)
    ativo = db.Column(db.Boolean, default=True)
    importacao_id = db.Column(db.Integer, db.ForeignKey('importacoes.id'))

    def to_dict(self):
        return {
            'id': self.id,
            'numero_contrato': self.numero_contrato,
            'linha_telefone': self.linha_telefone,
            'valor_contratado': self.valor_contratado,
            'cliente': self.cliente,
            'operadora': self.operadora,
            'vigencia_inicio': self.vigencia_inicio.strftime('%d/%m/%Y') if self.vigencia_inicio else None,
            'vigencia_fim': self.vigencia_fim.strftime('%d/%m/%Y') if self.vigencia_fim else None,
            'arquivo_pdf_origem': self.arquivo_pdf_origem,
            'data_importacao': self.data_importacao.strftime('%d/%m/%Y %H:%M') if self.data_importacao else None,
            'observacoes': self.observacoes,
            'ativo': self.ativo,
            'importacao_id': self.importacao_id
        }


class Conta(db.Model):
    __tablename__ = 'contas'

    id = db.Column(db.Integer, primary_key=True, autoincrement=True)
    linha_telefone = db.Column(db.String(20), nullable=False, index=True)
    valor_fatura = db.Column(db.Float, nullable=False)
    competencia = db.Column(db.String(7), index=True)  # YYYY-MM
    operadora = db.Column(db.String(100))
    numero_fatura = db.Column(db.String(100))
    arquivo_pdf_origem = db.Column(db.String(400))
    data_importacao = db.Column(db.DateTime, default=datetime.utcnow)
    importacao_id = db.Column(db.Integer, db.ForeignKey('importacoes.id'))

    def to_dict(self):
        return {
            'id': self.id,
            'linha_telefone': self.linha_telefone,
            'valor_fatura': self.valor_fatura,
            'competencia': self.competencia,
            'operadora': self.operadora,
            'numero_fatura': self.numero_fatura,
            'arquivo_pdf_origem': self.arquivo_pdf_origem,
            'data_importacao': self.data_importacao.strftime('%d/%m/%Y %H:%M') if self.data_importacao else None,
            'importacao_id': self.importacao_id
        }


class Comparacao(db.Model):
    __tablename__ = 'comparacoes'

    id = db.Column(db.Integer, primary_key=True, autoincrement=True)
    contrato_id = db.Column(db.Integer, db.ForeignKey('contratos.id'), nullable=True)
    conta_id = db.Column(db.Integer, db.ForeignKey('contas.id'), nullable=True)
    linha_telefone = db.Column(db.String(20), index=True)
    valor_contratado = db.Column(db.Float)
    valor_fatura = db.Column(db.Float)
    diferenca_valor = db.Column(db.Float)
    diferenca_percentual = db.Column(db.Float)
    # ok | aproximado | divergente | sem_contrato | sem_fatura | ambiguo
    status_comparacao = db.Column(db.String(20), index=True)
    observacao = db.Column(db.Text)
    data_processamento = db.Column(db.DateTime, default=datetime.utcnow)
    competencia = db.Column(db.String(7), index=True)
    numero_contrato = db.Column(db.String(100))
    operadora = db.Column(db.String(100))

    contrato = db.relationship('Contrato', backref='comparacoes', lazy='select')
    conta = db.relationship('Conta', backref='comparacoes', lazy='select')

    def to_dict(self):
        return {
            'id': self.id,
            'contrato_id': self.contrato_id,
            'conta_id': self.conta_id,
            'linha_telefone': self.linha_telefone,
            'valor_contratado': self.valor_contratado,
            'valor_fatura': self.valor_fatura,
            'diferenca_valor': self.diferenca_valor,
            'diferenca_percentual': self.diferenca_percentual,
            'status_comparacao': self.status_comparacao,
            'observacao': self.observacao,
            'data_processamento': self.data_processamento.strftime('%d/%m/%Y %H:%M') if self.data_processamento else None,
            'competencia': self.competencia,
            'numero_contrato': self.numero_contrato,
            'operadora': self.operadora
        }


# ============================================================
# CADASTRO COMPLETO DE LINHAS
# ============================================================
class CadastroLinha(db.Model):
    """Cadastro completo de linhas telefônicas com dados do funcionário e centro de custo."""
    __tablename__ = 'cadastro_linhas'

    id = db.Column(db.Integer, primary_key=True, autoincrement=True)
    numero_telefone = db.Column(db.String(20), nullable=False, index=True)
    operadora = db.Column(db.String(100))
    vencimento = db.Column(db.Integer)  # dia do mês (1-31)
    numero_conta = db.Column(db.String(100))
    matricula_funcionario = db.Column(db.String(100))
    nome_funcionario = db.Column(db.String(300))
    centro_custo = db.Column(db.String(100))
    nome_centro_custo = db.Column(db.String(300))
    plano = db.Column(db.String(300))
    valor_plano = db.Column(db.Float)
    empresa = db.Column(db.String(300))  # nome da unidade
    conferencia = db.Column(db.String(20), default='pendente')  # ok | divergente | pendente
    valor_contrato = db.Column(db.Float)
    valor_fatura = db.Column(db.Float)
    diferenca = db.Column(db.Float)
    data_criacao = db.Column(db.DateTime, default=datetime.utcnow)

    def to_dict(self):
        return {
            'id': self.id,
            'numero_telefone': self.numero_telefone,
            'operadora': self.operadora,
            'vencimento': self.vencimento,
            'numero_conta': self.numero_conta,
            'matricula_funcionario': self.matricula_funcionario,
            'nome_funcionario': self.nome_funcionario,
            'centro_custo': self.centro_custo,
            'nome_centro_custo': self.nome_centro_custo,
            'plano': self.plano,
            'valor_plano': self.valor_plano,
            'empresa': self.empresa,
            'conferencia': self.conferencia,
            'valor_contrato': self.valor_contrato,
            'valor_fatura': self.valor_fatura,
            'diferenca': self.diferenca,
            'data_criacao': self.data_criacao.strftime('%d/%m/%Y %H:%M') if self.data_criacao else None,
        }


# ============================================================
# COOPERNAC
# ============================================================
class CoopernacVoz(db.Model):
    """Números Coopernac — Voz."""
    __tablename__ = 'coopernac_voz'

    id = db.Column(db.Integer, primary_key=True, autoincrement=True)
    numero = db.Column(db.String(50))
    descricao = db.Column(db.String(300))
    total = db.Column(db.Float, default=0)

    def to_dict(self):
        return {'id': self.id, 'numero': self.numero, 'descricao': self.descricao, 'total': self.total}


class CoopernacDados(db.Model):
    """Números Coopernac — Dados."""
    __tablename__ = 'coopernac_dados'

    id = db.Column(db.Integer, primary_key=True, autoincrement=True)
    numero = db.Column(db.String(50))
    descricao = db.Column(db.String(300))
    total = db.Column(db.Float, default=0)

    def to_dict(self):
        return {'id': self.id, 'numero': self.numero, 'descricao': self.descricao, 'total': self.total}


class CoopernacResumo(db.Model):
    """Resumo Coopernac — soma de planos voz e dados."""
    __tablename__ = 'coopernac_resumo'

    id = db.Column(db.Integer, primary_key=True, autoincrement=True)
    valores = db.Column(db.Float, default=0)
    descricao = db.Column(db.String(300))
    observacao = db.Column(db.String(500))

    def to_dict(self):
        return {'id': self.id, 'valores': self.valores, 'descricao': self.descricao, 'observacao': self.observacao}
