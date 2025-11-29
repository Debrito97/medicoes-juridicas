import streamlit as st
import pandas as pd
from datetime import datetime, date
import tempfile
import os
import re
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.drawing.image import Image as OpenPyXLImage
from openpyxl.workbook.protection import WorkbookProtection
import io

# ============================================================
# CONFIGURAÇÃO INICIAL DO STREAMLIT
# ============================================================
st.set_page_config(
    page_title="Sistema de Medições Jurídicas",
    page_icon="⚖️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ============================================================
# FUNÇÕES DE APOIO
# ============================================================
def filtrar_numeros(texto: str) -> str:
    """Remove todos os caracteres não numéricos de uma string."""
    return "".join(ch for ch in texto if ch.isdigit())

def cnpj_valido(cnpj: str) -> bool:
    """Verifica se o CNPJ é válido (simplificado)."""
    cnpj = filtrar_numeros(cnpj)
    if len(cnpj) != 14 or cnpj == cnpj[0] * 14:
        return False

    pesos1 = [5, 4, 3, 2, 9, 8, 7, 6, 5, 4, 3, 2]
    pesos2 = [6, 5, 4, 3, 2, 9, 8, 7, 6, 5, 4, 3, 2]

    # Cálculo do primeiro dígito verificador
    soma = sum(int(cnpj[i]) * pesos1[i] for i in range(12))
    resto = soma % 11
    digito1 = 0 if resto < 2 else 11 - resto

    # Cálculo do segundo dígito verificador
    soma = sum(int(cnpj[i]) * pesos2[i] for i in range(13))
    resto = soma % 11
    digito2 = 0 if resto < 2 else 11 - resto

    return cnpj[12] == str(digito1) and cnpj[13] == str(digito2)

def formatar_moeda(valor):
    """Formata valor para moeda brasileira"""
    if not valor:
        return "0,00"
    
    texto_puro = re.sub(r'[^0-9]', '', str(valor))
    
    if not texto_puro:
        return "0,00"
    
    if len(texto_puro) < 3:
        texto_puro = '0' * (3 - len(texto_puro)) + texto_puro

    inteiro = texto_puro[:-2]
    decimal = texto_puro[-2:]

    if not inteiro: 
        inteiro = "0"

    try:
        inteiro_num = int(inteiro)
        inteiro = str(inteiro_num)
    except ValueError:
        inteiro = "0"

    inteiro_formatado = f"{int(inteiro):,}".replace(",", ".")
    return f"{inteiro_formatado},{decimal}"

# ============================================================
# DADOS DO SISTEMA
# ============================================================
MAPA_TIPO_COBRANCA = {
    "Consultorias": "CONS",
    "Despesas": "DESP",
    "Êxito": "EXITO",
    "Honorários Advocatícios (Prolabore)": "HON",
    "Parecer": "PARC",
    "Perícia": "PERI",
    "Publicações Legais": "PUB"
}

MAPA_MATERIA = {
    "4819": "4819",
    "Ambiental": "AMB",
    "Cálculos Trabalhistas": "CALCTR",
    "Cível": "CIV",
    "Contratos": "CONT",
    "Debêntures": "DEBS",
    "Fundiário": "FND",
    "Novos Negócios": "NNG",
    "Regulatório": "REG",
    "Publicações Legais": "PUBLEG",
    "BPO": "BPO",
    "Acompanhamento de Processos Judiciais": "ACPROC",
    "Recorte Eletrônico dos diários oficiais": "REDO",
    "Trabalhista": "TRB",
    "Tributário": "TRI",
    "Aduaneiro": "ADU",
    "Levantamento de Processos": "LEVPROC",
    "Societário": "SOC"
}

MATERIAS_JURIDICAS = [
    "4819", "Ambiental", "Cálculos Trabalhistas", "Cível", "Contratos",
    "Debêntures", "Fundiário", "Novos Negócios", "Regulatório", 
    "Publicações Legais", "BPO", "Acompanhamento de Processos Judiciais",
    "Recorte Eletrônico dos diários oficiais", "Trabalhista", "Tributário",
    "Aduaneiro", "Levantamento de Processos", "Societário"
]

TIPOS_COBRANCA = [
    "Consultorias", "Despesas", "Êxito", "Honorários Advocatícios (Prolabore)",
    "Parecer", "Perícia", "Publicações Legais"
]

EMPRESAS = [
    "ISA ENERGIA BRASIL", "Interligação Elétrica Evrecy", 
    "Interligação Elétrica Minas Gerais", "Interligação Elétrica Norte Nordeste",
    "Interligação Elétrica Pinheiros", "Interligação Elétrica Sul",
    "Interligação Elétrica Serra Japi", "Interligação Elétrica Itaúnas",
    "Interligação Elétrica Itapura", "Interligação Elétrica Aguapeí",
    "Interligação Elétrica Itaquerê", "Interligação Elétrica Tibagi",
    "Interligação Elétrica Ivaí", "Interligação Elétrica Biguaçu",
    "Interligação Elétrica Jaguar 6", "Interligação Elétrica Jaguar 8",
    "Interligação Elétrica Jaguar 9", "Interligação Elétrica Riacho Grande"
]

ADVOGADOS = [
    "Andrea Mazzaro Carlos de Vincenti", "Carlos Lopes",
    "Emerson Rodrigues do Nascimento", "Eric Tadao Pagani Fukai",
    "Erica Barbeiro Travassos", "Francisco Ricardo Tavian",
    "Gilvan Aparecido dos Santos", "Leonam Ricardo Alcantara Francisconi",
    "Leonardo Lupercio Garcia Martins", "Leonardo Silva Merces",
    "Letticia Pinheiro de Oliveira Barros", "Luciana Semenzato Garcia",
    "Marjorie Merida Chiesa", "Natalia Mendonca Goncalves",
    "Pedro Henrique Ribeiro e Silva", "Ricardo de Oliveira Beninca",
    "Rita Halabian"
]

PROJETOS_POR_EMPRESA = {
    "Interligação Elétrica Aguapeí": {
        "Projeto vinculado": ["IE Interligação Elétrica Aguapeí"],
        "Trecho": {"IE Interligação Elétrica Aguapeí": [""]}
    },
    "Interligação Elétrica Biguaçu": {
        "Projeto vinculado": ["IE Interligação Elétrica Biguaçu"],
        "Trecho": {"IE Interligação Elétrica Biguaçu": [""]}
    },
    "Interligação Elétrica Evrecy": {
        "Projeto vinculado": ["Minuano"],
        "Trecho": {"Minuano": [""]}
    },
    "ISA ENERGIA BRASIL": {
        "Projeto vinculado": ["Replan", "Fernão Dias", "Piraque", "Itatiaia", "Serra Dourada"],
        "Trecho": {
            "Replan": [""], "Fernão Dias": [""],
            "Piraque": [
                "LT 500 kv JAIBA-JANAUBA6", "LT 500 Kv JANAUBA6-JANAUBA3",
                "LT 500 kV JANAUBA6 -CAPELINHA 3", "LT 500 kV CAPELINHA 3 -GOVERNADOR VALADARES",
                "LT 500 kV João Neiv 2 - Viana 2 C1", "LT 345 kV Viana 2 - Viana, C3",
                "SE Janaúba 6", "SE Capelinha 3", "SE Jaíba", "SE Janaúba 3",
                "SE Governador Valadares 6", "SE João Neiva 2", "SE Viana 2", "SE Viana"
            ],
            "Itatiaia": ["Governador Valadares 6 - Leopoldina 2", "Leopoldina 2 - Terminal Rio"],
            "Serra Dourada": [
                "SE BURITIRAMA", "SE BARRA II", "SE CORRENTINA", "SE ARINOS", "SE CAMPO FORMOSO",
                "SE JUAZEIRO", "SE BOM JESUS DA LAPA", "SE RIO DAS ÉGUAS",
                "LT 500 kV Buritirama -Barra II C1, CS", "LT 500 kV Barra II -Correntina C1, CS",
                "LT 500 kV Correntina -Arinos 2 C1, CS", "LT 500 kV Juazeiro III -Campo Formoso II",
                "LT 500kV Campo Formoso II Barra II C1 C2", "LT 500kV Bom Jesus da Lapa- Rio das Égua"
            ]
        }
    },
    # ... (outras empresas com mesma estrutura)
}

# ============================================================
# INICIALIZAÇÃO DO ESTADO
# ============================================================
def inicializar_estado():
    """Inicializa todas as variáveis de estado"""
    if 'tela_atual' not in st.session_state:
        st.session_state.tela_atual = "inicio"
    
    if 'dados_iniciais' not in st.session_state:
        st.session_state.dados_iniciais = {}
    
    if 'dados_coletados' not in st.session_state:
        st.session_state.dados_coletados = []
    
    if 'dados_cobrancas_nao_validados' not in st.session_state:
        st.session_state.dados_cobrancas_nao_validados = []
    
    # Estados do processo
    if 'is_dados_validado' not in st.session_state:
        st.session_state.is_dados_validado = False
    if 'is_revisao_concluida' not in st.session_state:
        st.session_state.is_revisao_concluida = False
    if 'is_detalhamento_iniciado' not in st.session_state:
        st.session_state.is_detalhamento_iniciado = False
    if 'is_detalhamento_validado' not in st.session_state:
        st.session_state.is_detalhamento_validado = False
    if 'is_finalizado' not in st.session_state:
        st.session_state.is_finalizado = False

# ============================================================
# TELAS DO SISTEMA
# ============================================================
def mostrar_tela_inicio():
    """Tela inicial do sistema"""
    st.title("Sistema de Medições Jurídicas")
    st.subheader("Módulo de Cadastro Inicial")
    st.markdown("---")
    
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        if st.button("🚀 Iniciar Lançamento", use_container_width=True):
            st.session_state.tela_atual = "dados"
            st.rerun()

def mostrar_tela_dados():
    """Tela de dados iniciais"""
    st.title("Dados Iniciais - Faturamento")
    st.markdown("---")
    
    with st.form("dados_iniciais"):
        col1, col2 = st.columns(2)
        
        with col1:
            cnpj = st.text_input("CNPJ do fornecedor*", 
                               placeholder="00.000.000/0000-00",
                               help="CNPJ completo do fornecedor")
            
            empresa = st.selectbox("Empresa contratante*", 
                                 options=[""] + sorted(EMPRESAS),
                                 help="Selecione a empresa contratante")
            
            advogado = st.selectbox("Advogado (a) responsável*",
                                  options=[""] + sorted(ADVOGADOS),
                                  help="Selecione o advogado responsável")
            
            tipo_doc = st.selectbox("Tipo de documento de cobrança*",
                                  options=["", "Nota Fiscal", "Nota de Débito", "RPA (Recibo de Pagamento Autônomo)"],
                                  help="Tipo do documento de cobrança")
        
        with col2:
            data_prevista = st.date_input("Data prevista de emissão*",
                                        value=date.today(),
                                        help="Data prevista para emissão do documento")
            
            existe_contrato = st.selectbox("Existe contrato vinculado?",
                                         options=["Não", "Sim"],
                                         help="Indique se existe contrato vinculado")
            
            n_contrato = ""
            if existe_contrato == "Sim":
                n_contrato = st.text_input("Nº do contrato*",
                                         placeholder="Formato: XX99999999",
                                         max_chars=10,
                                         help="Número do contrato (10 caracteres)")
            
            existe_pedido = st.selectbox("Existe pedido vinculado?",
                                       options=["Não", "Sim"],
                                       help="Indique se existe pedido vinculado")
            
            n_pedido = ""
            if existe_pedido == "Sim":
                n_pedido = st.text_input("Nº do pedido*",
                                       placeholder="Apenas números (10 dígitos)",
                                       max_chars=10,
                                       help="Número do pedido (10 dígitos)")
        
        n_medicao = st.text_input("Nº medição (doc.interno fornecedor)*",
                                placeholder="Número interno do fornecedor",
                                help="Número da medição no sistema do fornecedor")
        
        breve_desc = st.text_area("Breve descrição da fatura*",
                                placeholder="Descrição resumida da fatura",
                                help="Descrição breve da fatura")
        
        # Validação e submissão
        if st.form_submit_button("✅ Validar Dados e Prosseguir para Revisão", use_container_width=True):
            erros = []
            
            # Validações
            if not cnpj or len(filtrar_numeros(cnpj)) != 14:
                erros.append("CNPJ inválido ou incompleto")
            elif not cnpj_valido(filtrar_numeros(cnpj)):
                erros.append("CNPJ inválido")
            
            if not empresa:
                erros.append("Empresa contratante é obrigatória")
            if not advogado:
                erros.append("Advogado responsável é obrigatório")
            if not tipo_doc:
                erros.append("Tipo de documento é obrigatório")
            if not data_prevista:
                erros.append("Data prevista é obrigatória")
            if not n_medicao:
                erros.append("Nº medição é obrigatório")
            if not breve_desc:
                erros.append("Breve descrição é obrigatória")
            if existe_contrato == "Sim" and (not n_contrato or len(n_contrato) != 10):
                erros.append("Nº do contrato deve ter 10 caracteres quando 'Sim' for selecionado")
            if existe_pedido == "Sim" and (not n_pedido or len(n_pedido) != 10):
                erros.append("Nº do pedido deve ter 10 dígitos quando 'Sim' for selecionado")
            
            if erros:
                for erro in erros:
                    st.error(erro)
            else:
                # Salvar dados
                st.session_state.dados_iniciais = {
                    "cnpj": cnpj,
                    "empresa": empresa,
                    "advogado": advogado,
                    "tipo_doc": tipo_doc,
                    "data_prevista": data_prevista,
                    "existe_contrato": existe_contrato,
                    "n_contrato": n_contrato,
                    "existe_pedido": existe_pedido,
                    "n_pedido": n_pedido,
                    "n_medicao": n_medicao,
                    "breve_desc": breve_desc
                }
                st.session_state.is_dados_validado = True
                st.session_state.tela_atual = "revisao"
                st.success("Dados validados com sucesso!")
                st.rerun()

def mostrar_tela_revisao():
    """Tela de revisão dos dados iniciais"""
    if not st.session_state.is_dados_validado:
        st.error("Você deve preencher e validar os Dados Iniciais antes de avançar para a Revisão.")
        st.session_state.tela_atual = "dados"
        st.rerun()
    
    st.title("Revisão dos Dados Iniciais")
    st.markdown("---")
    
    dados = st.session_state.dados_iniciais
    
    st.subheader("Resumo dos Dados Iniciais")
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.info(f"**CNPJ Fornecedor:** {dados.get('cnpj', 'N/A')}")
        st.info(f"**Empresa Contratante:** {dados.get('empresa', 'N/A')}")
        st.info(f"**Advogado Responsável:** {dados.get('advogado', 'N/A')}")
        st.info(f"**Tipo de Documento:** {dados.get('tipo_doc', 'N/A')}")
        st.info(f"**Data de Emissão:** {dados.get('data_prevista', 'N/A')}")
    
    with col2:
        contrato_texto = f"Sim (Nº {dados.get('n_contrato', 'N/A')})" if dados.get('existe_contrato') == "Sim" else "Não"
        st.info(f"**Contrato Vinculado:** {contrato_texto}")
        
        pedido_texto = f"Sim (Nº {dados.get('n_pedido', 'N/A')})" if dados.get('existe_pedido') == "Sim" else "Não"
        st.info(f"**Pedido Vinculado:** {pedido_texto}")
        
        st.info(f"**Nº medição:** {dados.get('n_medicao', 'N/A')}")
        st.info(f"**Breve descrição:** {dados.get('breve_desc', 'N/A')}")
    
    st.markdown("---")
    
    col1, col2 = st.columns(2)
    with col1:
        if st.button("← Voltar e Corrigir", use_container_width=True):
            st.session_state.tela_atual = "dados"
            st.rerun()
    
    with col2:
        if st.button("Iniciar Detalhamento →", use_container_width=True):
            st.session_state.is_revisao_concluida = True
            st.session_state.tela_atual = "detalhamento"
            st.rerun()

def mostrar_tela_detalhamento():
    """Tela de detalhamento das cobranças"""
    if not st.session_state.is_revisao_concluida:
        st.error("Você deve confirmar a Revisão dos Dados Iniciais antes de avançar para o Detalhamento.")
        st.session_state.tela_atual = "revisao"
        st.rerun()
    
    st.title("Detalhamento das Cobranças")
    st.markdown("---")
    
    # Gerenciar cobranças
    if 'cobrancas' not in st.session_state:
        st.session_state.cobrancas = [{}]
    
    # Adicionar/remover cobranças
    col1, col2 = st.columns([3, 1])
    with col2:
        if st.button("➕ Adicionar Cobrança"):
            st.session_state.cobrancas.append({})
            st.rerun()
    
    # Formulário para cada cobrança
    for i, cobranca in enumerate(st.session_state.cobrancas):
        with st.expander(f"Cobrança {i+1}", expanded=True):
            mostrar_formulario_cobranca(i)
    
    st.markdown("---")
    
    # Botões de ação
    col1, col2, col3 = st.columns([1, 1, 1])
    with col1:
        if st.button("← Voltar", use_container_width=True):
            st.session_state.tela_atual = "revisao"
            st.rerun()
    
    with col3:
        if st.button("Validar e Revisar →", use_container_width=True):
            if validar_detalhamento():
                st.session_state.is_detalhamento_validado = True
                st.session_state.tela_atual = "revisao_detalhada"
                st.rerun()

def mostrar_formulario_cobranca(index):
    """Mostra formulário para uma cobrança específica"""
    cobranca_key = f"cobranca_{index}"
    
    col1, col2 = st.columns(2)
    
    with col1:
        possui_espaider = st.selectbox(
            "Possui Nº Espaider?",
            options=["Não", "Sim"],
            key=f"possui_espaider_{index}"
        )
        
        n_espaider = ""
        if possui_espaider == "Sim":
            n_espaider = st.text_input(
                "Nº Espaider*",
                key=f"n_espaider_{index}"
            )
        
        materia = st.selectbox(
            "Matéria*",
            options=[""] + MATERIAS_JURIDICAS,
            key=f"materia_{index}"
        )
    
    with col2:
        possui_projeto = st.selectbox(
            "Possui projeto vinculado?",
            options=["Não", "Sim"],
            key=f"possui_projeto_{index}"
        )
        
        projeto_vinculado = ""
        trecho = ""
        if possui_projeto == "Sim":
            empresa = st.session_state.dados_iniciais.get('empresa', '')
            projetos = PROJETOS_POR_EMPRESA.get(empresa, {}).get("Projeto vinculado", [""])
            projeto_vinculado = st.selectbox(
                "Projeto vinculado*",
                options=projetos,
                key=f"projeto_{index}"
            )
            
            # Trechos baseados no projeto selecionado
            trechos = PROJETOS_POR_EMPRESA.get(empresa, {}).get("Trecho", {}).get(projeto_vinculado, [""])
            trecho = st.selectbox(
                "Trecho",
                options=trechos,
                key=f"trecho_{index}"
            )
    
    # Cobrança principal
    st.subheader("Cobrança Principal")
    col1, col2, col3 = st.columns([2, 1, 2])
    
    with col1:
        tipo_1 = st.selectbox(
            "Tipo de cobrança*",
            options=[""] + TIPOS_COBRANCA,
            key=f"tipo_1_{index}"
        )
    
    with col2:
        valor_1 = st.text_input(
            "Valor (R$)*",
            value="0,00",
            key=f"valor_1_{index}"
        )
    
    with col3:
        # Texto breve automático
        texto_breve_1 = "Aguardando seleção..."
        if tipo_1 and materia:
            sigla_tipo = MAPA_TIPO_COBRANCA.get(tipo_1, "TIPO?")
            sigla_materia = MAPA_MATERIA.get(materia, "MAT?")
            texto_breve_1 = f"{sigla_tipo}_{sigla_materia}"
        
        st.text_input(
            "Texto Breve Código Serviço",
            value=texto_breve_1,
            disabled=True,
            key=f"texto_1_{index}"
        )
    
    # Cobrança secundária
    st.subheader("Cobrança Secundária")
    mais_cobrancas = st.selectbox(
        "Existe mais de uma cobrança para este Nº Espaider?",
        options=["Não", "Sim"],
        key=f"mais_cobrancas_{index}"
    )
    
    if mais_cobrancas == "Sim":
        col1, col2, col3 = st.columns([2, 1, 2])
        
        with col1:
            tipo_2 = st.selectbox(
                "Tipo de cobrança*",
                options=[""] + TIPOS_COBRANCA,
                key=f"tipo_2_{index}"
            )
        
        with col2:
            valor_2 = st.text_input(
                "Valor (R$)*",
                value="0,00",
                key=f"valor_2_{index}"
            )
        
        with col3:
            # Texto breve automático
            texto_breve_2 = "Aguardando seleção..."
            if tipo_2 and materia:
                sigla_tipo = MAPA_TIPO_COBRANCA.get(tipo_2, "TIPO?")
                sigla_materia = MAPA_MATERIA.get(materia, "MAT?")
                texto_breve_2 = f"{sigla_tipo}_{sigla_materia}"
            
            st.text_input(
                "Texto Breve Código Serviço",
                value=texto_breve_2,
                disabled=True,
                key=f"texto_2_{index}"
            )
    
    # Botão para remover cobrança (exceto a primeira)
    if index > 0:
        if st.button(f"🗑️ Remover Cobrança {index+1}", key=f"remover_{index}"):
            st.session_state.cobrancas.pop(index)
            st.rerun()

def validar_detalhamento():
    """Valida todos os detalhamentos das cobranças"""
    erros = []
    
    for i in range(len(st.session_state.cobrancas)):
        # Validar campos obrigatórios básicos
        possui_espaider = st.session_state.get(f"possui_espaider_{i}", "Não")
        if possui_espaider == "Sim" and not st.session_state.get(f"n_espaider_{i}", "").strip():
            erros.append(f"Cobrança {i+1}: Nº Espaider é obrigatório quando 'Sim' é selecionado")
        
        if not st.session_state.get(f"materia_{i}", ""):
            erros.append(f"Cobrança {i+1}: Matéria é obrigatória")
        
        # Validar cobrança principal
        if not st.session_state.get(f"tipo_1_{i}", ""):
            erros.append(f"Cobrança {i+1}: Tipo de cobrança principal é obrigatório")
        
        valor_1 = st.session_state.get(f"valor_1_{i}", "0,00")
        valor_limpo_1 = re.sub(r'[^0-9]', '', valor_1)
        if not valor_limpo_1 or int(valor_limpo_1) == 0:
            erros.append(f"Cobrança {i+1}: Valor da cobrança principal deve ser maior que zero")
        
        # Validar cobrança secundária se existir
        if st.session_state.get(f"mais_cobrancas_{i}", "Não") == "Sim":
            if not st.session_state.get(f"tipo_2_{i}", ""):
                erros.append(f"Cobrança {i+1}: Tipo de cobrança secundária é obrigatório quando 'Sim' é selecionado")
            
            valor_2 = st.session_state.get(f"valor_2_{i}", "0,00")
            valor_limpo_2 = re.sub(r'[^0-9]', '', valor_2)
            if not valor_limpo_2 or int(valor_limpo_2) == 0:
                erros.append(f"Cobrança {i+1}: Valor da cobrança secundária deve ser maior que zero")
    
    if erros:
        for erro in erros:
            st.error(erro)
        return False
    
    # Coletar dados validados
    dados_coletados = []
    for i in range(len(st.session_state.cobrancas)):
        dados_cobranca = {
            'num_cobranca': i + 1,
            "Possui Nº Espaider?": st.session_state.get(f"possui_espaider_{i}", "Não"),
            "Nº Espaider": st.session_state.get(f"n_espaider_{i}", ""),
            "Possui projeto vinculado?": st.session_state.get(f"possui_projeto_{i}", "Não"),
            "Projeto vinculado": st.session_state.get(f"projeto_{i}", ""),
            "Trecho": st.session_state.get(f"trecho_{i}", ""),
            "Matéria": st.session_state.get(f"materia_{i}", ""),
            'bloco_1': {
                "tipo": st.session_state.get(f"tipo_1_{i}", ""),
                "materia": st.session_state.get(f"materia_{i}", ""),
                "valor": st.session_state.get(f"valor_1_{i}", "0,00"),
                "texto_breve": st.session_state.get(f"texto_1_{i}", "Aguardando seleção...")
            }
        }
        
        if st.session_state.get(f"mais_cobrancas_{i}", "Não") == "Sim":
            dados_cobranca['bloco_2'] = {
                "tipo": st.session_state.get(f"tipo_2_{i}", ""),
                "materia": st.session_state.get(f"materia_{i}", ""),
                "valor": st.session_state.get(f"valor_2_{i}", "0,00"),
                "texto_breve": st.session_state.get(f"texto_2_{i}", "Aguardando seleção...")
            }
        
        dados_coletados.append(dados_cobranca)
    
    st.session_state.dados_coletados = dados_coletados
    st.session_state.dados_cobrancas_nao_validados = dados_coletados.copy()
    
    return True

def mostrar_tela_revisao_detalhada():
    """Tela de revisão detalhada"""
    if not st.session_state.is_detalhamento_validado:
        st.error("Ocorreu um erro. A validação do detalhamento não foi concluída.")
        st.session_state.tela_atual = "detalhamento"
        st.rerun()
    
    st.title("Revisão do Detalhamento")
    st.markdown("---")
    
    # Resumo dos dados iniciais
    st.subheader("Resumo dos Dados Iniciais")
    dados = st.session_state.dados_iniciais
    
    col1, col2 = st.columns(2)
    with col1:
        st.info(f"**CNPJ:** {dados.get('cnpj', 'N/A')}")
        st.info(f"**Empresa:** {dados.get('empresa', 'N/A')}")
        st.info(f"**Advogado(a):** {dados.get('advogado', 'N/A')}")
        st.info(f"**Documento:** {dados.get('tipo_doc', 'N/A')}")
    
    with col2:
        st.info(f"**Data Emissão:** {dados.get('data_prevista', 'N/A')}")
        st.info(f"**Nº medição:** {dados.get('n_medicao', 'N/A')}")
        contrato_texto = f"Sim (Nº {dados.get('n_contrato', 'N/A')})" if dados.get('existe_contrato') == "Sim" else "Não"
        st.info(f"**Contrato:** {contrato_texto}")
    
    st.markdown("---")
    
    # Resumo das cobranças
    st.subheader("Resumo das Cobranças")
    
    for cobranca in st.session_state.dados_coletados:
        with st.expander(f"Cobrança {cobranca['num_cobranca']}", expanded=True):
            col1, col2 = st.columns(2)
            
            with col1:
                st.write(f"**Nº Espaider:** {cobranca.get('Nº Espaider', 'Não informado')}")
                st.write(f"**Projeto vinculado:** {cobranca.get('Projeto vinculado', 'Não informado')}")
                st.write(f"**Trecho:** {cobranca.get('Trecho', 'Não informado')}")
                st.write(f"**Matéria:** {cobranca.get('Matéria', 'N/A')}")
            
            with col2:
                bloco_1 = cobranca.get('bloco_1', {})
                st.write(f"**Cobrança Principal:**")
                st.write(f"- Tipo: {bloco_1.get('tipo', 'N/A')}")
                st.write(f"- Valor: R$ {bloco_1.get('valor', '0,00')}")
                st.write(f"- Texto Breve: {bloco_1.get('texto_breve', 'N/A')}")
                
                if 'bloco_2' in cobranca:
                    bloco_2 = cobranca['bloco_2']
                    st.write(f"**Cobrança Secundária:**")
                    st.write(f"- Tipo: {bloco_2.get('tipo', 'N/A')}")
                    st.write(f"- Valor: R$ {bloco_2.get('valor', '0,00')}")
                    st.write(f"- Texto Breve: {bloco_2.get('texto_breve', 'N/A')}")
    
    st.markdown("---")
    
    # Botões de ação
    col1, col2 = st.columns(2)
    
    with col1:
        if st.button("← Voltar e Corrigir Detalhes", use_container_width=True):
            st.session_state.tela_atual = "detalhamento"
            st.rerun()
    
    with col2:
        if st.button("🎯 FINALIZAR E GERAR EXCEL", use_container_width=True):
            gerar_excel()
            st.session_state.is_finalizado = True
            st.success("Processo finalizado com sucesso! Arquivo Excel gerado.")

# ============================================================
# FUNÇÕES PARA GERAR EXCEL
# ============================================================
def gerar_excel():
    """Gera o arquivo Excel final"""
    try:
        # Criar workbook
        workbook = openpyxl.Workbook()
        
        # Planilha principal
        sheet_principal = workbook.active
        sheet_principal.title = "Medições"
        formatar_planilha_principal(sheet_principal)
        
        # Planilha BD
        sheet_bd = workbook.create_sheet(title="BD")
        formatar_planilha_bd(sheet_bd)
        
        # Proteção
        sheet_principal.protection.password = 'SINAPSE4'
        sheet_principal.protection.sheet = True
        sheet_bd.protection.password = 'SINAPSE4'
        sheet_bd.protection.sheet = True
        sheet_bd.sheet_state = 'hidden'
        
        workbook.security = WorkbookProtection(
            workbookPassword='SINAPSE4',
            lockStructure=True
        )
        
        # Salvar para download
        buffer = io.BytesIO()
        workbook.save(buffer)
        buffer.seek(0)
        
        # Botão de download
        st.download_button(
            label="📥 Baixar Arquivo Excel",
            data=buffer,
            file_name=f"Medicoes_Juridicas_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
        
    except Exception as e:
        st.error(f"Erro ao gerar Excel: {e}")

def formatar_planilha_principal(sheet):
    """Formata a planilha principal do Excel"""
    # Configurações de estilo
    font_header_azul = Font(name='Segoe UI', size=11, bold=True, color="FFFFFF")
    fill_header_azul = PatternFill(start_color="002060", end_color="002060", fill_type="solid")
    font_bold = Font(name='Segoe UI', size=10, bold=True)
    font_normal = Font(name='Segoe UI', size=10)
    
    alignment_left = Alignment(horizontal='left', vertical='center', wrap_text=True)
    alignment_center = Alignment(horizontal='center', vertical='center')
    thin_border = Border(
        left=Side(style='thin'), right=Side(style='thin'),
        top=Side(style='thin'), bottom=Side(style='thin')
    )
    
    # Cabeçalho
    sheet['D2'] = f"Medição Jurídica Nº {st.session_state.dados_iniciais.get('n_medicao', 'N/A')}"
    sheet['D2'].font = Font(name='Segoe UI', size=16, bold=True)
    sheet['D2'].alignment = alignment_center
    sheet.merge_cells('D2:G2')
    
    # Dados iniciais
    sheet['B5'] = "Dados Iniciais - Faturamento"
    sheet['B5'].font = font_header_azul
    sheet['B5'].fill = fill_header_azul
    sheet['B5'].alignment = alignment_center
    sheet.merge_cells('B5:G5')
    
    # ... (continua com a formatação completa da planilha)

def formatar_planilha_bd(sheet):
    """Formata a planilha BD do Excel"""
    # Implementação similar à função original
    pass

# ============================================================
# BARRA LATERAL
# ============================================================
def mostrar_sidebar():
    """Mostra a barra lateral com navegação e status"""
    with st.sidebar:
        st.title("⚖️ Medições Jurídicas")
        st.markdown("---")
        
        # Navegação
        st.subheader("Navegação")
        
        if st.button("🏠 Início", use_container_width=True):
            st.session_state.tela_atual = "inicio"
            st.rerun()
        
        if st.button("📊 Dados Iniciais", use_container_width=True):
            st.session_state.tela_atual = "dados"
            st.rerun()
        
        if st.button("👁️ Revisão", use_container_width=True, 
                    disabled=not st.session_state.is_dados_validado):
            if st.session_state.is_dados_validado:
                st.session_state.tela_atual = "revisao"
                st.rerun()
        
        if st.button("📋 Detalhamento", use_container_width=True,
                    disabled=not st.session_state.is_revisao_concluida):
            if st.session_state.is_revisao_concluida:
                st.session_state.tela_atual = "detalhamento"
                st.rerun()
        
        if st.button("🔍 Revisão Detalhada", use_container_width=True,
                    disabled=not st.session_state.is_detalhamento_validado):
            if st.session_state.is_detalhamento_validado:
                st.session_state.tela_atual = "revisao_detalhada"
                st.rerun()
        
        st.markdown("---")
        
        # Status do processo
        st.subheader("Status do Processo")
        
        status_icons = {
            "done": "✅",
            "active": "🟢", 
            "pending": "⚪",
            "disabled": "⚫"
        }
        
        # Etapa 1
        if st.session_state.is_dados_validado:
            st.write(f"{status_icons['done']} 1. Dados Iniciais")
        else:
            st.write(f"{status_icons['active']} 1. Dados Iniciais")
        
        # Etapa 2
        if st.session_state.is_revisao_concluida:
            st.write(f"{status_icons['done']} 2. Revisão")
        elif st.session_state.is_dados_validado:
            st.write(f"{status_icons['active']} 2. Revisão")
        else:
            st.write(f"{status_icons['pending']} 2. Revisão")
        
        # Etapa 3
        if st.session_state.is_detalhamento_validado:
            st.write(f"{status_icons['done']} 3. Detalhamento")
        elif st.session_state.is_revisao_concluida:
            st.write(f"{status_icons['active']} 3. Detalhamento")
        else:
            st.write(f"{status_icons['pending']} 3. Detalhamento")
        
        # Etapa 4
        if st.session_state.is_finalizado:
            st.write(f"{status_icons['done']} 4. Revisão Detalhada")
        elif st.session_state.is_detalhamento_validado:
            st.write(f"{status_icons['active']} 4. Revisão Detalhada")
        else:
            st.write(f"{status_icons['pending']} 4. Revisão Detalhada")
        
        # Etapa 5
        if st.session_state.is_finalizado:
            st.write(f"{status_icons['done']} 5. Geração do Arquivo")
        else:
            st.write(f"{status_icons['pending']} 5. Geração do Arquivo")
        
        st.markdown("---")
        st.markdown("© Equipe de Desenvolvimento")

# ============================================================
# APLICAÇÃO PRINCIPAL
# ============================================================
def main():
    """Função principal da aplicação"""
    
    # Inicializar estado
    inicializar_estado()
    
    # Mostrar sidebar
    mostrar_sidebar()
    
    # Mostrar tela atual
    if st.session_state.tela_atual == "inicio":
        mostrar_tela_inicio()
    elif st.session_state.tela_atual == "dados":
        mostrar_tela_dados()
    elif st.session_state.tela_atual == "revisao":
        mostrar_tela_revisao()
    elif st.session_state.tela_atual == "detalhamento":
        mostrar_tela_detalhamento()
    elif st.session_state.tela_atual == "revisao_detalhada":
        mostrar_tela_revisao_detalhada()

if __name__ == "__main__":
    main()
