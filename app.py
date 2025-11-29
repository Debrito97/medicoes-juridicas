# app.py
import streamlit as st
from datetime import datetime, date
import re
from io import BytesIO
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.drawing.image import Image as OpenPyXLImage
from PIL import Image
import base64
import os

# ---------------------------
# Helpers (adaptados do main.py)
# ---------------------------
def filtrar_numeros(texto: str) -> str:
    return "".join(ch for ch in (texto or "") if ch.isdigit())

def cnpj_valido(cnpj: str) -> bool:
    c = filtrar_numeros(cnpj)
    if len(c) != 14 or c == c[0] * 14:
        return False
    pesos1 = [5,4,3,2,9,8,7,6,5,4,3,2]
    pesos2 = [6,5,4,3,2,9,8,7,6,5,4,3,2]
    soma = sum(int(c[i]) * pesos1[i] for i in range(12))
    resto = soma % 11
    digito1 = 0 if resto < 2 else 11 - resto
    soma = sum(int(c[i]) * pesos2[i] for i in range(13))
    resto = soma % 11
    digito2 = 0 if resto < 2 else 11 - resto
    return c[12] == str(digito1) and c[13] == str(digito2)

def formatar_valor_para_float(valor_str):
    if not valor_str:
        return 0.0
    v = re.sub(r'[^0-9,\.]', '', valor_str).replace(',', '.')
    try:
        return float(v)
    except:
        return 0.0

def criar_excel(dados_iniciais, cobrancas):
    """Gera um arquivo excel em memória (openpyxl) similar ao original."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Medições"
    # header style
    font_header_azul = Font(name='Segoe UI', size=11, bold=True, color="FFFFFF")
    fill_header_azul = PatternFill(start_color="002060", end_color="002060", fill_type="solid")
    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'),
                         top=Side(style='thin'), bottom=Side(style='thin'))

    # logo (se existir arquivo isa_energia_logo.png no repo)
    try:
        logo_path = "isa_energia_logo.png"
        if os.path.exists(logo_path):
            img = OpenPyXLImage(logo_path)
            img.width = 120
            img.height = 40
            ws.add_image(img, "B2")
        else:
            ws["B2"] = "ISA Energia"
    except Exception as e:
        ws["B2"] = "ISA Energia (logo falhou)"

    # escrever dados iniciais no topo
    ws["D2"] = "Nº medição jurídica"
    ws["D3"] = dados_iniciais.get("n_medicao", "")
    ws["B5"] = "CNPJ"
    ws["C5"] = dados_iniciais.get("cnpj", "")
    ws["B6"] = "Empresa"
    ws["C6"] = dados_iniciais.get("empresa", "")
    ws["B7"] = "Advogado"
    ws["C7"] = dados_iniciais.get("advogado", "")
    ws["B8"] = "Tipo Documento"
    ws["C8"] = dados_iniciais.get("tipo_doc", "")
    ws["B9"] = "Data Emissão"
    dt = dados_iniciais.get("data_prevista")
    ws["C9"] = dt.strftime("%d/%m/%Y") if isinstance(dt, (date, datetime)) else str(dt or "")

    # headers para cobranças
    start_row = 12
    headers = ["Idx", "Nº Espaider", "Projeto", "Trecho", "Matéria", "Tipo Cobrança", "Valor", "Resumo"]
    for col_idx, h in enumerate(headers, start=1):
        cell = ws.cell(row=start_row, column=col_idx, value=h)
        cell.font = font_header_azul
        cell.fill = fill_header_azul
        cell.alignment = Alignment(horizontal='center')
        cell.border = thin_border

    # popular cobranças
    for i, c in enumerate(cobrancas, start=1):
        r = start_row + i
        ws.cell(row=r, column=1, value=i)
        ws.cell(row=r, column=2, value=c.get("n_esp", ""))
        ws.cell(row=r, column=3, value=c.get("projeto", ""))
        ws.cell(row=r, column=4, value=c.get("trecho", ""))
        ws.cell(row=r, column=5, value=c.get("materia", ""))
        ws.cell(row=r, column=6, value=c.get("tipo_cobranca", ""))
        val = formatar_valor_para_float(c.get("valor", "0"))
        ws.cell(row=r, column=7, value=val)
        ws.cell(row=r, column=8, value=c.get("resumo", ""))

        for col in range(1, 9):
            ws.cell(row=r, column=col).border = thin_border

    # ajustar larguras básicas
    for col in range(1, 9):
        ws.column_dimensions[openpyxl.utils.get_column_letter(col)].width = 18

    # salvar em BytesIO
    bio = BytesIO()
    wb.save(bio)
    bio.seek(0)
    return bio

# ---------------------------
# Inicialização sessão
# ---------------------------
if 'step' not in st.session_state:
    st.session_state.step = "inicio"
if 'dados_iniciais' not in st.session_state:
    st.session_state.dados_iniciais = {}
if 'cobrancas' not in st.session_state:
    st.session_state.cobrancas = []

# ---------------------------
# Sidebar com o fluxo (mantendo etapas)
# ---------------------------
st.sidebar.title("Medições Jurídicas")
menu = ["Início", "Dados Iniciais", "Revisão", "Detalhamento", "Revisão Detalhada", "Geração do Arquivo"]
choice = st.sidebar.radio("Etapas", menu, index=menu.index({"Início":"Início","Dados Iniciais":"Dados Iniciais"}[st.session_state.get('step',"Início")] if False else 0))

# map radio to step
if choice == "Início":
    st.session_state.step = "inicio"
elif choice == "Dados Iniciais":
    st.session_state.step = "dados"
elif choice == "Revisão":
    st.session_state.step = "revisao"
elif choice == "Detalhamento":
    st.session_state.step = "detalhamento"
elif choice == "Revisão Detalhada":
    st.session_state.step = "revisao_detalhada"
elif choice == "Geração do Arquivo":
    st.session_state.step = "geracao"

# ---------------------------
# TELA: Início
# ---------------------------
if st.session_state.step == "inicio":
    st.title("Sistema de Medições Jurídicas")
    st.subheader("Módulo de Cadastro Inicial")
    st.markdown("Clique em **Dados Iniciais** no menu lateral para começar.")
    if st.button("Iniciar Lançamento"):
        st.session_state.step = "dados"
        st.experimental_rerun()

# ---------------------------
# TELA: Dados Iniciais
# ---------------------------
if st.session_state.step == "dados":
    st.header("Dados Iniciais - Faturamento")
    with st.form("form_dados"):
        cnpj = st.text_input("CNPJ do fornecedor", value=st.session_state.dados_iniciais.get("cnpj",""))
        empresa = st.text_input("Empresa contratante", value=st.session_state.dados_iniciais.get("empresa",""))
        advogado = st.text_input("Advogado (a) responsável", value=st.session_state.dados_iniciais.get("advogado",""))
        tipo_doc = st.selectbox("Tipo de documento de cobrança", ["","Consultoria","Nota","Fatura","Outro"], index=0)
        data_prevista = st.date_input("Data prevista de emissão", value=st.session_state.dados_iniciais.get("data_prevista", date.today()))
        existe_contrato = st.selectbox("Existe contrato vinculado?", ["Não", "Sim"], index=0)
        n_contrato = st.text_input("Nº do contrato (se aplicável)", value=st.session_state.dados_iniciais.get("n_contrato",""))
        existe_pedido = st.selectbox("Existe pedido vinculado?", ["Não", "Sim"], index=0)
        n_pedido = st.text_input("Nº do pedido (se aplicável)", value=st.session_state.dados_iniciais.get("n_pedido",""))
        n_medicao = st.text_input("Nº medição (doc.interno fornecedor)", value=st.session_state.dados_iniciais.get("n_medicao",""))
        breve_desc = st.text_area("Breve descrição da fatura", value=st.session_state.dados_iniciais.get("breve_desc",""), height=80)

        submitted = st.form_submit_button("Validar Dados e Prosseguir para Revisão")
        if submitted:
            # validações
            erros = []
            if not cnpj or not cnpj_valido(cnpj):
                erros.append("CNPJ inválido ou vazio.")
            if existe_contrato == "Sim" and (not n_contrato or len(re.sub(r'[^0-9]','', n_contrato)) < 5):
                erros.append("Nº do contrato inválido (quando 'Sim' selecionado).")
            if existe_pedido == "Sim" and (not n_pedido or len(re.sub(r'[^0-9]','', n_pedido)) != 10):
                erros.append("Nº do pedido deve ter 10 dígitos (quando 'Sim' selecionado).")
            if not n_medicao:
                erros.append("Nº medição é obrigatório.")
            if not breve_desc:
                erros.append("Breve descrição é obrigatória.")

            if erros:
                st.error("Erros:\n- " + "\n- ".join(erros))
            else:
                st.success("Dados iniciais validados com sucesso.")
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
                st.session_state.step = "revisao"
                st.experimental_rerun()

# ---------------------------
# TELA: Revisão
# ---------------------------
if st.session_state.step == "revisao":
    st.header("Revisão dos Dados Iniciais")
    di = st.session_state.dados_iniciais
    if not di:
        st.warning("Nenhum dado inicial validado. Volte para 'Dados Iniciais'.")
    else:
        st.markdown("**Revise e confirme os dados antes de prosseguir.**")
        st.write("**CNPJ:**", di.get("cnpj"))
        st.write("**Empresa:**", di.get("empresa"))
        st.write("**Advogado:**", di.get("advogado"))
        st.write("**Tipo Documento:**", di.get("tipo_doc"))
        st.write("**Data de Emissão:**", di.get("data_prevista").strftime("%d/%m/%Y"))
        st.write("**Contrato:**", di.get("n_contrato") if di.get("existe_contrato")=="Sim" else "Não")
        st.write("**Pedido:**", di.get("n_pedido") if di.get("existe_pedido")=="Sim" else "Não")
        st.write("**Nº medição:**", di.get("n_medicao"))
        st.write("**Breve descrição:**", di.get("breve_desc"))

        col1, col2 = st.columns(2)
        with col1:
            if st.button("← Voltar e Corrigir"):
                st.session_state.step = "dados"
                st.experimental_rerun()
        with col2:
            if st.button("Iniciar Detalhamento →"):
                st.session_state.step = "detalhamento"
                st.experimental_rerun()

# ---------------------------
# TELA: Detalhamento (adicionar cobranças)
# ---------------------------
if st.session_state.step == "detalhamento":
    st.header("Detalhamento das Cobranças")
    st.markdown("Preencha as informações de cada cobrança. Você pode adicionar várias cobranças.")

    # formulário para adicionar cobrança
    with st.form("form_cobranca"):
        n_esp = st.text_input("Nº Espaider")
        possui_projeto = st.selectbox("Possui projeto vinculado?", ["Não","Sim"])
        projeto = st.text_input("Projeto vinculado (se aplicável)")
        trecho = st.text_input("Trecho")
        materia = st.selectbox("Matéria", ["","Trabalhista","Cível","Regulatório","Ambiental","Outro"])
        tipo_cobranca = st.selectbox("Tipo de Cobrança", ["Consultorias","Despesas","Êxito","Honorários","Parecer","Perícia","Publicações Legais"])
        valor = st.text_input("Valor (ex: 1.234,56)")
        resumo = st.text_area("Resumo / breve descrição", height=60)

        add_sub = st.form_submit_button("Adicionar Cobrança")
        if add_sub:
            st.session_state.cobrancas.append({
                "n_esp": n_esp,
                "projeto": projeto if possui_projeto=="Sim" else "",
                "trecho": trecho,
                "materia": materia,
                "tipo_cobranca": tipo_cobranca,
                "valor": valor,
                "resumo": resumo
            })
            st.success("Cobrança adicionada.")
            st.experimental_rerun()

    # listar cobranças atuais
    if st.session_state.cobrancas:
        st.markdown("### Cobranças adicionadas")
        for idx, c in enumerate(st.session_state.cobrancas, start=1):
            with st.expander(f"Cobrança #{idx} — {c.get('tipo_cobranca')} — {c.get('valor')}"):
                st.write(c)
                if st.button(f"Remover cobrança #{idx}"):
                    st.session_state.cobrancas.pop(idx-1)
                    st.experimental_rerun()

    col1, col2 = st.columns(2)
    with col1:
        if st.button("← Voltar"):
            st.session_state.step = "revisao"
            st.experimental_rerun()
    with col2:
        if st.button("Validar e Revisar →"):
            # validar existência de pelo menos 1 cobrança
            if not st.session_state.cobrancas:
                st.error("Adicione pelo menos uma cobrança antes de prosseguir.")
            else:
                st.session_state.step = "revisao_detalhada"
                st.experimental_rerun()

# ---------------------------
# TELA: Revisão Detalhada
# ---------------------------
if st.session_state.step == "revisao_detalhada":
    st.header("Revisão do Detalhamento")
    st.markdown("Revise todas as cobranças antes de gerar o arquivo.")
    if not st.session_state.cobrancas:
        st.warning("Nenhuma cobrança adicionada.")
    else:
        total = sum(formatar_valor_para_float(c.get("valor","0")) for c in st.session_state.cobrancas)
        st.write("Total (soma das cobranças):", f"{total:,.2f}")
        for idx, c in enumerate(st.session_state.cobrancas, start=1):
            st.write(f"**#{idx}** — {c.get('tipo_cobranca')} — {c.get('valor')}")
            st.write(c.get("resumo",""))

    col1, col2 = st.columns(2)
    with col1:
        if st.button("← Voltar e Corrigir"):
            st.session_state.step = "detalhamento"
            st.experimental_rerun()
    with col2:
        if st.button("Gerar Arquivo →"):
            st.session_state.step = "geracao"
            st.experimental_rerun()

# ---------------------------
# TELA: Geração do Arquivo
# ---------------------------
if st.session_state.step == "geracao":
    st.header("Geração do Arquivo")
    st.write("Revise e gere o Excel final.")

    if not st.session_state.dados_iniciais:
        st.error("Dados iniciais ausentes.")
    else:
        if st.button("Gerar e Baixar Excel"):
            bio = criar_excel(st.session_state.dados_iniciais, st.session_state.cobrancas)
            fn = f"Medições_{st.session_state.dados_iniciais.get('n_medicao','sem_num')}.xlsx"
            st.download_button("📥 Baixar Excel", data=bio, file_name=fn, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    if st.button("Finalizar e Voltar ao Início"):
        # reset simples
        st.session_state.step = "inicio"
        st.session_state.dados_iniciais = {}
        st.session_state.cobrancas = []
        st.experimental_rerun()
