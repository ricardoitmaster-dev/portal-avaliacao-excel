import streamlit as st
import pandas as pd
import random
import io
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from datetime import datetime, timedelta

# --- CONFIGURAÇÕES VIA SECRETS ---
try:
    EMAIL_PROFESSOR = st.secrets["EMAIL_PROFESSOR"]
    SENHA_APP_GOOGLE = st.secrets["SENHA_APP_GOOGLE"]
except:
    EMAIL_PROFESSOR = "ricardoitmaster@gmail.com"
    SENHA_APP_GOOGLE = "ugjhusmwnbmgzspv"

# ==========================================
# CUSTOMIZAÇÃO DE IDENTIDADE VISUAL SENAI
# ==========================================
CORE_SENAI = "#FF0000"
CORE_TEXTO = "#404040"

st.set_page_config(
    page_title="Portal de Avaliação Excel - SENAI", 
    page_icon="https://upload.wikimedia.org/wikipedia/commons/8/8c/SENAI_S%C3%A3o_Paulo_logo.png",
    layout="centered"
)

st.markdown(f"""
    <style>
        h1 {{ color: {CORE_SENAI} !important; font-weight: bold; }}
        h2, h3 {{ color: {CORE_TEXTO} !important; }}
        .stButton>button {{ color: white; background-color: {CORE_TEXTO}; border-radius: 5px; }}
        .stButton>button:hover {{ background-color: {CORE_SENAI}; color: white; }}
        .stDownloadButton>button {{ color: white !important; background-color: {CORE_SENAI} !important; font-weight: bold; }}
        hr {{ border-top: 2px solid {CORE_SENAI}; }}
    </style>
""", unsafe_allow_html=True)

# --- FUNÇÕES DE MOTOR E CORREÇÃO ---
def enviar_email(destinatario, assunto, corpo, arquivo_bytes=None, nome_arquivo=None):
    try:
        msg = MIMEMultipart()
        msg['From'] = EMAIL_PROFESSOR
        msg['To'] = destinatario
        msg['Subject'] = assunto
        msg.attach(MIMEText(corpo, 'plain'))
        if arquivo_bytes:
            part = MIMEBase('application', 'octet-stream')
            part.set_payload(arquivo_bytes)
            encoders.encode_base64(part)
            part.add_header('Content-Disposition', f"attachment; filename={nome_arquivo}")
            msg.attach(part)
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(EMAIL_PROFESSOR, SENHA_APP_GOOGLE)
        server.send_message(msg)
        server.quit()
        return True
    except Exception as e:
        st.error(f"Erro no envio de e-mail: {e}")
        return False

def gerar_prova_excel(nome_aluno, turma):
    itens = ["Notebook", "Mouse", "Teclado", "Monitor", "Impressora"]
    dados = []
    for i in range(1, 11): # 10 linhas para teste de correção
        qtd = random.randint(5, 50)
        preco = round(random.uniform(20, 100), 2)
        dados.append({
            "ID": i, "Produto": random.choice(itens),
            "Quantidade": qtd, "Preço Unitário": preco,
            "Venda Total": 0, "Status": ""
        })
    df = pd.DataFrame(dados)
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, sheet_name='Base_de_Dados', index=False)
        inst = [["AVALIAÇÃO EXCEL"], [f"ALUNO: {nome_aluno}"], [f"TURMA: {turma}"], [""],
                ["1. Coluna Venda Total: Quantidade * Preço Unitário"],
                ["2. Coluna Status: Se Venda Total >= 500 escrever 'META', senão 'REVISAR'"]]
        pd.DataFrame(inst).to_excel(writer, sheet_name='Instruções', index=False, header=False)
    return output.getvalue()

def corrigir_prova(arquivo_aluno):
    try:
        df = pd.read_excel(arquivo_aluno, sheet_name='Base_de_Dados')
        pontos_venda = 0
        pontos_status = 0
        total_linhas = len(df)
        for index, row in df.iterrows():
            if round(row['Venda Total'], 2) == round(row['Quantidade'] * row['Preço Unitário'], 2):
                pontos_venda += 1
            esperado = "META" if row['Venda Total'] >= 500 else "REVISAR"
            if str(row['Status']).strip().upper() == esperado:
                pontos_status += 1
        nota = round(((pontos_venda / total_linhas) * 5) + ((pontos_status / total_linhas) * 5), 1)
        feedback = f"Resultado: {pontos_venda}/{total_linhas} cálculos e {pontos_status}/{total_linhas} funções SE corretas."
        return nota, feedback
    except Exception:
        return 0, "Erro ao ler as colunas. Verifique se usou o arquivo original."

# --- INTERFACE ---
if 'etapa' not in st.session_state: st.session_state.etapa = 'login'

if st.session_state.etapa == 'login':
    col_logo, col_tit = st.columns([1, 4])
    with col_logo: st.image("https://upload.wikimedia.org/wikipedia/commons/8/8c/SENAI_S%C3%A3o_Paulo_logo.png", width=120)
    with col_tit:
        st.title("Portal de Avaliação Prática")
        st.subheader("Cursos de Tecnologia da Informação - Excel Completo")
    nome = st.text_input("Nome Completo")
    turma = st.text_input("Sua Turma")
    email = st.text_input("Seu E-mail")
    if st.button("Acessar Avaliação"):
        if nome and turma and email:
            st.session_state.excel_data = gerar_prova_excel(nome, turma)
            st.session_state.nome_arquivo = f"Prova_{nome.replace(' ','_')}.xlsx"
            st.session_state.aluno = {"nome": nome, "turma": turma, "email": email}
            st.session_state.etapa = 'prova'
            st.rerun()
else:
    col_logo, col_tit = st.columns([1, 4])
    with col_logo: st.image("https://upload.wikimedia.org/wikipedia/commons/8/8c/SENAI_S%C3%A3o_Paulo_logo.png", width=120)
    with col_tit: st.title("Área de Entrega e Correção")
    
    st.info(f"Aluno: {st.session_state.aluno['nome']} | Turma: {st.session_state.aluno['turma']}")
    st.download_button("📥 Baixar minha Prova", st.session_state.excel_data, st.session_state.nome_arquivo)
    st.divider()
    
    arquivo_upload = st.file_uploader("Anexe sua prova resolvida", type=['xlsx'])
    if st.button("🚀 Enviar e Corrigir Agora"):
        if arquivo_upload:
            with st.spinner('Corrigindo...'):
                nota, feedback = corrigir_prova(arquivo_upload)
                corpo = f"Aluno: {st.session_state.aluno['nome']}\nNota: {nota}\n{feedback}"
                enviar_email(EMAIL_PROFESSOR, f"NOTA {nota}: {st.session_state.aluno['nome']}", corpo, arquivo_upload.getvalue(), arquivo_upload.name)
                enviar_email(st.session_state.aluno['email'], "Seu Resultado - Prova Excel SENAI", corpo)
                st.success(f"Nota Final: {nota}")
                st.write(feedback)
                st.balloons()
        else:
            st.error("Anexe o arquivo.")
    if st.button("🚪 Sair"):
        st.session_state.clear()
        st.rerun()
