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

# --- IDENTIDADE VISUAL (MODO ESCURO - DARK MASTER) ---
CORE_SENAI = "#FF0000"
CORE_FUNDO = "#0E1117" 
CORE_TEXTO_BRANCO = "#FFFFFF"

st.set_page_config(
    page_title="Portal de Avaliação Excel - SENAI", 
    page_icon="https://upload.wikimedia.org/wikipedia/commons/8/8c/SENAI_S%C3%A3o_Paulo_logo.png",
    layout="centered"
)

# CSS para Fundo Escuro e Centralização
st.markdown(f"""
    <style>
        .stApp {{ background-color: {CORE_FUNDO} !important; }}
        h1 {{ 
            color: {CORE_SENAI} !important; 
            font-weight: bold; 
            text-align: center !important;
        }}
        .centered-subtitle {{
            text-align: center !important;
            color: {CORE_TEXTO_BRANCO} !important;
            font-size: 1.2rem;
            font-weight: 500;
            margin-bottom: 30px;
        }}
        label, p, span {{ color: {CORE_TEXTO_BRANCO} !important; }}
        .stButton>button {{ 
            color: white; 
            background-color: #262730; 
            border-radius: 8px; 
            width: 100%;
            border: 1px solid {CORE_SENAI};
        }}
        .stButton>button:hover {{ background-color: {CORE_SENAI}; }}
        .stDownloadButton>button {{ 
            color: white !important; 
            background-color: {CORE_SENAI} !important; 
            font-weight: bold; 
            width: 100%;
        }}
        .stTextInput input {{ background-color: #262730 !important; color: white !important; }}
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
    except:
        return False

def gerar_prova_excel(nome_aluno, turma):
    itens = ["Notebook", "Mouse", "Teclado", "Monitor", "Impressora", "Cabo HDMI", "SSD 480GB"]
    dados = []
    for i in range(1, 31):
        qtd = random.randint(5, 50)
        preco = round(random.uniform(20, 300), 2)
        dados.append({
            "ID": i, "Produto": random.choice(itens),
            "Quantidade": qtd, "Preço Unitário": preco,
            "Venda Total": 0, "Status": ""
        })
    df = pd.DataFrame(dados)
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, sheet_name='Base_de_Dados', index=False)
        inst = [
            ["INSTRUÇÕES TÉCNICAS (NÃO ALTERE OS NOMES DAS COLUNAS):"],
            ["1. Na coluna Venda Total: Use a função de MULTIPLICAÇÃO ou MULT (Quantidade * Preço)."],
            ["2. Na coluna Status: Use a função lógica SE: Se Venda Total >= 500 escrever 'META', senão 'REVISAR'."],
            ["3. Salve o arquivo e envie de volta no portal."]
        ]
        pd.DataFrame(inst).to_excel(writer, sheet_name='Instrucoes', index=False, header=False)
    return output.getvalue()

def corrigir_prova(arquivo_aluno):
    try:
        # Forçamos o motor openpyxl para evitar o erro de leitura
        df = pd.read_excel(arquivo_aluno, sheet_name='Base_de_Dados', engine='openpyxl')
        pv, ps, total = 0, 0, len(df)
        for index, row in df.iterrows():
            if round(float(row['Venda Total']), 2) == round(float(row['Quantidade'] * row['Preço Unitário']), 2):
                pv += 1
            esp = "META" if row['Venda Total'] >= 500 else "REVISAR"
            if str(row['Status']).strip().upper() == esp:
                ps += 1
        nota = round(((pv / total) * 5) + ((ps / total) * 5), 1)
        return nota, f"Cálculos: {pv}/{total} | Lógica SE: {ps}/{total}"
    except Exception as e:
        return 0, f"Erro Técnico: Certifique-se de preencher a aba 'Base_de_Dados' corretamente."

# --- INTERFACE ---
if 'etapa' not in st.session_state: st.session_state.etapa = 'login'

if st.session_state.etapa == 'login':
    st.image("https://upload.wikimedia.org/wikipedia/commons/8/8c/SENAI_S%C3%A3o_Paulo_logo.png", width=120)
    st.title("Portal de Avaliação Prática")
    st.markdown('<p class="centered-subtitle">Cursos de Tecnologia da Informação - Excel Completo</p>', unsafe_allow_html=True)
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
    st.image("https://upload.wikimedia.org/wikipedia/commons/8/8c/SENAI_S%C3%A3o_Paulo_logo.png", width=100)
    st.title("Entrega e Correção")
    st.markdown(f'<p class="centered-subtitle">Aluno: {st.session_state.aluno["nome"]} | Turma: {st.session_state.aluno["turma"]}</p>', unsafe_allow_html=True)
    st.download_button("📥 Baixar minha Prova Individual", st.session_state.excel_data, st.session_state.nome_arquivo)
    st.divider()
    arquivo_upload = st.file_uploader("Anexe sua prova resolvida (.xlsx)", type=['xlsx'])
    if st.button("🚀 Enviar e Obter Resultado"):
        if arquivo_upload:
            with st.spinner('Analisando sua prova...'):
                nota, feedback = corrigir_prova(arquivo_upload)
                corpo = f"Aluno: {st.session_state.aluno['nome']}\nNota: {nota}\n{feedback}"
                enviar_email(EMAIL_PROFESSOR, f"NOTA {nota}: {st.session_state.aluno['nome']}", corpo, arquivo_upload.getvalue(), arquivo_upload.name)
                enviar_email(st.session_state.aluno['email'], "Resultado - SENAI", corpo)
                if nota > 0:
                    st.success(f"Sua nota final é: {nota}")
                    st.info(feedback)
                    st.balloons()
                else:
                    st.error(feedback)
    if st.button("🚪 Sair"):
        st.session_state.clear()
        st.rerun()
