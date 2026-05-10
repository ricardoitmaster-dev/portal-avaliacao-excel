import streamlit as st
import pandas as pd
import random
import io
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

# --- CONFIGURAÇÕES VIA SECRETS ---
try:
    EMAIL_PROFESSOR = st.secrets["EMAIL_PROFESSOR"]
    SENHA_APP_GOOGLE = st.secrets["SENHA_APP_GOOGLE"]
except:
    EMAIL_PROFESSOR = "ricardoitmaster@gmail.com"
    SENHA_APP_GOOGLE = "ugjhusmwnbmgzspv"

# --- IDENTIDADE VISUAL ---
CORE_SENAI = "#FF0000"
CORE_FUNDO = "#0E1117" 
CORE_TEXTO_BRANCO = "#FFFFFF"

st.set_page_config(page_title="Portal de Avaliação Excel - SENAI", layout="centered")

st.markdown(f"""
    <style>
        .stApp {{ background-color: {CORE_FUNDO} !important; }}
        h1 {{ color: {CORE_SENAI} !important; font-weight: bold; text-align: center !important; }}
        .centered-subtitle {{ text-align: center !important; color: {CORE_TEXTO_BRANCO} !important; font-size: 1.2rem; margin-bottom: 30px; }}
        label, p, span {{ color: {CORE_TEXTO_BRANCO} !important; }}
        .stButton>button {{ color: white; background-color: #262730; border-radius: 8px; width: 100%; border: 1px solid {CORE_SENAI}; }}
        .stButton>button:hover {{ background-color: {CORE_SENAI}; }}
        .stDownloadButton>button {{ color: white !important; background-color: {CORE_SENAI} !important; font-weight: bold; width: 100%; }}
        .stTextInput input {{ background-color: #262730 !important; color: white !important; }}
    </style>
""", unsafe_allow_html=True)

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

def gerar_prova_excel(nome_aluno):
    itens = ["Notebook", "Mouse", "Teclado", "Monitor", "Impressora", "Cabo HDMI"]
    dados = [{"ID": i, "Produto": random.choice(itens), "Quantidade": random.randint(5, 50), 
              "Preço Unitário": round(random.uniform(20, 300), 2), "Venda Total": 0, "Status": ""} for i in range(1, 31)]
    df = pd.DataFrame(dados)
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, sheet_name='Base_de_Dados', index=False)
    return output.getvalue()

def calcular_nota(arquivo_bytes):
    try:
        df = pd.read_excel(arquivo_bytes, sheet_name='Base_de_Dados', engine='openpyxl')
        pv, ps, total = 0, 0, len(df)
        for _, row in df.iterrows():
            correto = round(float(row['Quantidade'] * row['Preço Unitário']), 2)
            if round(float(row['Venda Total']), 2) == correto: pv += 1
            meta = "META" if correto >= 500 else "REVISAR"
            if str(row['Status']).strip().upper() == meta: ps += 1
        nota = round(((pv / total) * 5) + ((ps / total) * 5), 1)
        return nota, f"Cálculos: {pv}/{total} | Lógica SE: {ps}/{total}"
    except:
        return 0, "Erro na leitura das colunas."

# --- LOGICA DE NAVEGAÇÃO ---
if 'etapa' not in st.session_state: st.session_state.etapa = 'login'

if st.session_state.etapa == 'login':
    st.image("https://upload.wikimedia.org/wikipedia/commons/8/8c/SENAI_S%C3%A3o_Paulo_logo.png", width=120)
    st.title("Portal de Avaliação Prática")
    nome = st.text_input("Nome Completo")
    turma = st.text_input("Turma")
    email = st.text_input("Seu E-mail")
    
    if st.button("Acessar Prova"):
        if nome and turma and email:
            st.session_state.aluno = {"nome": nome, "turma": turma, "email": email}
            # Criamos o nome do arquivo esperado EXATAMENTE aqui
            st.session_state.nome_arquivo_esperado = f"Avaliacao_{nome.replace(' ','_')}.xlsx"
            st.session_state.excel_data = gerar_prova_excel(nome)
            st.session_state.etapa = 'prova'
            st.rerun()

else:
    st.image("https://upload.wikimedia.org/wikipedia/commons/8/8c/SENAI_S%C3%A3o_Paulo_logo.png", width=100)
    st.title("Entrega e Validação")
    st.write(f"Aluno: **{st.session_state.aluno['nome']}**")
    
    st.download_button("📥 1. Baixar Minha Prova Individual", 
                       st.session_state.excel_data, 
                       st.session_state.nome_arquivo_esperado)
    
    st.divider()
    
    arquivo_upload = st.file_uploader("2. Envie o arquivo baixado (não altere o nome)", type=['xlsx'])
    
    if st.button("🚀 3. Validar e Enviar"):
        if arquivo_upload is not None:
            # A TRAVA REAL: Compara o nome do arquivo enviado com o esperado
            if arquivo_upload.name != st.session_state.nome_arquivo_esperado:
                st.error(f"ARQUIVO INCORRETO! Você enviou '{arquivo_upload.name}', mas deve enviar o seu arquivo oficial: '{st.session_state.nome_arquivo_esperado}'.")
            else:
                # SE PASSOU NA TRAVA DE NOME, SEGUE PARA O ENVIO
                with st.spinner('Corrigindo e Enviando...'):
                    nota, res = calcular_nota(arquivo_upload)
                    corpo = f"Aluno: {st.session_state.aluno['nome']}\nNota: {nota}\n{res}"
                    
                    # Envia e-mails
                    enviar_email(EMAIL_PROFESSOR, f"NOTA {nota}: {st.session_state.aluno['nome']}", corpo, arquivo_upload.getvalue(), arquivo_upload.name)
                    enviar_email(st.session_state.aluno['email'], "Resultado Avaliação SENAI", corpo)
                    
                    st.success(f"Enviado com sucesso! Nota Final: {nota}")
                    st.info(res)
                    st.balloons()
        else:
            st.warning("Selecione o arquivo antes de clicar em enviar.")

    if st.button("Sair"):
        st.session_state.clear()
        st.rerun()
