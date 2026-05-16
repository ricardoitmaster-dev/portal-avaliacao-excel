import streamlit as st
import pandas as pd
import random
import io
import smtplib
import os
import hashlib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from openpyxl import load_workbook

# --- CONFIGURAÇÕES DE AMBIENTE ---
try:
    EMAIL_PROFESSOR = st.secrets["EMAIL_PROFESSOR"]
    SENHA_APP_GOOGLE = st.secrets["SENHA_APP_GOOGLE"]
except:
    EMAIL_PROFESSOR = "ricardoitmaster@gmail.com"
    SENHA_APP_GOOGLE = "ugjhusmwnbmgzspv"

# Cores Identidade Visual (BMW Portinari Blue, Black, Gold)
COR_AZUL_BMW = "#002366"
COR_PRETO_BRILHANTE = "#000000"
COR_DOURADO = "#D4AF37"
COR_TEXTO = "#FFFFFF"

st.set_page_config(page_title="Portal de Avaliação Excel", layout="wide")

# --- ESTILIZAÇÃO CSS ---
st.markdown(f"""
    <style>
        .stApp {{ background-color: {COR_PRETO_BRILHANTE} !important; }}
        h1, h2, h3 {{ color: {COR_DOURADO} !important; font-weight: bold; text-align: center; }}
        label, p, span {{ color: {COR_TEXTO} !important; }}
        .stButton>button {{ 
            color: {COR_TEXTO}; 
            background-color: {COR_AZUL_BMW}; 
            border-radius: 10px; 
            border: 1px solid {COR_DOURADO};
            height: 4em;
            font-weight: bold;
            width: 100%;
        }}
        .stButton>button:hover {{ border: 2px solid {COR_TEXTO}; color: {COR_DOURADO}; }}
        .stDownloadButton>button {{ 
            color: {COR_PRETO_BRILHANTE} !important; 
            background-color: {COR_DOURADO} !important; 
            font-weight: bold; 
            width: 100%; 
        }}
        .stTextInput input {{ background-color: #1A1A1A !important; color: white !important; border: 1px solid {COR_AZUL_BMW} !important; }}
        [data-testid="stHorizontalBlock"] {{ align-items: center !important; }}
        
        /* Estilo para o link que parece um botão de sair */
        .btn-sair-link {{
            display: flex;
            justify-content: center;
            align-items: center;
            text-decoration: none !important;
            background-color: #7B0000;
            color: white !important;
            font-weight: bold;
            border: 1px solid white;
            border-radius: 10px;
            height: 3.5em;
            width: 100%;
            margin-top: 30px;
            text-align: center;
        }}
        .btn-sair-link:hover {{
            background-color: #A00000;
            border: 2px solid white;
        }}
    </style>
""", unsafe_allow_html=True)

# --- FUNÇÕES DE APOIO ---
def enviar_email(destinatario, assunto, corpo, arquivos=None):
    try:
        msg = MIMEMultipart()
        msg['From'] = EMAIL_PROFESSOR
        msg['To'] = destinatario
        msg['Subject'] = assunto
        msg.attach(MIMEText(corpo, 'plain'))
        if arquivos:
            for arquivo_bytes, nome_arquivo in arquivos:
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
    seed = int(hashlib.md5(nome_aluno.encode()).hexdigest(), 16) % 10000
    random.seed(seed)
    
    itens = ["Notebook", "Mouse", "Teclado", "Monitor", "Impressora", "Cabo HDMI", "SSD 480GB"]
    categorias = ["Informática", "Periféricos", "Periféricos", "Informática", "Periféricos", "Acessórios", "Hardware"]
    
    df_apoio = pd.DataFrame({
        "ID_Prod": range(1, 8),
        "Produto": itens,
        "Categoria": categorias,
        "Preço Base": [3500.0, 80.0, 150.0, 900.0, 600.0, 45.0, 280.0]
    })
    
    dados = []
    for i in range(1, 31):
        id_prod = random.randint(1, 7)
