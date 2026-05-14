import streamlit as st
import pandas as pd
import random
import io
import smtplib
import os
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from openpyxl import load_workbook

# --- CONFIGURAÇÕES ---
try:
    EMAIL_PROFESSOR = st.secrets["EMAIL_PROFESSOR"]
    SENHA_APP_GOOGLE = st.secrets["SENHA_APP_GOOGLE"]
except:
    EMAIL_PROFESSOR = "ricardoitmaster@gmail.com"
    SENHA_APP_GOOGLE = "ugjhusmwnbmgzspv"

CORE_SENAI = "#FF0000"
CORE_FUNDO = "#0E1117" 

st.set_page_config(page_title="Portal de Avaliação Excel - SENAI", layout="centered")

# --- CSS E ALINHAMENTO ---
st.markdown(f"""
    <style>
        .stApp {{ background-color: {CORE_FUNDO} !important; }}
        h1 {{ color: {CORE_SENAI} !important; font-weight: bold; text-align: center; }}
        label, p, span {{ color: #FFFFFF !important; }}
        .stButton>button {{ color: white; background-color: #262730; border-radius: 8px; width: 100%; border: 1px solid {CORE_SENAI}; }}
        /* Alinhamento vertical das logos */
        [data-testid="stHorizontalBlock"] {{ align-items: center; }}
    </style>
""", unsafe_allow_html=True)

def encerrar_sessao():
    for key in list(st.session_state.keys()):
        del st.session_state[key]
    st.rerun()

# --- FUNÇÃO DE CABEÇALHO (LOGOS) ---
def exibir_cabecalho():
    col_logo_senai, col_vazia, col_logo_ricardo = st.columns([1, 1, 1])
    with col_logo_senai:
        st.image("
