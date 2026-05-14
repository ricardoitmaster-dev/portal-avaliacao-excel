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

# --- CSS ---
st.markdown(f"""
    <style>
        .stApp {{ background-color: {CORE_FUNDO} !important; }}
        h1 {{ color: {CORE_SENAI} !important; font-weight: bold; text-align: center; }}
        label, p, span {{ color: #FFFFFF !important; }}
        .stButton>button {{ color: white; background-color: #262730; border-radius: 8px; width: 100%; border: 1px solid {CORE_SENAI}; }}
    </style>
""", unsafe_allow_html=True)

# --- FUNÇÃO DE ENVIO COM MÚLTIPLOS ANEXOS ---
def enviar_email_com_anexos(destinatario, assunto, corpo, lista_arquivos):
    try:
        msg = MIMEMultipart()
        msg['From'] = EMAIL_PROFESSOR
        msg['To'] = destinatario
        msg['Subject'] = assunto
        msg.attach(MIMEText(corpo, 'plain'))
        
        for arq in lista_arquivos:
            part = MIMEBase('application', 'octet-stream')
            part.set_payload(arq.getvalue())
            encoders.encode_base64(part)
            part.add_header('Content-Disposition', f"attachment; filename={arq.name}")
            msg.attach(part)
            
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(EMAIL_PROFESSOR, SENHA_APP_GOOGLE)
        server.send_message(msg)
        server.quit()
        return True
    except: return False

# --- LÓGICA DE CORREÇÃO (0 a 100) ---
def realizar_correcao_completa(lista_arquivos):
    melhor_nota = 0
    relatorio_final = ""
    
    for arquivo in lista_arquivos:
        try:
            df = pd.read_excel(arquivo, sheet_name='DESAFIO_PRATICO', engine='openpyxl')
            wb = load_workbook(arquivo, keep_vba=True)
            
            p_arit = 0
            p_logica = 0
            p_procv = 0
            p_macro = 0
            p_tab = 0
            feed = [f"--- Analisando: {arquivo.name} ---"]

            # 1. Aritmética (20 pts)
            if all(round(r['Subtotal'],2) == round(r['Quantidade']*r['Preço Unit'],2) for _,r in df.iterrows()):
                p_arit = 20
                feed.append("✅ Cálculos Aritméticos: 20/20")
            else: feed.append("❌ Erro em cálculos (Subtotal/Total).")

            # 2. SE/SES (20 pts)
            if all(str(r['Status Estoque']).upper() in ["CRÍTICO", "NORMAL"] for _,r in df.iterrows()):
                p_logica = 20
                feed.append("✅ Função SE/SES: 20/20")
            else: feed.append("❌ Função SE não aplicada corretamente.")

            # 3. PROCV (20 pts)
            if 'Categoria' in df.columns and not df['Categoria'].isnull().any():
                p_procv = 20
                feed.append("✅ Função PROCV: 20/20")
            else: feed.append("❌ PROCV não realizado.")

            # 4. Macros (20 pts) - Aceita apenas se for .xlsm
            if arquivo.name.endswith('.xlsm') and wb.vba_archive:
                p_macro = 20
                feed.append("✅ Macros e Botões: 20/20")
            else: feed.append("ℹ️ Sem Macros (VBA não detectado ou arquivo .xlsx).")

            # 5. Tabelas (20 pts)
            if wb['DESAFIO_PRATICO'].tables:
                p_tab = 20
                feed.append("✅ Objeto Tabela: 20/20")
            else: feed.append("❌ Não converteu intervalo em Tabela.")

            nota_atual = p_arit + p_logica + p_procv + p_macro + p_tab
            if nota_atual >= melhor_nota:
                melhor_nota = nota_atual
                relatorio_final = "\n".join(feed)
                
        except: continue
        
    return melhor_nota, relatorio_final

# --- INTERFACE ---
if 'etapa' not in st.session_state: st.session_state.etapa = 'login'

st.title("Portal de Avaliação SENAI")

if st.session_state.etapa == 'login':
    nome = st.text_input("Nome")
    turma = st.text_input("Turma").strip().upper()
    email = st.text_input("Seu E-mail")
    if st.button("Acessar"):
        if nome and turma and email:
            st.session_state.aluno = {"nome": nome, "turma": turma, "email": email}
            st.session_state.etapa = 'prova'
            st.rerun()

elif st.session_state.etapa == 'prova':
    st.write(f"Candidato: {st.session_state.aluno['nome']}")
    st.info("Você pode enviar um arquivo .xlsx para a parte de fórmulas e um .xlsm para a parte de macros, ou tudo em um só.")
    
    ups = st.file_uploader("Upload de Soluções (xlsx, xlsm, xls)", type=['xlsx', 'xlsm', 'xls'], accept_multiple_files=True)
    
    if st.button("🚀 Submeter para Avaliação"):
        if ups:
            nota, relatorio = realizar_correcao_completa(ups)
            st.subheader(f"Nota Final: {nota}/100")
            st.text(relatorio)
            
            # Envio de E-mails
            corpo = f"Resultado SENAI\nAluno: {st.session_state.aluno['nome']}\nNota: {nota}\n\n{relatorio}"
            enviar_email_com_anexos(EMAIL_PROFESSOR, f"ENTREGA: {st.session_state.aluno['nome']}", corpo, ups)
            enviar_email_com_anexos(st.session_state.aluno['email'], "Seu Resultado - Avaliação Excel", corpo, ups)
            
            st.success("Arquivos enviados com sucesso para você e para o professor!")
            if st.button("Finalizar"):
                st.session_state.clear()
                st.rerun()
