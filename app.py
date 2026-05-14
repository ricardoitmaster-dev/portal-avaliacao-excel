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
        .stDownloadButton>button {{ background-color: {CORE_SENAI} !important; color: white !important; font-weight: bold; }}
    </style>
""", unsafe_allow_html=True)

# --- GERADOR DE PLANILHA DINÂMICA ---
def gerar_desafio_excel(nome_aluno):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        # Aba 1: Desafio Principal
        itens = ["Monitor LED", "Teclado Mecânico", "Mouse Óptico", "SSD 1TB", "Placa de Vídeo", "Fonte 600W"]
        dados = []
        for i in range(1, 21):
            dados.append({
                "ID": i,
                "Produto": random.choice(itens),
                "Quantidade": random.randint(5, 50),
                "Preço Unit": round(random.uniform(100, 1500), 2),
                "Subtotal": 0,          # Aluno faz Mult
                "IPI (10%)": 0,         # Aluno faz Percentual
                "Total Geral": 0,       # Aluno faz Soma
                "Status": "",           # Aluno faz SE/SES
                "Cod_Ref": random.randint(101, 106) # Para PROCV
            })
        df_base = pd.DataFrame(dados)
        df_base.to_excel(writer, sheet_name='DESAFIO_PRATICO', index=False)
        
        # Aba 2: Referência para PROCV
        df_ref = pd.DataFrame({
            "Cod": [101, 102, 103, 104, 105, 106],
            "Categoria": ["Periféricos", "Armazenamento", "Hardware", "Energia", "Vídeo", "Redes"]
        })
        df_ref.to_excel(writer, sheet_name='REFERENCIA', index=False)
        
        # Aba 3: Instruções
        inst = [
            ["INSTRUÇÕES PARA O ALUNO:", nome_aluno],
            ["1. Converta o intervalo da aba DESAFIO_PRATICO em TABELA."],
            ["2. Calcule Subtotal, IPI e Total Geral usando operadores aritméticos."],
            ["3. Use a função SE: Se Total > 5000 'ALTO INVESTIMENTO', senão 'NORMAL'."],
            ["4. Use PROCV para buscar a 'Categoria' na aba REFERENCIA usando o 'Cod_Ref'."],
            ["5. Crie 2 Macros de Ordenação e insira BOTÕES na planilha."],
            ["6. Salve como .XLSM se fizer macros, ou .XLSX se não fizer."]
        ]
        pd.DataFrame(inst).to_excel(writer, sheet_name='INSTRUCOES', index=False, header=False)
    return output.getvalue()

# --- FUNÇÕES DE E-MAIL E CORREÇÃO (Mantidas conforme solicitado) ---
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
    except: pass

def realizar_correcao_completa(lista_arquivos):
    melhor_nota = 0
    relatorio = ""
    for arq in lista_arquivos:
        try:
            df = pd.read_excel(arq, sheet_name='DESAFIO_PRATICO')
            wb = load_workbook(arq, keep_vba=True)
            p = 0
            res = [f"Análise: {arq.name}"]
            # Aritmética (25 pts)
            if all(round(r['Subtotal'],1) == round(r['Quantidade']*r['Preço Unit'],1) for _,r in df.iterrows()):
                p += 25
                res.append("✅ Operações Aritméticas: OK")
            # SE (25 pts)
            if 'Status' in df.columns and not df['Status'].isnull().all():
                p += 25
                res.append("✅ Função SE/SES: OK")
            # PROCV (25 pts)
            if 'Categoria' in df.columns:
                p += 25
                res.append("✅ PROCV: OK")
            # Macros (25 pts)
            if arq.name.endswith('.xlsm') and wb.vba_archive:
                p += 25
                res.append("✅ Macros Detectadas: OK")
            
            if p >= melhor_nota:
                melhor_nota = p
                relatorio = "\n".join(res)
        except: continue
    return melhor_nota, relatorio

# --- INTERFACE ---
if 'etapa' not in st.session_state: st.session_state.etapa = 'login'

if st.session_state.etapa == 'login':
    st.title("Portal de Avaliação SENAI")
    n = st.text_input("Nome Completo")
    t = st.text_input("Turma").strip().upper()
    e = st.text_input("E-mail")
    if st.button("Entrar no Laboratório"):
        if n and t and e:
            st.session_state.aluno = {"nome": n, "turma": t, "email": e}
            st.session_state.prova_origem = gerar_desafio_excel(n)
            st.session_state.etapa = 'prova'
            st.rerun()

elif st.session_state.etapa == 'prova':
    st.title("Atividade Prática")
    st.write(f"Bem-vindo, **{st.session_state.aluno['nome']}**")
    
    # PASSO 1: BAIXAR
    st.subheader("1º Passo: Baixe seu arquivo de trabalho")
    st.download_button(label="📥 Baixar Planilha de Desafios", 
                       data=st.session_state.prova_origem, 
                       file_name=f"Desafio_Excel_{st.session_state.aluno['nome']}.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    
    st.divider()
    
    # PASSO 2: ENVIAR
    st.subheader("2º Passo: Entregar exercício resolvido")
    st.info("Você pode enviar um ou mais arquivos (.xlsx ou .xlsm).")
    uploads = st.file_uploader("Selecione os arquivos", type=['xlsx', 'xlsm'], accept_multiple_files=True)
    
    if st.button("🚀 Finalizar Avaliação"):
        if uploads:
            nota, relat = realizar_correcao_completa(uploads)
            st.subheader(f"Nota Final: {nota}/100")
            st.text(relat)
            
            corpo = f"Aluno: {st.session_state.aluno['nome']}\nNota: {nota}\n\n{relat}"
            enviar_email_com_anexos(EMAIL_PROFESSOR, f"PROVA: {st.session_state.aluno['nome']}", corpo, uploads)
            enviar_email_com_anexos(st.session_state.aluno['email'], "Seu Resultado SENAI", corpo, uploads)
            
            st.success("Tudo pronto! Sua nota e arquivos foram enviados para o professor.")
            if st.button("Sair"):
                st.session_state.clear()
                st.rerun()
