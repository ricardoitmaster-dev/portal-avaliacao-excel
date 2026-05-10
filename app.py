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

# --- IDENTIDADE VISUAL (DARK MASTER) ---
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

# --- FUNÇÕES DE MOTOR ---
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
    dados = [{"ID": i, "Produto": random.choice(itens), "Quantidade": random.randint(5, 50), 
              "Preço Unitário": round(random.uniform(20, 300), 2), "Venda Total": 0, "Status": ""} for i in range(1, 31)]
    
    df = pd.DataFrame(dados)
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, sheet_name='Base_de_Dados', index=False)
        # TOKEN DE IDENTIDADE ÚNICO NA ABA INSTRUCOES
        inst = [["AVALIAÇÃO OFICIAL SENAI"], 
                [f"ID_ALUNO: {nome_aluno}"],
                [f"TURMA: {turma}"]]
        pd.DataFrame(inst).to_excel(writer, sheet_name='Instrucoes', index=False, header=False)
    return output.getvalue()

def validar_e_corrigir(arquivo_aluno, nome_logado):
    try:
        xls = pd.ExcelFile(arquivo_aluno, engine='openpyxl')
        if 'Instrucoes' not in xls.sheet_names:
            return None, "ARQUIVO INVÁLIDO! Use o modelo oficial do SENAI."

        # VERIFICAÇÃO RIGOROSA DO NOME DENTRO DO ARQUIVO
        df_inst = pd.read_excel(xls, sheet_name='Instrucoes', header=None)
        nome_no_arquivo = str(df_inst.iloc[1, 0]).replace("ID_ALUNO: ", "").strip()
        
        # Só prossegue se o nome no arquivo for exatamente igual ao logado
        if nome_no_arquivo.upper() != nome_logado.upper():
            return None, f"ERRO: Este arquivo foi gerado para {nome_no_arquivo}. Você deve enviar o SEU arquivo."

        df = pd.read_excel(xls, sheet_name='Base_de_Dados')
        pv, ps, total = 0, 0, len(df)
        
        for _, row in df.iterrows():
            calc_correto = round(float(row['Quantidade'] * row['Preço Unitário']), 2)
            if round(float(row['Venda Total']), 2) == calc_correto:
                pv += 1
            meta = "META" if calc_correto >= 500 else "REVISAR"
            if str(row['Status']).strip().upper() == meta:
                ps += 1
        
        nota = round(((pv / total) * 5) + ((ps / total) * 5), 1)
        return nota, f"Cálculos: {pv}/{total} | Lógica SE: {ps}/{total}"
    except:
        return None, "ERRO TÉCNICO: Não foi possível processar o arquivo. Verifique a estrutura."

# --- INTERFACE ---
if 'etapa' not in st.session_state: st.session_state.etapa = 'login'

if st.session_state.etapa == 'login':
    st.image("https://upload.wikimedia.org/wikipedia/commons/8/8c/SENAI_S%C3%A3o_Paulo_logo.png", width=120)
    st.title("Portal de Avaliação Prática")
    st.markdown('<p class="centered-subtitle">Cursos de Tecnologia da Informação - SENAI</p>', unsafe_allow_html=True)
    nome = st.text_input("Seu Nome Completo")
    turma = st.text_input("Turma")
    email = st.text_input("E-mail para receber nota")
    if st.button("Iniciar Avaliação"):
        if nome and turma and email:
            st.session_state.excel_data = gerar_prova_excel(nome, turma)
            st.session_state.nome_arquivo = f"Avaliacao_{nome.replace(' ','_')}.xlsx"
            st.session_state.aluno = {"nome": nome, "turma": turma, "email": email}
            st.session_state.etapa = 'prova'
            st.rerun()
else:
    st.image("https://upload.wikimedia.org/wikipedia/commons/8/8c/SENAI_S%C3%A3o_Paulo_logo.png", width=100)
    st.title("Entrega da Prova")
    st.markdown(f'<p class="centered-subtitle">Aluno: {st.session_state.aluno["nome"]}</p>', unsafe_allow_html=True)
    st.download_button("📥 1. Baixar Minha Prova", st.session_state.excel_data, st.session_state.nome_arquivo)
    st.divider()
    
    arquivo_upload = st.file_uploader("2. Enviar Prova Resolvida", type=['xlsx'])
    
    if st.button("🚀 Finalizar e Corrigir"):
        if arquivo_upload:
            with st.spinner('Validando...'):
                nota, resultado = validar_e_corrigir(arquivo_upload, st.session_state.aluno["nome"])
                
                if nota is not None:
                    corpo = f"Aluno: {st.session_state.aluno['nome']}\nTurma: {st.session_state.aluno['turma']}\nNota: {nota}\n{resultado}"
                    enviar_email(EMAIL_PROFESSOR, f"NOTA {nota}: {st.session_state.aluno['nome']}", corpo, arquivo_upload.getvalue(), arquivo_upload.name)
                    enviar_email(st.session_state.aluno['email'], "Seu Resultado - Excel SENAI", corpo)
                    st.success(f"Nota Final: {nota}")
                    st.info(resultado)
                    st.balloons()
                else:
                    st.error(resultado)
        else:
            st.warning("Selecione seu arquivo.")

    if st.button("Sair"):
        st.session_state.clear()
        st.rerun()
