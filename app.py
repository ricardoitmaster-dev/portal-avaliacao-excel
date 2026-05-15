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

st.set_page_config(page_title="Portal de Avaliação Excel", layout="centered")

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
            height: 3em;
            font-weight: bold;
        }}
        .stButton>button:hover {{ border: 2px solid {COR_TEXTO}; color: {COR_DOURADO}; }}
        .stDownloadButton>button {{ 
            color: {COR_PRETO_BRILHANTE} !important; 
            background-color: {COR_DOURADO} !important; 
            font-weight: bold; 
            width: 100%; 
        }}
        .stTextInput input {{ background-color: #1A1A1A !important; color: white !important; border: 1px solid {COR_AZUL_BMW} !important; }}
    </style>
""", unsafe_allow_html=True)

# --- FUNÇÕES DE APOIO ---
def gerar_feedback_pedagogico():
    return """
--------------------------------------------------
🎓 GUIA DE CORREÇÃO E BOAS PRÁTICAS
--------------------------------------------------
1. CÁLCULO DE FATURAMENTO: Fórmula: =C2*D2
2. LÓGICA CONDICIONAL: Fórmula: =SE(E2>=500;"META";"REVISAR")
3. FORMATAÇÃO: Use padrão Moeda (R$) e salve como .XLSM para macros.
--------------------------------------------------
"""

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
    except: return False

def gerar_prova_excel(nome_aluno):
    itens = ["Notebook", "Mouse", "Teclado", "Monitor", "Impressora", "Cabo HDMI", "SSD 480GB"]
    dados = [{"ID": i, "Produto": random.choice(itens), "Quantidade": random.randint(5, 50), 
              "Preço Unitário": round(random.uniform(20, 300), 2), "Venda Total": 0, "Status": ""} for i in range(1, 31)]
    df = pd.DataFrame(dados)
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, sheet_name='Base_de_Dados', index=False)
        inst = [["AVALIAÇÃO PRÁTICA"], [f"Aluno: {nome_aluno}"], ["1. Calcule Venda Total"], ["2. Use SE para Status (Meta 500)"], ["3. Crie Macros e salve como .xlsm"]]
        pd.DataFrame(inst).to_excel(writer, sheet_name='Instrucoes', index=False, header=False)
    return output.getvalue()

def calcular_nota(arquivo_bytes):
    try:
        df = pd.read_excel(arquivo_bytes, sheet_name='Base_de_Dados', engine='openpyxl')
        pv, ps, total = 0, 0, len(df)
        for _, row in df.iterrows():
            calc = round(float(row['Quantidade'] * row['Preço Unitário']), 2)
            if round(float(row['Venda Total']), 2) == calc: pv += 1
            meta = "META" if calc >= 500 else "REVISAR"
            if str(row['Status']).strip().upper() == meta: ps += 1
        try:
            wb = load_workbook(arquivo_bytes, keep_vba=True)
            tem_macro = 2.0 if wb.vba_archive else 0.0
        except: tem_macro = 0.0
        nota = round(((pv / total) * 4) + ((ps / total) * 4) + tem_macro, 1)
        return nota, f"Cálculos: {pv}/{total} | Lógica SE: {ps}/{total} | Macro: {'Sim' if tem_macro > 0 else 'Não'}"
    except: return 0, "Erro na leitura do arquivo."

# --- LÓGICA DE NAVEGAÇÃO ---
if 'perfil' not in st.session_state:
    st.session_state.perfil = None

# Cabeçalho Fixo
col_l, col_r = st.columns([1, 1])
with col_l: st.image("https://upload.wikimedia.org/wikipedia/commons/8/8c/SENAI_S%C3%A3o_Paulo_logo.png", width=100)
with col_r: st.image("Imagem para o app avaliação Excel_RicardoItmaster.png", width=180)

# --- TELA DE SELEÇÃO DE PERFIL ---
if st.session_state.perfil is None:
    st.title("Sistema de Avaliação Técnica")
    st.write("### Bem-vindo! Selecione seu perfil de acesso:")
    
    col1, col2 = st.columns(2)
    with col1:
        if st.button("🎓 SOU ALUNO"):
            st.session_state.perfil = "aluno"
            st.rerun()
    with col2:
        if st.button("👨‍🏫 SOU PROFESSOR / GESTOR"):
            st.session_state.perfil = "admin"
            st.rerun()

# --- ÁREA DO ALUNO ---
elif st.session_state.perfil == "aluno":
    if 'etapa_aluno' not in st.session_state: st.session_state.etapa_aluno = 'login'
    
    if st.button("⬅️ Voltar ao Início"):
        st.session_state.perfil = None
        st.rerun()

    if st.session_state.etapa_aluno == 'login':
        st.subheader("Acesso do Estudante")
        with st.form("login_aluno"):
            nome = st.text_input("Nome Completo")
            turma = st.text_input("Turma")
            email = st.text_input("Seu E-mail")
            if st.form_submit_button("Entrar no Ambiente"):
                if nome and turma and email:
                    st.session_state.aluno_dados = {"nome": nome, "turma": turma.upper(), "email": email}
                    st.session_state.excel_data = gerar_prova_excel(nome)
                    st.session_state.etapa_aluno = 'prova'
                    st.rerun()

    elif st.session_state.etapa_aluno == 'prova':
        st.subheader(f"Área de Entrega: {st.session_state.aluno_dados['nome']}")
        st.download_button("📥 Baixar Planilha de Avaliação", st.session_state.excel_data, f"Avaliacao_{st.session_state.aluno_dados['nome']}.xlsx")
        
        st.divider()
        up = st.file_uploader("Upload da Planilha Respondida", type=['xls', 'xlsx', 'xlsm'], accept_multiple_files=True)
        
        if st.button("🚀 Enviar para Correção"):
            if up:
                # Lógica de validação de nome e correção (igual ao anterior)
                validos = [f for f in up if f.name.split('.')[0] == f"Avaliacao_{st.session_state.aluno_dados['nome']}".replace(' ','_')]
                if not validos:
                    st.error("Erro: O arquivo enviado não pertence a este aluno ou o nome foi alterado.")
                else:
                    arq = next((f for f in validos if f.name.endswith('xlsm')), validos[0])
                    nota, feedback = calcular_nota(arq)
                    
                    # Salva no DB
                    novo = pd.DataFrame([[st.session_state.aluno_dados['nome'], st.session_state.aluno_dados['turma'], nota]], columns=['Aluno', 'Turma', 'Nota'])
                    novo.to_csv("db_notas.csv", mode='a', header=not os.path.exists("db_notas.csv"), index=False)
                    
                    # Envia E-mails
                    corpo = f"Resultado: {nota}\n{feedback}\n{gerar_feedback_pedagogico()}"
                    anexos = [(f.getvalue(), f.name) for f in validos]
                    enviar_email(EMAIL_PROFESSOR, f"PROVA: {st.session_state.aluno_dados['nome']}", corpo, anexos)
                    enviar_email(st.session_state.aluno_dados['email'], "Seu Feedback SENAI", corpo, anexos)
                    
                    st.success(f"Enviado com sucesso! Nota: {nota}")
                    st.balloons()

# --- ÁREA ADMINISTRATIVA ---
elif st.session_state.perfil == "admin":
    if st.button("⬅️ Voltar ao Início"):
        st.session_state.perfil = None
        st.rerun()

    tabs = st.tabs(["Acesso Professor", "Novo Cadastro", "Gerência (ADM)"])

    with tabs[0]: # Login Professor
        if not st.session_state.get('prof_logado', False):
            with st.form("l_prof"):
                t = st.text_input("Turma").upper()
                s = st.text_input("Senha", type="password")
                if st.form_submit_button("Entrar"):
                    if os.path.exists("professores.csv"):
                        dfp = pd.read_csv("professores.csv")
                        if not dfp[(dfp['Turma'] == t) & (dfp['Senha'] == str(s))].empty:
                            st.session_state.prof_logado = True
                            st.session_state.prof_turma = t
                            st.rerun()
                    st.error("Dados inválidos.")
        else:
            st.write(f"### Resultados da Turma {st.session_state.prof_turma}")
            if os.path.exists("db_notas.csv"):
                dfn = pd.read_csv("db_notas.csv")
                st.dataframe(dfn[dfn['Turma'] == st.session_state.prof_turma], use_container_width=True)
            if st.button("Sair"): st.session_state.prof_logado = False; st.rerun()

    with tabs[1]: # Cadastro
        with st.form("cad_p", clear_on_submit=True):
            np = st.text_input("Nome Professor")
            tp = st.text_input("Turma").upper()
            sp = st.text_input("Senha", type="password")
            if st.form_submit_button("Cadastrar"):
                pd.DataFrame([[np, tp, sp]], columns=['Professor','Turma','Senha']).to_csv("professores.csv", mode='a', header=not os.path.exists("professores.csv"), index=False)
                st.success("Cadastrado!")

    with tabs[2]: # ADM Central
        if not st.session_state.get('adm_logado', False):
            with st.form("l_adm"):
                sadm = st.text_input("Senha Mestra", type="password")
                if st.form_submit_button("Acessar"):
                    if sadm == "Celina2610$$":
                        st.session_state.adm_logado = True
                        st.rerun()
                    else: st.error("Senha Incorreta.")
        else:
            st.write("### Relatório Geral (Todas as Turmas)")
            if os.path.exists("db_notas.csv"): st.dataframe(pd.read_csv("db_notas.csv"), use_container_width=True)
            if st.button("Logout ADM"): st.session_state.adm_logado = False; st.rerun()
