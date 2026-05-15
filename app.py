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
            height: 3em;
            font-weight: bold;
            width: 100%; /* Adicionado para garantir preenchimento total da coluna */
        }}
        .stButton>button:hover {{ border: 2px solid {COR_TEXTO}; color: {COR_DOURADO}; }}
        .stDownloadButton>button {{ 
            color: {COR_PRETO_BRILHANTE} !important; 
            background-color: {COR_DOURADO} !important; 
            font-weight: bold; 
            width: 100%; 
        }}
        .stTextInput input {{ background-color: #1A1A1A !important; color: white !important; border: 1px solid {COR_AZUL_BMW} !important; }}
        /* Ajuste fino para alinhamento de imagens */
        [data-testid="stHorizontalBlock"] {{ align-items: center !important; }}
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

# --- CABEÇALHO CENTRALIZADO (LOGOS) ---
col_logo_l, col_centro, col_logo_r = st.columns([1, 2, 1])
with col_logo_l:
    st.image("https://upload.wikimedia.org/wikipedia/commons/8/8c/SENAI_S%C3%A3o_Paulo_logo.png", width=120)
with col_logo_r:
    st.image("Imagem para o app avaliação Excel_RicardoItmaster.png", width=220)

# --- LÓGICA DE NAVEGAÇÃO ---
if 'perfil' not in st.session_state: st.session_state.perfil = None

# --- TELA DE SELEÇÃO DE PERFIL CENTRALIZADA ---
if st.session_state.perfil is None:
    st.title("Sistema de Avaliação Técnica")
    st.write("### Bem-vindo! Selecione seu perfil de acesso:")
    # Alteração: Uso de 4 colunas para centralizar os botões
    _, col1, col2, _ = st.columns([1, 1, 1, 1])
    with col1:
        if st.button("🎓 SOU ALUNO"):
            st.session_state.perfil = "aluno"; st.rerun()
    with col2:
        if st.button("👨‍🏫 SOU PROFESSOR / GESTOR"):
            st.session_state.perfil = "admin"; st.rerun()

# --- ÁREA DO ALUNO ---
elif st.session_state.perfil == "aluno":
    if 'etapa_aluno' not in st.session_state: st.session_state.etapa_aluno = 'login'
    
    if st.button("⬅️ Voltar"):
        st.session_state.clear(); st.rerun()

    if st.session_state.etapa_aluno == 'login':
        st.subheader("Acesso do Estudante")
        with st.form("login_aluno"):
            nome = st.text_input("Nome Completo").strip()
            turma = st.text_input("Turma").strip().upper()
            email = st.text_input("E-mail").strip()
            if st.form_submit_button("Acessar Prova"):
                if nome and turma and email:
                    st.session_state.aluno_dados = {"nome": nome, "turma": turma, "email": email}
                    st.session_state.excel_data = gerar_prova_excel(nome)
                    st.session_state.etapa_aluno = 'prova'; st.rerun()

    elif st.session_state.etapa_aluno == 'prova':
        st.subheader(f"Área de Entrega: {st.session_state.aluno_dados['nome']}")
        nome_arquivo_esperado = f"Avaliacao_{st.session_state.aluno_dados['nome'].replace(' ', '_')}"
        st.download_button("📥 Baixar Planilha de Avaliação", st.session_state.excel_data, f"{nome_arquivo_esperado}.xlsx")
        
        st.divider()
        up = st.file_uploader("Upload da Planilha Respondida", type=['xls', 'xlsx', 'xlsm'], accept_multiple_files=True)
        
        if st.button("🚀 Enviar para Correção"):
            if up:
                # Validação robusta (Case insensitive e remove espaços extras)
                validos = []
                for f in up:
                    nome_enviado = f.name.split('.')[0].lower().strip()
                    if nome_enviado == nome_arquivo_esperado.lower().strip():
                        validos.append(f)
                
                if not validos:
                    st.error(f"Erro: O arquivo enviado não pertence a este aluno. O nome do arquivo deve ser exatamente: {nome_arquivo_esperado}")
                else:
                    arq = next((f for f in validos if f.name.endswith('xlsm')), validos[0])
                    nota, feedback = calcular_nota(arq)
                    pd.DataFrame([[st.session_state.aluno_dados['nome'], st.session_state.aluno_dados['turma'], nota]], columns=['Aluno', 'Turma', 'Nota']).to_csv("db_notas.csv", mode='a', header=not os.path.exists("db_notas.csv"), index=False)
                    enviar_email(EMAIL_PROFESSOR, f"PROVA: {st.session_state.aluno_dados['nome']}", f"Nota: {nota}\n{feedback}", [(f.getvalue(), f.name) for f in validos])
                    st.success(f"Enviado com sucesso! Nota: {nota}"); st.balloons()

# --- ÁREA ADMINISTRATIVA ---
elif st.session_state.perfil == "admin":
    if st.button("⬅️ Voltar"):
        st.session_state.clear(); st.rerun()

    tabs = st.tabs(["Acesso Professor", "Novo Cadastro", "Gerência (ADM)"])

    with tabs[0]: # Login Professor (Corrigido)
        if not st.session_state.get('prof_autenticado', False):
            with st.form("login_prof_v2"):
                n_prof = st.text_input("Nome do Professor").strip()
                s_prof = st.text_input("Senha", type="password")
                if st.form_submit_button("Validar Cadastro"):
                    if os.path.exists("professores.csv"):
                        dfp = pd.read_csv("professores.csv")
                        match = dfp[(dfp['Professor'].str.lower() == n_prof.lower()) & (dfp['Senha'] == str(s_prof))]
                        if not match.empty:
                            st.session_state.prof_autenticado = True
                            st.session_state.prof_nome_logado = n_prof
                            st.session_state.turmas_disponiveis = match['Turma'].unique().tolist()
                            st.rerun()
                    st.error("Professor não cadastrado ou senha incorreta.")
        else:
            st.write(f"### Olá, Prof. {st.session_state.prof_nome_logado}")
            turma_sel = st.selectbox("Selecione a Turma que deseja acessar:", st.session_state.turmas_disponiveis)
            if os.path.exists("db_notas.csv"):
                dfn = pd.read_csv("db_notas.csv")
                st.dataframe(dfn[dfn['Turma'] == turma_sel], use_container_width=True)
            if st.button("Sair da Sessão"): st.session_state.clear(); st.rerun()

    with tabs[1]: # Cadastro
        with st.form("cad_p"):
            np = st.text_input("Nome Completo do Professor").strip()
            tp = st.text_input("Turma para Vincular").strip().upper()
            sp = st.text_input("Definir Senha de Acesso", type="password")
            if st.form_submit_button("Cadastrar Docente"):
                if np and tp and sp:
                    pd.DataFrame([[np, tp, sp]], columns=['Professor','Turma','Senha']).to_csv("professores.csv", mode='a', header=not os.path.exists("professores.csv"), index=False)
                    st.success(f"Professor {np} cadastrado na turma {tp}!"); st.balloons()

    with tabs[2]: # ADM Central
        if not st.session_state.get('adm_logado', False):
            with st.form("l_adm"):
                if st.form_submit_button("Acessar Painel Mestre") and st.text_input("Senha Mestra", type="password") == "Celina2610$$":
                    st.session
