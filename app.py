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
    except: return False

def gerar_prova_excel(nome_aluno):
    itens = ["Notebook", "Mouse", "Teclado", "Monitor", "Impressora", "Cabo HDMI", "SSD 480GB"]
    dados = []
    for i in range(1, 31):
        item = random.choice(itens)
        qtd = random.randint(5, 50)
        p_u = round(random.uniform(20, 300), 2)
        dados.append({"ID": i, "Produto": item, "Quantidade": qtd, "Preço Unitário": p_u, "Venda Total": 0, "Status": ""})
    df = pd.DataFrame(dados)
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, sheet_name='Base_de_Dados', index=False)
        instrucoes = [
            ["AVALIAÇÃO PRÁTICA DE EXCEL"], [f"Nome do Aluno: {nome_aluno}"], [""],
            ["INSTRUÇÕES:"], ["1. Calcule 'Venda Total' (Quantidade x Preço Unitário)"],
            ["2. Na coluna 'Status', use a função SE (>= 500 META, senão REVISAR)"],
            ["3. Salve o arquivo final como .xlsm (habilitado para macros)"]
        ]
        pd.DataFrame(instrucoes).to_excel(writer, sheet_name='Instrucoes', index=False, header=False)
    return output.getvalue()

def calcular_nota(arquivo_bytes):
    try:
        df = pd.read_excel(arquivo_bytes, sheet_name='Base_de_Dados', engine='openpyxl')
        ac_v, ac_s, total = 0, 0, len(df)
        for _, row in df.iterrows():
            calc = round(float(row['Quantidade'] * row['Preço Unitário']), 2)
            if round(float(row['Venda Total']), 2) == calc: ac_v += 1
            meta = "META" if calc >= 500 else "REVISAR"
            if str(row['Status']).strip().upper() == meta: ac_s += 1
        try:
            wb = load_workbook(arquivo_bytes, keep_vba=True)
            macro = 2.0 if wb.vba_archive else 0.0
        except: macro = 0.0
        nota = round(((ac_v / total) * 4) + ((ac_s / total) * 4) + macro, 1)
        return nota, f"Cálculos: {ac_v}/{total} | Status: {ac_s}/{total} | Macro: {macro}"
    except Exception as e: return 0, f"Erro: {str(e)}"

# --- INICIALIZAÇÃO ---
if 'perfil' not in st.session_state: st.session_state.perfil = None
if 'etapa_aluno' not in st.session_state: st.session_state.etapa_aluno = 'login'

# --- CABEÇALHO ---
c1, c2, c3 = st.columns([1, 2, 1])
with c1: st.image("https://upload.wikimedia.org/wikipedia/commons/8/8c/SENAI_S%C3%A3o_Paulo_logo.png", width=150)
with c3: st.image("Imagem para o app avaliação Excel_RicardoItmaster.png", width=220)

# --- TELA INICIAL ---
if st.session_state.perfil is None:
    st.markdown("<br>", unsafe_allow_html=True)
    st.title("Sistema de Avaliação Técnica")
    st.write("### Bem-vindo! Selecione seu perfil de acesso:")
    
    # DISPOSIÇÃO HORIZONTAL
    _, col1, col2, _ = st.columns([0.8, 1, 1, 0.8])
    with col1:
        if st.button("🎓 SOU ALUNO"):
            st.session_state.perfil = "aluno"; st.rerun()
    with col2:
        if st.button("👨‍🏫 PROFESSOR / GESTOR"):
            st.session_state.perfil = "admin"; st.rerun()
    
    # SAIR COM LINK PARA SENAI GUARULHOS
    _, col_sair, _ = st.columns([1.5, 1, 1.5])
    with col_sair:
        st.markdown(
            '<a href="https://guarulhos.sp.senai.br/" target="_self" class="btn-sair-link">❌ SAIR DO SISTEMA</a>', 
            unsafe_allow_html=True
        )

# --- ÁREA DO ALUNO ---
elif st.session_state.perfil == "aluno":
    if st.button("⬅️ Voltar"): st.session_state.perfil = None; st.rerun()
    if st.session_state.etapa_aluno == 'login':
        st.subheader("Identificação")
        with st.form("f_aluno"):
            n, t, e = st.text_input("Nome"), st.text_input("Turma").upper(), st.text_input("E-mail")
            if st.form_submit_button("Gerar Minha Prova"):
                if n and t and e:
                    st.session_state.aluno_dados, st.session_state.excel_data = {"nome":n,"turma":t,"email":e}, gerar_prova_excel(n)
                    st.session_state.etapa_aluno = 'prova'; st.rerun()
    elif st.session_state.etapa_aluno == 'prova':
        nome_e = f"Avaliacao_{st.session_state.aluno_dados['nome'].replace(' ', '_')}"
        st.info(f"Envie o arquivo como: {nome_e}")
        st.download_button("📥 Baixar Prova", st.session_state.excel_data, f"{nome_e}.xlsx")
        up = st.file_uploader("Upload", type=['xlsx', 'xlsm'], accept_multiple_files=True)
        if st.button("🚀 Enviar Prova"):
            validos = [f for f in up if f.name.split('.')[0].lower() == nome_e.lower()]
            if validos:
                arq = next((f for f in validos if f.name.endswith('xlsm')), validos[0])
                nota, info = calcular_nota(arq)
                pd.DataFrame([[st.session_state.aluno_dados['nome'], st.session_state.aluno_dados['turma'], nota]], columns=['Aluno','Turma','Nota']).to_csv("db_notas.csv", mode='a', header=not os.path.exists("db_notas.csv"), index=False)
                enviar_email(EMAIL_PROFESSOR, f"Prova - {st.session_state.aluno_dados['nome']}", f"Nota: {nota}", [(f.getvalue(), f.name) for f in up])
                st.success(f"Enviado! Nota: {nota}"); st.balloons()

# --- ÁREA ADMINISTRATIVA ---
elif st.session_state.perfil == "admin":
    if st.button("⬅️ Voltar"): st.session_state.perfil = None; st.rerun()
    t1, t2, t3 = st.tabs(["📊 Notas", "📝 Professores", "🛡️ Área ADM"])
    with t1:
        if os.path.exists("professores.csv"):
            np, sp = st.text_input("Professor"), st.text_input("Senha", type="password")
            if st.button("Consultar"):
                dfp = pd.read_csv("professores.csv")
                auth = dfp[(dfp['Professor'].str.lower() == np.lower()) & (dfp['Senha'] == str(sp))]
                if not auth.empty and os.path.exists("db_notas.csv"):
                    db = pd.read_csv("db_notas.csv")
                    st.dataframe(db[db['Turma'].isin(auth['Turma'].unique())], use_container_width=True)
    with t2:
        with st.form("c_p"):
            n_p, t_p, s_p = st.text_input("Nome"), st.text_input("Turma"), st.text_input("Senha")
            if st.form_submit_button("Cadastrar"):
                pd.DataFrame([[n_p, t_p, s_p]], columns=['Professor','Turma','Senha']).to_csv("professores.csv", mode='a', header=not os.path.exists("professores.csv"), index=False)
                st.success("OK!")
    with t3:
        if st.text_input("Senha ADM", type="password") == "Celina2610$$":
            if os.path.exists("db_notas.csv"):
                st.dataframe(pd.read_csv("db_notas.csv"), use_container_width=True)
                if st.button("Limpar Dados"): os.remove("db_notas.csv"); st.rerun()
