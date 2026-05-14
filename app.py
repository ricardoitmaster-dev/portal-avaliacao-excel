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

CORE_SENAI = "#FF0000"
CORE_FUNDO = "#0E1117" 
CORE_TEXTO_BRANCO = "#FFFFFF"

st.set_page_config(page_title="Portal de Avaliação Excel - SENAI", layout="centered")

# --- ESTILIZAÇÃO CSS AVANÇADA ---
st.markdown(f"""
    <style>
        .stApp {{ background-color: {CORE_FUNDO} !important; }}
        h1 {{ color: {CORE_SENAI} !important; font-weight: bold; text-align: center !important; }}
        label, p, span {{ color: {CORE_TEXTO_BRANCO} !important; }}
        .stButton>button {{ color: white; background-color: #262730; border-radius: 8px; width: 100%; border: 1px solid {CORE_SENAI}; }}
        .stButton>button:hover {{ background-color: {CORE_SENAI}; }}
        .stDownloadButton>button {{ color: white !important; background-color: {CORE_SENAI} !important; font-weight: bold; width: 100%; }}
        .stTextInput input {{ background-color: #262730 !important; color: white !important; }}
        [data-testid="stHorizontalBlock"] {{ align-items: center; }}
    </style>
""", unsafe_allow_html=True)

# --- FUNÇÕES CORE (PEDAGÓGICO E ENVIO) ---
def gerar_feedback_pedagogico():
    return """
--------------------------------------------------
🎓 GUIA DE CORREÇÃO E BOAS PRÁTICAS (SENAI)
--------------------------------------------------
1. CÁLCULO DE FATURAMENTO: Fórmula correta: =C2*D2
2. LÓGICA CONDICIONAL: Fórmula esperada: =SE(E2>=500;"META";"REVISAR")
3. FORMATAÇÃO: Use formato 'Contábil' e aplique bordas e negrito.
4. MACROS: Salve como .XLSM para manter a automação.
--------------------------------------------------
"""

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
    except: return False

def gerar_prova_excel(nome_aluno):
    itens = ["Notebook", "Mouse", "Teclado", "Monitor", "Impressora", "Cabo HDMI", "SSD 480GB"]
    dados = [{"ID": i, "Produto": random.choice(itens), "Quantidade": random.randint(5, 50), 
              "Preço Unitário": round(random.uniform(20, 300), 2), "Venda Total": 0, "Status": ""} for i in range(1, 31)]
    df = pd.DataFrame(dados)
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, sheet_name='Base_de_Dados', index=False)
        inst = [["AVALIAÇÃO EXCEL - SENAI"], [f"Candidato: {nome_aluno}"], ["1. Calcule Venda Total."], ["2. Lógica SE (500)."], ["3. Macros de Ordenação."]]
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
        return nota, f"Cálculos: {pv}/{total} | Lógica SE: {ps}/{total} | Macro: {'Sim' if tem_macro>0 else 'Não'}"
    except: return 0, "Erro na leitura."

# --- INTERFACE ALUNO ---
if 'etapa' not in st.session_state: st.session_state.etapa = 'login'

col_logo, col_espaco, col_assinatura = st.columns([1, 1, 1])
with col_logo: st.image("https://upload.wikimedia.org/wikipedia/commons/8/8c/SENAI_S%C3%A3o_Paulo_logo.png", width=120)
with col_assinatura:
    st.markdown('<div style="display: flex; justify-content: flex-end;">', unsafe_allow_html=True)
    st.image("Imagem para o app avaliação Excel_RicardoItmaster.png", width=220)
    st.markdown('</div>', unsafe_allow_html=True)

if st.session_state.etapa == 'login':
    st.title("Portal de Avaliação Profissional")
    nome = st.text_input("Nome Completo")
    turma = st.text_input("Turma").strip().upper()
    email = st.text_input("E-mail")
    if st.button("Acessar Prova"):
        if nome and turma and email:
            st.session_state.aluno = {"nome": nome, "turma": turma, "email": email}
            st.session_state.nome_esp = f"Avaliacao_{nome.replace(' ','_')}.xlsx"
            st.session_state.excel_data = gerar_prova_excel(nome)
            st.session_state.etapa = 'prova'
            st.rerun()
else:
    st.title("Ambiente de Entrega")
    st.write(f"Aluno: **{st.session_state.aluno['nome']}**")
    st.download_button("📥 Baixar Prova", st.session_state.excel_data, st.session_state.nome_esp)
    st.divider()
    up = st.file_uploader("Enviar Arquivo", type=['xlsx', 'xlsm'])
    if st.button("🚀 Submeter"):
        if up:
            nota, feedback = calcular_nota(up)
            novo_reg = pd.DataFrame([[st.session_state.aluno['nome'], st.session_state.aluno['turma'], nota]], columns=['Aluno', 'Turma', 'Nota'])
            novo_reg.to_csv("db_notas.csv", mode='a', header=not os.path.exists("db_notas.csv"), index=False)
            st.success(f"Nota: {nota}")
            st.balloons()
    if st.button("Sair"):
        st.session_state.clear()
        st.rerun()

# --- PAINEL ADMINISTRATIVO COM LIMPEZA AUTOMÁTICA ---
st.divider()
with st.expander("👤 ÁREA RESTRITA (Professor / ADM)"):
    aba_adm = st.tabs(["Acesso Professor", "Novo Cadastro", "Gerente Geral (Ricardo)"])

    # 1. LOGIN PROFESSOR
    with aba_adm[0]:
        l_turma = st.text_input("Turma", key="l_t").strip().upper()
        l_pass = st.text_input("Senha", type="password", key="l_p")
        if l_pass:
            if os.path.exists("professores.csv"):
                df_p = pd.read_csv("professores.csv")
                if not df_p[(df_p['Turma'] == l_turma) & (df_p['Senha'] == str(l_pass))].empty:
                    st.success(f"Bem-vindo, Professor da Turma {l_turma}")
                    if os.path.exists("db_notas.csv"):
                        dados = pd.read_csv("db_notas.csv")
                        turma_dados = dados[dados['Turma'] == l_turma]
                        st.dataframe(turma_dados, use_container_width=True)
                    if st.button("Encerrar e Limpar Tela"): st.rerun()
                else: st.error("Dados incorretos.")

    # 2. CADASTRO
    with aba_adm[1]:
        c_prof = st.text_input("Nome Prof.")
        c_turma = st.text_input("Turma Prof.").strip().upper()
        c_pass = st.text_input("Definir Senha Prof.", type="password")
        if st.button("Confirmar Cadastro"):
            pd.DataFrame([[c_prof, c_turma, c_pass]], columns=['Professor','Turma','Senha']).to_csv("professores.csv", mode='a', header=not os.path.exists("professores.csv"), index=False)
            st.success("Cadastrado! Limpando dados...")
            st.rerun()

    # 3. ADM RICARDO (COM DASHBOARD COMPLETO)
    with aba_adm[2]:
        master = st.text_input("Senha Mestra ADM", type="password")
        if master == "ricardoitmaster":
            st.subheader("📊 DASHBOARD GERENCIAL (Todas as Turmas)")
            if os.path.exists("db_notas.csv"):
                full_db = pd.read_csv("db_notas.csv")
                
                # Métricas do Dashboard
                c1, c2, c3 = st.columns(3)
                c1.metric("Total Alunos", len(full_db))
                c2.metric("Média Geral", round(full_db['Nota'].mean(), 1))
                c3.metric("Turmas Ativas", full_db['Turma'].nunique())
                
                # Gráfico
                st.write("### Desempenho por Nota")
                st.bar_chart(full_db['Nota'].value_counts())
                
                # Dados
                st.write("### Tabela Completa")
                st.dataframe(full_db, use_container_width=True)
                
                if st.button("🔐 Sair do Modo ADM"): st.rerun()
            else: st.info("Nenhum dado capturado ainda.")
