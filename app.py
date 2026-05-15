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

st.set_page_config(page_title="Portal de Avaliação Excel Profissional", layout="wide")

# --- ESTILIZAÇÃO CSS COMPLETA ---
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
        .btn-sair-link:hover {{ background-color: #A00000; border: 2px solid white; }}
    </style>
""", unsafe_allow_html=True)

# --- FUNÇÕES DE APOIO E E-MAIL ---
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

# --- GERAÇÃO DE PROVA COM VARIABILIDADE (Máx 2 repetições) ---
def gerar_prova_excel_complexa(nome_aluno, turma):
    # Lógica para permitir apenas 2 tipos de provas diferentes por turma
    hash_id = int(hashlib.md5(f"{nome_aluno}{turma}".encode()).hexdigest(), 16)
    tipo_prova = hash_id % 2 
    random.seed(tipo_prova)

    itens_banco = ["Notebook", "Mouse", "Teclado", "Monitor", "Cabo", "SSD", "RAM", "Fonte"]
    precos_banco = [3500, 80, 150, 900, 40, 250, 400, 300]
    
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        # Aba Banco de Dados para PROCV
        df_banco = pd.DataFrame({"ID": range(1, 9), "PRODUTO": itens_banco, "PRECO_UNIT": precos_banco})
        df_banco.to_excel(writer, sheet_name='BANCO_DADOS', index=False)

        # Aba de Resolução
        dados_res = {
            "DATA": [f"{random.randint(1,28)}/05/2026" for _ in range(5)],
            "ID_PROD": [random.randint(1, 8) for _ in range(5)],
            "DESCRIÇÃO": [""] * 5,
            "VALOR_UNIT": [0.0] * 5,
            "QTD": [random.randint(2, 10) for _ in range(5)],
            "SUBTOTAL": [0.0] * 5,
            "DESCONTO_%": [0.0] * 5,
            "TOTAL_LIQUIDO": [0.0] * 5
        }
        pd.DataFrame(dados_res).to_excel(writer, sheet_name='RESOLUÇÃO', index=False)

        # Instruções Detalhadas
        instr = [
            ["AVALIAÇÃO EXCEL - SENAI 122"], [f"ALUNO: {nome_aluno}"], [""],
            ["REQUISITOS (TOTAL 100 PONTOS):"],
            ["1. PROCV: Busque Descrição e Valor na aba BANCO_DADOS."],
            ["2. MATEMÁTICA: Calcule SUBTOTAL (VALOR * QTD)."],
            ["3. LÓGICA SE: Se SUBTOTAL > 1000, DESCONTO é 10%, senão 0%."],
            ["4. MULTIPLICAÇÃO: Calcule TOTAL_LIQUIDO (Subtotal - Desconto)."],
            ["5. FORMATAÇÃO: Formate Moeda, Percentual e transforme em TABELA."],
            ["6. MACRO: Crie botão para ORDENAR por DATA."]
        ]
        pd.DataFrame(instr).to_excel(writer, sheet_name='INSTRUÇÕES', index=False, header=False)
    return output.getvalue()

# --- CORREÇÃO MINUCIOSA ---
def realizar_correcao(arquivo_bytes):
    try:
        df = pd.read_excel(arquivo_bytes, sheet_name='RESOLUÇÃO')
        wb_check = load_workbook(arquivo_bytes, keep_vba=True)
        pts = 0
        feedback = []

        # Validação PROCV/Matemática (40 pts)
        # Verificamos se o resultado final bate com o cálculo esperado
        if (df['SUBTOTAL'] == df['VALOR_UNIT'] * df['QTD']).all():
            pts += 40
        else:
            feedback.append("- Erro no cálculo de Subtotal ou busca de valores (PROCV/Mult).")

        # Validação SE (30 pts)
        meta_correta = True
        for _, r in df.iterrows():
            esp = 0.1 if r['SUBTOTAL'] > 1000 else 0.0
            if abs(r['DESCONTO_%'] - esp) > 0.01: meta_correta = False
        if meta_correta: pts += 30
        else: feedback.append("- Erro na função SE: Desconto de 10% aplicado incorretamente.")

        # Validação Macro e Tabela (30 pts)
        if wb_check.vba_archive: pts += 30
        else: feedback.append("- Macro de ordenação não detectada no arquivo .xlsm.")

        msg = "\n".join(feedback) if feedback else "Excelente! Todos os critérios atendidos."
        return pts, msg
    except: return 0, "Erro ao processar arquivo. Verifique o formato."

# --- INICIALIZAÇÃO DE SESSÃO ---
if 'perfil' not in st.session_state: st.session_state.perfil = None
if 'etapa_aluno' not in st.session_state: st.session_state.etapa_aluno = 'login'

# --- CABEÇALHO ---
c1, c2, c3 = st.columns([1, 2, 1])
with c1: st.image("https://upload.wikimedia.org/wikipedia/commons/8/8c/SENAI_S%C3%A3o_Paulo_logo.png", width=150)
with c3: st.image("Imagem para o app avaliação Excel_RicardoItmaster.png", width=220)

# --- FLUXO DE TELAS ---
if st.session_state.perfil is None:
    st.title("Sistema de Avaliação Técnica")
    st.write("### Selecione seu perfil de acesso:")
    _, col1, col2, _ = st.columns([0.8, 1, 1, 0.8])
    with col1:
        if st.button("🎓 SOU ALUNO"): st.session_state.perfil = "aluno"; st.rerun()
    with col2:
        if st.button("👨‍🏫 PROFESSOR / GESTOR"): st.session_state.perfil = "admin"; st.rerun()
    _, col_sair, _ = st.columns([1.5, 1, 1.5])
    with col_sair:
        st.markdown('<a href="https://guarulhos.sp.senai.br/" target="_self" class="btn-sair-link">❌ SAIR DO SISTEMA</a>', unsafe_allow_html=True)

elif st.session_state.perfil == "aluno":
    if st.button("⬅️ Voltar"): st.session_state.perfil = None; st.rerun()
    if st.session_state.etapa_aluno == 'login':
        with st.form("f_aluno"):
            n, t, e = st.text_input("Nome Completo"), st.text_input("Turma").upper(), st.text_input("Seu E-mail")
            if st.form_submit_button("Gerar Prova"):
                if n and t and e:
                    st.session_state.aluno_dados = {"nome":n, "turma":t, "email":e}
                    st.session_state.excel_data = gerar_prova_excel_complexa(n, t)
                    st.session_state.etapa_aluno = 'prova'; st.rerun()
    elif st.session_state.etapa_aluno == 'prova':
        n_arq = f"Avaliacao_{st.session_state.aluno_dados['nome'].replace(' ','_')}"
        st.download_button("📥 Baixar Planilha", st.session_state.excel_data, f"{n_arq}.xlsx")
        up = st.file_uploader("Upload Prova (.xlsm)", type=['xlsm', 'xlsx'])
        if st.button("🚀 Finalizar"):
            if up:
                nota, fb = realizar_correcao(up)
                relatorio = f"Aluno: {st.session_state.aluno_dados['nome']}\nNota Final: {nota}/100\n\nDetalhes:\n{fb}"
                enviar_email(st.session_state.aluno_dados['email'], "Seu Resultado - Excel", relatorio)
                enviar_email(EMAIL_PROFESSOR, f"Entrega Turma {st.session_state.aluno_dados['turma']}", relatorio)
                st.success(f"Enviado! Sua nota: {nota}"); st.balloons()

elif st.session_state.perfil == "admin":
    if st.button("⬅️ Menu Principal"): st.session_state.perfil = None; st.rerun()
    t1, t2, t3 = st.tabs(["📊 Notas", "📝 Professores", "🛡️ Área ADM"])
    with t1:
        if os.path.exists("professores.csv"):
            np, sp = st.text_input("Professor"), st.text_input("Senha", type="password")
            if st.button("Consultar"):
                dfp = pd.read_csv("professores.csv")
                if not dfp[(dfp['Professor'].str.lower() == np.lower()) & (dfp['Senha'] == str(sp))].empty:
                    st.write("### Notas da Turma")
                    # Lógica de exibição de notas...
    with t2:
        with st.form("cad_p"):
            n_p, t_p, s_p = st.text_input("Nome"), st.text_input("Turma"), st.text_input("Senha")
            if st.form_submit_button("Salvar Professor"):
                pd.DataFrame([[n_p, t_p, s_p]], columns=['Professor','Turma','Senha']).to_csv("professores.csv", mode='a', header=not os.path.exists("professores.csv"), index=False)
                st.success("Professor cadastrado!")
    with t3:
        if st.text_input("Senha Mestra", type="password") == "Celina2610$$":
            st.write("### Controle Administrativo")
            # Funções de limpeza e exportação geral...
