import streamlit as st
import pandas as pd
import random
import io
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from datetime import datetime, timedelta

# --- CONFIGURAÇÕES VIA SECRETS ---
try:
    EMAIL_PROFESSOR = st.secrets["EMAIL_PROFESSOR"]
    SENHA_APP_GOOGLE = st.secrets["SENHA_APP_GOOGLE"]
except:
    EMAIL_PROFESSOR = "ricardoitmaster@gmail.com"
    SENHA_APP_GOOGLE = "ugjhusmwnbmgzspv"

# --- IDENTIDADE VISUAL ---
CORE_SENAI = "#FF0000"
CORE_TEXTO = "#404040"
CORE_GELO = "#F8F9FA"  # Cor de fundo padrão SENAI (clara)

st.set_page_config(
    page_title="Portal de Avaliação Excel - SENAI", 
    page_icon="https://upload.wikimedia.org/wikipedia/commons/8/8c/SENAI_S%C3%A3o_Paulo_logo.png",
    layout="centered"
)

# CSS para forçar Fundo Branco/Gelo e Centralizar Subtitle
st.markdown(f"""
    <style>
        /* Força fundo claro independente do modo do navegador */
        .stApp {{
            background-color: {CORE_GELO} !important;
        }}
        h1 {{ 
            color: {CORE_SENAI} !important; 
            font-weight: bold; 
            text-align: center !important;
            margin-bottom: 5px !important;
        }}
        /* Estilo para centralizar o curso */
        .centered-subtitle {{
            text-align: center !important;
            color: {CORE_TEXTO} !important;
            font-size: 1.2rem;
            font-weight: 500;
            margin-bottom: 30px;
        }}
        .stButton>button {{ 
            color: white; 
            background-color: {CORE_TEXTO}; 
            border-radius: 8px; 
            width: 100%;
        }}
        .stButton>button:hover {{ background-color: {CORE_SENAI}; color: white; }}
        .stDownloadButton>button {{ 
            color: white !important; 
            background-color: {CORE_SENAI} !important; 
            font-weight: bold; 
            width: 100%;
        }}
        hr {{ border-top: 2px solid {CORE_SENAI}; }}
    </style>
""", unsafe_allow_html=True)

# --- FUNÇÕES DE MOTOR E CORREÇÃO ---
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
    dados = []
    # VOLTANDO PARA 30 REGISTROS
    for i in range(1, 31):
        qtd = random.randint(5, 50)
        preco = round(random.uniform(20, 300), 2)
        dados.append({
            "ID": i, 
            "Produto": random.choice(itens),
            "Quantidade": qtd, 
            "Preço Unitário": preco,
            "Venda Total": 0, 
            "Status": ""
        })
    df = pd.DataFrame(dados)
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, sheet_name='Base_de_Dados', index=False)
        # INSTRUÇÕES PEDAGÓGICAS EXPLÍCITAS
        inst = [
            ["AVALIAÇÃO PRÁTICA DE EXCEL"], 
            [f"ALUNO: {nome_aluno}"], 
            [f"TURMA: {turma}"], 
            [""],
            ["INSTRUÇÕES TÉCNICAS:"],
            ["1. Na coluna Venda Total: Utilize a função de MULTIPLICAÇÃO ou a função MULT (Quantidade * Preço)."],
            ["2. Na coluna Status: Utilize a função lógica SE. Regra: Se Venda Total >= 500 escrever 'META', caso contrário 'REVISAR'."],
            ["3. Crie uma Tabela Dinâmica em uma nova aba resumindo o total de vendas por Produto."]
        ]
        pd.DataFrame(inst).to_excel(writer, sheet_name='Instruções', index=False, header=False)
    return output.getvalue()

def corrigir_prova(arquivo_aluno):
    try:
        df = pd.read_excel(arquivo_aluno, sheet_name='Base_de_Dados')
        pontos_venda = 0
        pontos_status = 0
        total_linhas = len(df)
        for index, row in df.iterrows():
            # Validação Matemática
            if round(row['Venda Total'], 2) == round(row['Quantidade'] * row['Preço Unitário'], 2):
                pontos_venda += 1
            # Validação Lógica SE
            esperado = "META" if row['Venda Total'] >= 500 else "REVISAR"
            if str(row['Status']).strip().upper() == esperado:
                pontos_status += 1
        
        nota = round(((pontos_venda / total_linhas) * 5) + ((pontos_status / total_linhas) * 5), 1)
        feedback = f"Resultados: {pontos_venda}/{total_linhas} cálculos de multiplicação e {pontos_status}/{total_linhas} funções SE corretas."
        return nota, feedback
    except:
        return 0, "Erro ao processar. Verifique se os nomes das abas e colunas não foram alterados."

# ==========================================
# INTERFACE DO USUÁRIO
# ==========================================
if 'etapa' not in st.session_state: st.session_state.etapa = 'login'

if st.session_state.etapa == 'login':
    # Centralização do Header
    st.image("https://upload.wikimedia.org/wikipedia/commons/8/8c/SENAI_S%C3%A3o_Paulo_logo.png", width=120)
    st.title("Portal de Avaliação Prática")
    # SUBHEADER CENTRALIZADO VIA HTML
    st.markdown('<p class="centered-subtitle">Cursos de Tecnologia da Informação - Excel Completo</p>', unsafe_allow_html=True)
    
    nome = st.text_input("Nome Completo")
    turma = st.text_input("Sua Turma")
    email = st.text_input("Seu E-mail")
    
    if st.button("Acessar Avaliação"):
        if nome and turma and email:
            st.session_state.excel_data = gerar_prova_excel(nome, turma)
            st.session_state.nome_arquivo = f"Prova_Excel_{nome.replace(' ','_')}.xlsx"
            st.session_state.aluno = {"nome": nome, "turma": turma, "email": email}
            st.session_state.etapa = 'prova'
            st.rerun()
        else:
            st.error("Preencha todos os campos.")

else:
    st.image("https://upload.wikimedia.org/wikipedia/commons/8/8c/SENAI_S%C3%A3o_Paulo_logo.png", width=100)
    st.title("Entrega e Correção")
    st.markdown(f'<p class="centered-subtitle">Aluno: {st.session_state.aluno["nome"]} | Turma: {st.session_state.aluno["turma"]}</p>', unsafe_allow_html=True)
    
    st.download_button("📥 Baixar minha Prova Individual", st.session_state.excel_data, st.session_state.nome_arquivo)
    st.divider()
    
    arquivo_upload = st.file_uploader("Anexe sua prova resolvida (.xlsx)", type=['xlsx'])
    
    if st.button("🚀 Enviar e Obter Resultado"):
        if arquivo_upload:
            with st.spinner('Aguarde, corrigindo sua prova...'):
                nota, feedback = corrigir_prova(arquivo_upload)
                corpo = f"O aluno {st.session_state.aluno['nome']} obteve nota {nota}.\n{feedback}"
                enviar_email(EMAIL_PROFESSOR, f"NOTA {nota}: {st.session_state.aluno['nome']}", corpo, arquivo_upload.getvalue(), arquivo_upload.name)
                enviar_email(st.session_state.aluno['email'], "Resultado Avaliação Excel SENAI", corpo)
                
                st.success(f"Sua nota final é: {nota}")
                st.info(feedback)
                st.balloons()
        else:
            st.error("Por favor, anexe o arquivo antes de enviar.")
            
    if st.button("🚪 Sair"):
        st.session_state.clear()
        st.rerun()

Seu portal agora está blindado contra o modo noturno e pedagogicamente mais robusto. Fico à disposição para qualquer outro ajuste!
