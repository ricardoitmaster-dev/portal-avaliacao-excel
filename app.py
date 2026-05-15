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
        .stApp {{ 
            background-color: {COR_PRETO_BRILHANTE} !important; 
        }}
        h1, h2, h3 {{ 
            color: {COR_DOURADO} !important; 
            font-weight: bold; 
            text-align: center; 
        }}
        label, p, span {{ 
            color: {COR_TEXTO} !important; 
        }}
        .stButton>button {{ 
            color: {COR_TEXTO}; 
            background-color: {COR_AZUL_BMW}; 
            border-radius: 10px; 
            border: 1px solid {COR_DOURADO};
            height: 3.8em;
            font-weight: bold;
            width: 100%;
            margin-top: 10px;
        }}
        .stButton>button:hover {{ 
            border: 2px solid {COR_TEXTO}; 
            color: {COR_DOURADO}; 
        }}
        .stDownloadButton>button {{ 
            color: {COR_PRETO_BRILHANTE} !important; 
            background-color: {COR_DOURADO} !important; 
            font-weight: bold; 
            width: 100%; 
        }}
        .stTextInput input {{ 
            background-color: #1A1A1A !important; 
            color: white !important; 
            border: 1px solid {COR_AZUL_BMW} !important; 
        }}
        [data-testid="stHorizontalBlock"] {{ 
            align-items: center !important; 
        }}
        
        /* Botão Sair com estilo diferenciado */
        .btn-sair-container button {{
            background-color: #7B0000 !important;
            border: 1px solid #FFFFFF !important;
            margin-top: 20px;
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
    except:
        return False

def gerar_prova_excel(nome_aluno):
    itens = ["Notebook", "Mouse", "Teclado", "Monitor", "Impressora", "Cabo HDMI", "SSD 480GB"]
    dados = []
    for i in range(1, 31):
        item = random.choice(itens)
        qtd = random.randint(5, 50)
        preco = round(random.uniform(20, 300), 2)
        dados.append({
            "ID": i, 
            "Produto": item, 
            "Quantidade": qtd, 
            "Preço Unitário": preco, 
            "Venda Total": 0, 
            "Status": ""
        })
    
    df = pd.DataFrame(dados)
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, sheet_name='Base_de_Dados', index=False)
        instrucoes = [
            ["AVALIAÇÃO PRÁTICA DE EXCEL"],
            [f"Nome do Aluno: {nome_aluno}"],
            [""],
            ["INSTRUÇÕES:"],
            ["1. Calcule a coluna 'Venda Total' (Quantidade x Preço Unitário)"],
            ["2. Na coluna 'Status', use a função SE:"],
            ["   - Se Venda Total for >= 500, escreva 'META'"],
            ["   - Caso contrário, escreva 'REVISAR'"],
            ["3. Crie uma macro simples de formatação e salve o arquivo como .xlsm"]
        ]
        pd.DataFrame(instrucoes).to_excel(writer, sheet_name='Instrucoes', index=False, header=False)
    return output.getvalue()

def calcular_nota(arquivo_bytes):
    try:
        df = pd.read_excel(arquivo_bytes, sheet_name='Base_de_Dados', engine='openpyxl')
        acertos_venda = 0
        acertos_status = 0
        total_linhas = len(df)
        
        for _, row in df.iterrows():
            calc_venda = round(float(row['Quantidade'] * row['Preço Unitário']), 2)
            if round(float(row['Venda Total']), 2) == calc_venda:
                acertos_venda += 1
            
            status_esperado = "META" if calc_venda >= 500 else "REVISAR"
            if str(row['Status']).strip().upper() == status_esperado:
                acertos_status += 1
        
        try:
            wb = load_workbook(arquivo_bytes, keep_vba=True)
            tem_macro = 2.0 if wb.vba_archive else 0.0
        except:
            tem_macro = 0.0
            
        nota_venda = (acertos_venda / total_linhas) * 4
        nota_status = (acertos_status / total_linhas) * 4
        nota_final = round(nota_venda + nota_status + tem_macro, 1)
        
        detalhes = f"Vendas: {acertos_venda}/{total_linhas} | Status: {acertos_status}/{total_linhas} | Macro: {tem_macro}"
        return nota_final, detalhes
    except Exception as e:
        return 0, f"Erro na correção: {str(e)}"

# --- INICIALIZAÇÃO DE ESTADO ---
if 'perfil' not in st.session_state:
    st.session_state.perfil = None
if 'etapa_aluno' not in st.session_state:
    st.session_state.etapa_aluno = 'login'

# --- CABEÇALHO ---
col_l, col_c, col_r = st.columns([1, 2, 1])
with col_l:
    st.image("https://upload.wikimedia.org/wikipedia/commons/8/8c/SENAI_S%C3%A3o_Paulo_logo.png", width=150)
with col_r:
    st.image("Imagem para o app avaliação Excel_RicardoItmaster.png", width=220)

# --- FLUXO DE TELAS ---

if st.session_state.perfil is None:
    st.markdown("<br><br>", unsafe_allow_html=True)
    st.title("Sistema de Avaliação Técnica")
    st.write("### Olá! Por favor, escolha uma opção para continuar:")
    
    # Centralização dos três botões principais
    _, col_btn, _ = st.columns([1.3, 1, 1.3])
    
    with col_btn:
        if st.button("🎓 ACESSO DO ALUNO"):
            st.session_state.perfil = "aluno"
            st.rerun()
            
        if st.button("👨‍🏫 PAINEL DO PROFESSOR"):
            st.session_state.perfil = "admin"
            st.rerun()
            
        st.markdown('<div class="btn-sair-container">', unsafe_allow_html=True)
        if st.button("❌ FECHAR / SAIR DO SISTEMA"):
            # Limpeza absoluta da sessão para garantir que o botão "faça algo"
            for key in list(st.session_state.keys()):
                del st.session_state[key]
            st.rerun()
        st.markdown('</div>', unsafe_allow_html=True)

elif st.session_state.perfil == "aluno":
    if st.button("⬅️ Voltar ao Menu Principal"):
        st.session_state.perfil = None
        st.rerun()

    if st.session_state.etapa_aluno == 'login':
        st.subheader("Identificação do Estudante")
        with st.form("form_identificacao"):
            nome = st.text_input("Nome Completo").strip()
            turma = st.text_input("Sua Turma").strip().upper()
            email = st.text_input("Seu E-mail").strip()
            
            if st.form_submit_button("Gerar Minha Prova"):
                if nome and turma and email:
                    st.session_state.aluno_dados = {"nome": nome, "turma": turma, "email": email}
                    st.session_state.excel_data = gerar_prova_excel(nome)
                    st.session_state.etapa_aluno = 'prova'
                    st.rerun()
                else:
                    st.warning("Preencha todos os campos corretamente.")

    elif st.session_state.etapa_aluno == 'prova':
        st.write(f"### Bem-vindo, {st.session_state.aluno_dados['nome']}!")
        nome_esperado = f"Avaliacao_{st.session_state.aluno_dados['nome'].replace(' ', '_')}"
        
        st.info(f"Baixe o arquivo abaixo, resolva os desafios e envie-o com o nome: **{nome_esperado}**")
        st.download_button("📥 Baixar Planilha de Prova", st.session_state.excel_data, f"{nome_esperado}.xlsx")
        
        st.divider()
        up_files = st.file_uploader("Envie sua prova respondida (.xlsx ou .xlsm)", type=['xlsx', 'xlsm'], accept_multiple_files=True)
        
        if st.button("🚀 Finalizar e Enviar"):
            if up_files:
                validos = [f for f in up_files if f.name.split('.')[0].lower
