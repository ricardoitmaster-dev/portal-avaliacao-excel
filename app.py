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
    SENHA_APP_GOOGLE = "ugjhusmwnbmgzspv" # Mantendo sua senha válida

# ==========================================
# CUSTOMIZAÇÃO DE IDENTIDADE VISUAL SENAI
# ==========================================
# Baseado no manual de marca do SENAI-SP (Vermelho: #FF0000, Cinza: #404040)
CORE_SENAI = "#FF0000"
CORE_TEXTO = "#404040"
CORE_FUNDO = "#FFFFFF"

# Configuração de Página (Logo e Título na aba)
st.set_page_config(
    page_title="Portal de Avaliação Excel - SENAI", 
    page_icon="https://upload.wikimedia.org/wikipedia/commons/8/8c/SENAI_S%C3%A3o_Paulo_logo.png",
    layout="centered"
)

# Estilização CSS para forçar cores SENAI
st.markdown(f"""
    <style>
        /* Título Principal */
        h1 {{
            color: {CORE_SENAI} !important;
            font-family: 'Helvetica Neue', Helvetica, Arial, sans-serif;
            font-weight: bold;
        }}
        /* Subtítulos */
        h2, h3 {{
            color: {CORE_TEXTO} !important;
        }}
        /* Botões Primários (Enviar/Sair) */
        .stButton>button {{
            color: white;
            background-color: {CORE_TEXTO};
            border-radius: 5px;
            border: none;
        }}
        .stButton>button:hover {{
            background-color: {CORE_SENAI};
            color: white;
        }}
        /* Botão de Download (Especial) */
        .stDownloadButton>button {{
            color: white !important;
            background-color: {CORE_SENAI} !important;
            border-radius: 5px;
            font-weight: bold;
            border: none;
        }}
        /* Traço divisor */
        hr {{
            border-top: 2px solid {CORE_SENAI};
        }}
    </style>
""", unsafe_allow_html=True)

# --- FUNÇÕES DE MOTOR (Sem Alteração) ---
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
    except Exception as e:
        st.error(f"Erro técnico no envio: {e}")
        return False

def gerar_prova_excel(nome_aluno, turma):
    temas = [
        {"nome": "Vendas de Tecnologia", "cat": "Hardware", "itens": ["Notebook", "Mouse", "Teclado"]},
        {"nome": "Gestão de PetShop", "cat": "Serviços", "itens": ["Ração", "Banho", "Tosa"]},
        {"nome": "Logística de Alimentos", "cat": "Perecíveis", "itens": ["Arroz", "Feijão", "Azeite"]}
    ]
    tema = random.choice(temas)
    dados = []
    for i in range(1, 31):
        dados.append({
            "ID": i, 
            "Data": (datetime.now() - timedelta(days=random.randint(1, 365))).strftime("%d/%m/%Y"),
            "Produto": random.choice(tema["itens"]), 
            "Categoria": tema["cat"],
            "Quantidade": random.randint(5, 100), 
            "Preço Unitário": round(random.uniform(15, 600), 2), 
            "Venda Total": 0 
        })
    df = pd.DataFrame(dados)
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, sheet_name='Base_de_Dados', index=False)
        instrucoes = [
            ["AVALIAÇÃO PRÁTICA DE EXCEL"], [f"ALUNO: {nome_aluno.upper()}"], [f"TURMA: {turma}"],
            [""], ["1. Calcule a Venda Total (Quantidade x Preço)"], ["2. Crie Status com a função SE"], ["3. Faça uma Tabela Dinâmica em nova aba"]
        ]
        pd.DataFrame(instrucoes).to_excel(writer, sheet_name='Instruções', index=False, header=False)
    return output.getvalue()

# ==========================================
# INTERFACE DO ALUNO (Visual SENAI)
# ==========================================
if 'etapa' not in st.session_state:
    st.session_state.etapa = 'login'

if st.session_state.etapa == 'login':
    # Cabeçalho com o Logo Oficial
    col_logo, col_titulo = st.columns([1, 4])
    with col_logo:
        st.image("https://upload.wikimedia.org/wikipedia/commons/8/8c/SENAI_S%C3%A3o_Paulo_logo.png", width=120)
    with col_titulo:
        st.title("Portal de Avaliação Prática")
        st.subheader("Cursos de Tecnologia da Informação")
        
    st.write("Identifique-se para baixar sua prova personalizada de Excel.")
    nome = st.text_input("Nome Completo")
    turma = st.text_input("Sua Turma")
    email = st.text_input("Seu E-mail Corporativo ou Pessoal")
    
    if st.button("Acessar Avaliação"):
        if nome and turma and email:
            with st.spinner('Gerando prova...'):
                excel_data = gerar_prova_excel(nome, turma)
                nome_limpo = nome.replace(' ', '_')
                turma_limpa = turma.replace(' ', '_')
                st.session_state.nome_arquivo = f"Prova_Excel_{nome_limpo}_{turma_limpa}.xlsx"
                st.session_state.excel_data = excel_data
                st.session_state.aluno = {"nome": nome, "turma": turma, "email": email}
                st.session_state.etapa = 'prova'
                st.rerun()
        else:
            st.error("⚠️ Por favor, preencha todos os campos para continuar.")

else:
    # Cabeçalho da Área do Aluno
    col_logo, col_titulo = st.columns([1, 4])
    with col_logo:
        st.image("https://upload.wikimedia.org/wikipedia/commons/8/8c/SENAI_S%C3%A3o_Paulo_logo.png", width=120)
    with col_titulo:
        st.title("Área de Entrega")

    st.header(f"Bem-vindo, {st.session_state.aluno['nome']}")
    st.write(f"Turma: {st.session_state.aluno['turma']}")
    
    st.warning(f"⚠️ **Importante:** Ao salvar o seu trabalho no computador, use o nome exato: **{st.session_state.nome_arquivo}**")
    
    # Botão Vermelho SENAI para download
    st.download_button("📥 Baixar minha Prova Individual", st.session_state.excel_data, st.session_state.nome_arquivo)
    
    st.divider()
    
    st.subheader("Finalizou a prova? Envie o arquivo aqui:")
    arquivo_upload = st.file_uploader("Anexe seu arquivo .xlsx resolvido", type=['xlsx'])
    
    col1, col2 = st.columns(2)
    
    with col1:
        if st.button("🚀 Enviar Prova Final"):
            if arquivo_upload:
                if arquivo_upload.name != st.session_state.nome_arquivo:
                    st.error(f"❌ **NOME DO ARQUIVO INVÁLIDO!** O sistema só aceita arquivos com o nome: **{st.session_state.nome_arquivo}**")
                else:
                    with st.spinner('Processando envio...'):
                        conteudo_arquivo = arquivo_upload.getvalue()
                        nome_do_arquivo = arquivo_upload.name
                        
                        # Envio para o Professor
                        corpo_prof = f"O aluno {st.session_state.aluno['nome']} (Turma {st.session_state.aluno['turma']}) entregou a prova."
                        p = enviar_email(EMAIL_PROFESSOR, f"PROVA RECEBIDA (SENAI): {st.session_state.aluno['nome']}", corpo_prof, conteudo_arquivo, nome_do_arquivo)
                        
                        # Envio de comprovante com anexo para o Aluno
                        corpo_aluno = f"Olá {st.session_state.aluno['nome']}!\n\nConfirmamos o recebimento da sua prova prática de Excel.\nSegue em anexo o arquivo que você enviou como comprovante de entrega.\n\nAtenciosamente,\nProf. Ricardo - SENAI"
                        enviar_email(st.session_state.aluno['email'], "Comprovante de Entrega - Prova Excel (SENAI)", corpo_aluno, conteudo_arquivo, nome_do_arquivo)
                        
                        if p:
                            st.balloons()
                            st.success("✅ Prova enviada com sucesso! Você e o professor receberam o arquivo por e-mail.")
            else:
                st.error("⚠️ Anexe o arquivo antes de tentar enviar.")

    with col2:
        if st.button("🚪 Sair do Portal"):
            st.session_state.clear()
            st.rerun()
