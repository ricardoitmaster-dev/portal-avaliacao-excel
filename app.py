import streamlit as st
import pandas as pd
import random
import io
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

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

st.markdown(f"""
    <style>
        .stApp {{ background-color: {CORE_FUNDO} !important; }}
        h1 {{ color: {CORE_SENAI} !important; font-weight: bold; text-align: center !important; }}
        label, p, span {{ color: {CORE_TEXTO_BRANCO} !important; }}
        .stButton>button {{ color: white; background-color: #262730; border-radius: 8px; width: 100%; border: 1px solid {CORE_SENAI}; }}
        .stButton>button:hover {{ background-color: {CORE_SENAI}; }}
        .stDownloadButton>button {{ color: white !important; background-color: {CORE_SENAI} !important; font-weight: bold; width: 100%; }}
        .stTextInput input {{ background-color: #262730 !important; color: white !important; }}
    </style>
""", unsafe_allow_html=True)

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

def gerar_prova_excel(nome_aluno):
    itens = ["Notebook", "Mouse", "Teclado", "Monitor", "Impressora", "Cabo HDMI", "SSD 480GB"]
    dados = [{"ID": i, "Produto": random.choice(itens), "Quantidade": random.randint(5, 50), 
              "Preço Unitário": round(random.uniform(20, 300), 2), "Venda Total": 0, "Status": ""} for i in range(1, 31)]
    
    df = pd.DataFrame(dados)
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        # Aba 1: Dados para Processamento
        df.to_excel(writer, sheet_name='Base_de_Dados', index=False)
        
        # Aba 2: Contexto Inteligente e Instruções Pedagógicas
        inst = [
            ["AVALIAÇÃO PRÁTICA: GESTÃO DE ATIVOS E INDICADORES COMERCIAIS"],
            [""],
            ["CONTEXTO PROFISSIONAL:"],
            [f"Prezado(a) {nome_aluno}, você foi designado para automatizar o relatório de vendas de ativos de TI."],
            ["Sua missão é estruturar os dados para que a diretoria possa identificar a performance de cada item."],
            [""],
            ["DESAFIOS TÉCNICOS:"],
            ["1. CÁLCULO DE FATURAMENTO: Na coluna 'Venda Total', utilize operadores aritméticos para"],
            ["   determinar o montante total baseado no volume estocado e no valor unitário."],
            [""],
            ["2. ANÁLISE DE PERFORMANCE (LÓGICA CONDICIONAL): Na coluna 'Status', você deve criar uma inteligência"],
            ["   utilizando a função 'SE'. O critério estabelecido pela gerência é de 500 unidades monetárias."],
            ["   - Se atingir ou superar o critério, o status deve retornar 'META'."],
            ["   - Caso contrário, o sistema deve apontar a necessidade de 'REVISAR'."],
            [""],
            ["REGRAS DE INTEGRIDADE:"],
            ["- Não altere a estrutura das colunas ou os nomes das abas."],
            ["- O arquivo deve ser salvo e enviado sem alteração no nome original gerado pelo portal."],
            [""],
            ["Bom trabalho!"]
        ]
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
        nota = round(((pv / total) * 5) + ((ps / total) * 5), 1)
        return nota, f"Cálculos: {pv}/{total} | Lógica SE: {ps}/{total}"
    except:
        return 0, "Erro: Certifique-se de preencher a aba 'Base_de_Dados' corretamente."

# --- INTERFACE STREAMLIT ---
if 'etapa' not in st.session_state: st.session_state.etapa = 'login'

if st.session_state.etapa == 'login':
    st.image("https://upload.wikimedia.org/wikipedia/commons/8/8c/SENAI_S%C3%A3o_Paulo_logo.png", width=120)
    st.title("Portal de Avaliação Profissional")
    nome = st.text_input("Nome Completo do Aluno")
    turma = st.text_input("Identificação da Turma")
    email = st.text_input("E-mail Institucional/Pessoal")
    
    if st.button("Acessar Ambiente de Prova"):
        if nome and turma and email:
            st.session_state.aluno = {"nome": nome, "turma": turma, "email": email}
            st.session_state.nome_esperado = f"Avaliacao_{nome.replace(' ','_')}.xlsx"
            st.session_state.excel_data = gerar_prova_excel(nome)
            st.session_state.etapa = 'prova'
            st.rerun()
else:
    st.image("https://upload.wikimedia.org/wikipedia/commons/8/8c/SENAI_S%C3%A3o_Paulo_logo.png", width=100)
    st.title("Laboratório de Entrega")
    st.write(f"Candidato: **{st.session_state.aluno['nome']}**")
    
    st.download_button("📥 1. Baixar Caderno de Questões (Excel)", 
                       st.session_state.excel_data, 
                       st.session_state.nome_esperado)
    
    st.divider()
    arquivo_upload = st.file_uploader("2. Enviar Solução Finalizada", type=['xlsx'])
    
    if st.button("🚀 3. Submeter para Correção"):
        if arquivo_upload:
            if arquivo_upload.name != st.session_state.nome_esperado:
                st.error(f"SISTEMA DE SEGURANÇA: Nome do arquivo divergente. Envie o arquivo '{st.session_state.nome_esperado}'.")
            else:
                with st.spinner('Analisando fórmulas e lógica...'):
                    nota, feedback = calcular_nota(arquivo_upload)
                    corpo = f"Aluno: {st.session_state.aluno['nome']}\nTurma: {st.session_state.aluno['turma']}\nNota: {nota}\n{feedback}"
                    
                    enviar_email(EMAIL_PROFESSOR, f"RESULTADO {nota}: {st.session_state.aluno['nome']}", corpo, arquivo_upload.getvalue(), arquivo_upload.name)
                    enviar_email(st.session_state.aluno['email'], "Confirmação de Entrega - SENAI", corpo)
                    
                    st.success(f"Submissão realizada com sucesso! Nota: {nota}")
                    st.info(feedback)
                    st.balloons()
        else:
            st.warning("Nenhum arquivo detectado para submissão.")

    if st.button("Encerrar Sessão"):
        st.session_state.clear()
        st.rerun()
