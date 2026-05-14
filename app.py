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

# --- CONFIGURAÇÕES ---
try:
    EMAIL_PROFESSOR = st.secrets["EMAIL_PROFESSOR"]
    SENHA_APP_GOOGLE = st.secrets["SENHA_APP_GOOGLE"]
except:
    EMAIL_PROFESSOR = "ricardoitmaster@gmail.com"
    SENHA_APP_GOOGLE = "ugjhusmwnbmgzspv"

CORE_SENAI = "#FF0000"
CORE_FUNDO = "#0E1117" 

st.set_page_config(page_title="Portal de Avaliação Excel - SENAI", layout="centered")

# --- CSS ---
st.markdown(f"""
    <style>
        .stApp {{ background-color: {CORE_FUNDO} !important; }}
        h1 {{ color: {CORE_SENAI} !important; font-weight: bold; text-align: center; }}
        label, p, span {{ color: #FFFFFF !important; }}
        .stButton>button {{ color: white; background-color: #262730; border-radius: 8px; width: 100%; border: 1px solid {CORE_SENAI}; }}
    </style>
""", unsafe_allow_html=True)

def encerrar_sessao():
    for key in list(st.session_state.keys()):
        del st.session_state[key]
    st.rerun()

def gerar_desafio_excel(nome_aluno):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        itens = ["Monitor LED", "Teclado Mecânico", "Mouse Óptico", "SSD 1TB", "Placa de Vídeo", "Fonte 600W"]
        dados = [{"ID": i, "Produto": random.choice(itens), "Quantidade": random.randint(5, 50), "Preço Unit": round(random.uniform(100, 1500), 2), "Subtotal": 0, "IPI (10%)": 0, "Total Geral": 0, "Status": "", "Cod_Ref": random.randint(101, 106)} for i in range(1, 21)]
        pd.DataFrame(dados).to_excel(writer, sheet_name='DESAFIO_PRATICO', index=False)
        pd.DataFrame({"Cod": [101, 102, 103, 104, 105, 106], "Categoria": ["Periféricos", "Armazenamento", "Hardware", "Energia", "Vídeo", "Redes"]}).to_excel(writer, sheet_name='REFERENCIA', index=False)
        
        inst = [
            ["CANDIDATO:", nome_aluno.upper()],
            ["REQUISITOS DA PROVA:"],
            ["1. TABELA:", "Converta os dados em Objeto Tabela."],
            ["2. CÁLCULOS:", "Subtotal, IPI e Total Geral."],
            ["3. LÓGICA:", "Função SE para Status."],
            ["4. BUSCA:", "PROCV na aba REFERENCIA."],
            ["5. MACROS:", "Crie macros de ordenação e botões."],
            ["SALVE COMO:", f"Prova_{nome_aluno.replace(' ','_')}"]
        ]
        pd.DataFrame(inst).to_excel(writer, sheet_name='INSTRUCOES', index=False, header=False)
    return output.getvalue()

def realizar_correcao(lista_arquivos):
    melhor_nota = 0
    relatorio = ""
    for arq in lista_arquivos:
        try:
            df = pd.read_excel(arq, sheet_name='DESAFIO_PRATICO')
            wb = load_workbook(arq, keep_vba=True)
            p = 0
            res = [f"Análise: {arq.name}"]
            if all(round(r['Subtotal'],1) == round(r['Quantidade']*r['Preço Unit'],1) for _,r in df.iterrows()): p += 25
            if 'Status' in df.columns and not df['Status'].isnull().all(): p += 25
            if 'Categoria' in df.columns: p += 25
            if arq.name.endswith('.xlsm') and wb.vba_archive: p += 25
            if p >= melhor_nota:
                melhor_nota = p
                relatorio = "\n".join(res) + f"\nNota: {p}/100"
        except: continue
    return melhor_nota, relatorio

def enviar_emails(aluno_email, nome_aluno, nota, relato, arquivos):
    corpo = f"Resultado SENAI\nAluno: {nome_aluno}\nNota: {nota}/100\n\n{relato}"
    try:
        for dest in [EMAIL_PROFESSOR, aluno_email]:
            msg = MIMEMultipart()
            msg['From'] = EMAIL_PROFESSOR
            msg['To'] = dest
            msg['Subject'] = f"Prova Excel - {nome_aluno}"
            msg.attach(MIMEText(corpo, 'plain'))
            for arq in arquivos:
                part = MIMEBase('application', 'octet-stream')
                part.set_payload(arq.getvalue())
                encoders.encode_base64(part)
                part.add_header('Content-Disposition', f"attachment; filename={arq.name}")
                msg.attach(part)
            server = smtplib.SMTP('smtp.gmail.com', 587)
            server.starttls()
            server.login(EMAIL_PROFESSOR, SENHA_APP_GOOGLE)
            server.send_message(msg)
            server.quit()
        return True
    except: return False

# --- INTERFACE ---
if 'etapa' not in st.session_state: st.session_state.etapa = 'login'

if st.session_state.etapa == 'login':
    st.title("Portal de Avaliação SENAI")
    n = st.text_input("Nome Completo")
    t = st.text_input("Turma").strip().upper()
    e = st.text_input("E-mail")
    if st.button("Acessar Prova"):
        if n and t and e:
            st.session_state.aluno = {"nome": n, "turma": t, "email": e}
            st.session_state.prova_origem = gerar_desafio_excel(n)
            st.session_state.nome_limpo = n.replace(" ","_")
            st.session_state.etapa = 'prova'
            st.rerun()

elif st.session_state.etapa == 'prova':
    st.title("Avaliação Prática")
    st.write(f"Candidato: **{st.session_state.aluno['nome']}**")
    
    st.download_button("📥 1. Baixar Arquivo de Prova", st.session_state.prova_origem, f"Prova_{st.session_state.nome_limpo}.xlsx")
    
    st.divider()
    st.subheader("2. Entregar Resolução")
    ups = st.file_uploader("Selecione seus arquivos (.xlsx ou .xlsm)", type=['xlsx', 'xlsm'], accept_multiple_files=True)
    
    # --- VALIDAÇÃO DE SEGURANÇA EM TEMPO REAL ---
    arquivos_validos = []
    if ups:
        for arq in ups:
            if st.session_state.nome_limpo.lower() in arq.name.lower():
                arquivos_validos.append(arq)
            else:
                st.error(f"⚠️ ARQUIVO INVÁLIDO: O arquivo '{arq.name}' não contém seu nome. Por favor, renomeie-o corretamente antes de enviar.")

    # Só habilita o botão se houver arquivos e todos forem válidos
    if st.button("🚀 Finalizar e Enviar para Correção"):
        if len(arquivos_validos) == len(ups) and ups:
            nota, relato = realizar_correcao(ups)
            sucesso = enviar_emails(st.session_state.aluno['email'], st.session_state.aluno['nome'], nota, relato, ups)
            if sucesso:
                st.success(f"Nota: {nota}/100. Relatório enviado para você e para o professor!")
                st.text(relato)
        elif not ups:
            st.warning("Por favor, faça o upload de pelo menos um arquivo.")
        else:
            st.error("Corrija os nomes dos arquivos acima para habilitar o envio.")

    st.divider()
    if st.button("🛑 Encerrar Sessão"):
        encerrar_sessao()
