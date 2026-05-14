import streamlit as st
import pandas as pd
import random
import io
import smtplib
import base64
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

# --- ESTILIZAÇÃO CSS AVANÇADA (Sua Versão Original Mantida) ---
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

# --- FUNÇÃO DE FEEDBACK DIDÁTICO (Original 227 Linhas) ---
def gerar_feedback_pedagogico():
    return """
--------------------------------------------------
🎓 GUIA DE CORREÇÃO E BOAS PRÁTICAS (SENAI)
--------------------------------------------------
1. CÁLCULO DE FATURAMENTO:
   - A fórmula correta para a 'Venda Total' é: =C2*D2
   
2. LÓGICA CONDICIONAL (Função SE):
   - A fórmula esperada é: =SE(E2>=500;"META";"REVISAR")
   
3. FORMATAÇÃO E APRESENTAÇÃO PROFISSIONAL:
   - MOEDA: Formate valores financeiros como 'Contábil' (R$).
   - ESTÉTICA: Use bordas, negrito nos cabeçalhos e cores sóbrias.
   - ALINHAMENTO: Centralize IDs e Quantidades para melhor leitura.
   - MACROS: O arquivo deve ser salvo como .XLSM para manter a automação.
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
        inst = [
            ["AVALIAÇÃO PRÁTICA: GESTÃO DE ATIVOS E INDICADORES COMERCIAIS"],
            [""],
            ["CONTEXTO PROFISSIONAL:"],
            [f"Prezado(a) {nome_aluno}, você foi designado para automatizar o relatório de vendas."],
            ["Sua missão é estruturar os dados para que a diretoria possa identificar a performance."],
            [""],
            ["DESAFIOS TÉCNICOS:"],
            ["1. CÁLCULO DE FATURAMENTO: Na coluna 'Venda Total', utilize operadores aritméticos."],
            ["   determine o montante total baseado no volume estocado e no valor unitário."],
            [""],
            ["2. ANÁLISE DE PERFORMANCE (LÓGICA CONDICIONAL): Na coluna 'Status', use a função 'SE'."],
            ["   - Se atingir ou superar 500, o status deve retornar 'META'."],
            ["   - Caso contrário, o sistema deve apontar a necessidade de 'REVISAR'."],
            [""],
            ["3. AUTOMAÇÃO (MACROS): Para facilitar a operação, a diretoria exige a criação de macros:"],
            ["   - Crie uma Macro para ORDENAR a tabela pelo campo 'Produto' de A-Z."],
            ["   - Crie uma Macro para ORDENAR a tabela pelo campo 'Venda Total' do maior para o menor."],
            ["   - Insira BOTÕES na planilha e atribua as macros correspondentes a eles."],
            [""],
            ["REGRAS DE INTEGRIDADE:"],
            ["- Não altere a estrutura das colunas ou os nomes das abas."],
            ["- IMPORTANTE: Para que as macros funcionem, salve o arquivo como .xlsm."],
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
        try:
            wb = load_workbook(arquivo_bytes, keep_vba=True)
            tem_macro = 2.0 if wb.vba_archive else 0.0
        except: tem_macro = 0.0
        nota = round(((pv / total) * 4) + ((ps / total) * 4) + tem_macro, 1)
        feedback_macro = "Macro detectada (+2.0)" if tem_macro > 0 else "Nenhuma Macro detectada (0.0)"
        return nota, f"Cálculos: {pv}/{total} | Lógica SE: {ps}/{total} | {feedback_macro}"
    except: return 0, "Erro: Certifique-se de preencher a aba 'Base_de_Dados' corretamente."

# --- LOGICA DE INTERFACE ---
if 'etapa' not in st.session_state: st.session_state.etapa = 'login'

col_logo, col_espaco, col_assinatura = st.columns([1, 1, 1])
with col_logo: st.image("https://upload.wikimedia.org/wikipedia/commons/8/8c/SENAI_S%C3%A3o_Paulo_logo.png", width=120)
with col_assinatura:
    st.markdown('<div style="display: flex; justify-content: flex-end;">', unsafe_allow_html=True)
    st.image("Imagem para o app avaliação Excel_RicardoItmaster.png", width=220)
    st.markdown('</div>', unsafe_allow_html=True)

if st.session_state.etapa == 'login':
    st.title("Portal de Avaliação Profissional")
    nome = st.text_input("Nome Completo do Aluno")
    turma = st.text_input("Identificação da Turma").strip().upper()
    email = st.text_input("E-mail Institucional/Pessoal")
    if st.button("Acessar Ambiente de Prova"):
        if nome and turma and email:
            st.session_state.aluno = {"nome": nome, "turma": turma, "email": email}
            st.session_state.nome_esp = f"Avaliacao_{nome.replace(' ','_')}.xlsx"
            st.session_state.excel_data = gerar_prova_excel(nome)
            st.session_state.etapa = 'prova'
            st.rerun()
else:
    st.title("Laboratório de Entrega")
    st.write(f"Candidato: **{st.session_state.aluno['nome']}** | Turma: **{st.session_state.aluno['turma']}**")
    st.download_button("📥 1. Baixar Caderno de Questões", st.session_state.excel_data, st.session_state.nome_esp)
    st.divider()
    up_file = st.file_uploader("2. Enviar Solução (xlsx ou xlsm)", type=['xlsx', 'xlsm'])
    if st.button("🚀 3. Submeter para Correção"):
        if up_file:
            if up_file.name.split('.')[0] != st.session_state.nome_esp.split('.')[0]:
                st.error("SISTEMA DE SEGURANÇA: Nome do arquivo divergente.")
            else:
                nota, feedback = calcular_nota(up_file)
                tutorial = gerar_feedback_pedagogico()
                # SALVAR DADOS
                pd.DataFrame([[st.session_state.aluno['nome'], st.session_state.aluno['turma'], nota]], columns=['Aluno', 'Turma', 'Nota']).to_csv("db_notas.csv", mode='a', header=not os.path.exists("db_notas.csv"), index=False)
                
                corpo = f"Aluno: {st.session_state.aluno['nome']}\nNota: {nota}\n{feedback}\n\n{tutorial}"
                enviar_email(EMAIL_PROFESSOR, f"RESULTADO {nota}: {st.session_state.aluno['nome']}", corpo, up_file.getvalue(), up_file.name)
                enviar_email(st.session_state.aluno['email'], "Gabarito SENAI", corpo)
                st.success(f"Submissão realizada! Nota: {nota}")
                st.balloons()
    if st.button("Sair/Trocar Aluno"):
        st.session_state.clear()
        st.rerun()

# --- PAINEL DE GESTÃO (INTEGRADO E SEGURO) ---
st.divider()
with st.expander("👤 ÁREA DO PROFESSOR & GERÊNCIA"):
    tabs = st.tabs(["Acesso Professor", "Novo Cadastro Professor", "Painel ADM (Ricardo)"])
    
    with tabs[1]: # Cadastro
        st.write("#### Registrar Docente")
        reg_n = st.text_input("Nome Professor")
        reg_t = st.text_input("Turma Autorizada", key="reg_t").strip().upper()
        reg_s = st.text_input("Criar Senha", type="password", key="reg_s")
        if st.button("Salvar Cadastro"):
            if reg_n and reg_t and reg_s:
                pd.DataFrame([[reg_n, reg_t, reg_s]], columns=['Prof', 'Turma', 'Senha']).to_csv("professores.csv", mode='a', header=not os.path.exists("professores.csv"), index=False)
                st.success("Professor cadastrado! Limpando campos...")
                st.rerun()

    with tabs[0]: # Login Professor
        st.write("#### Dashboard da Turma")
        log_t = st.text_input("Turma Cadastrada", key="log_t").strip().upper()
        log_s = st.text_input("Sua Senha", type="password", key="log_s")
        if log_s:
            if os.path.exists("professores.csv"):
                df_p = pd.read_csv("professores.csv")
                if not df_p[(df_p['Turma'] == log_t) & (df_p['Senha'] == str(log_s))].empty:
                    if os.path.exists("db_notas.csv"):
                        db_n = pd.read_csv("db_notas.csv")
                        turma_res = db_n[db_n['Turma'] == log_t]
                        st.subheader(f"📊 Resultados Turma {log_t}")
                        st.dataframe(turma_res, use_container_width=True)
                    if st.button("🔐 Fechar Sessão Professor"): st.rerun()
                else: st.error("Acesso Negado.")

    with tabs[2]: # ADM Ricardo
        st.write("#### Controle Geral Manager")
        m_s = st.text_input("Senha Mestra", type="password", key="m_s")
        if m_s == "ricardoitmaster":
            if os.path.exists("db_notas.csv"):
                full_db = pd.read_csv("db_notas.csv")
                st.subheader("📊 DASHBOARD GERAL")
                c1, c2, c3 = st.columns(3)
                c1.metric("Total Alunos", len(full_db))
                c2.metric("Média Geral", round(full_db['Nota'].mean(), 1))
                c3.metric("Turmas Ativas", full_db['Turma'].nunique())
                st.write("### Desempenho Global")
                st.bar_chart(full_db['Nota'].value_counts())
                st.write("### Base de Dados Completa")
                st.dataframe(full_db, use_container_width=True)
                if st.button("🔐 Trancar Painel ADM"): st.rerun()
            else: st.info("Nenhum registro encontrado.")
