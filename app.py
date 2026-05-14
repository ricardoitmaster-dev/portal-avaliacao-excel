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
            [""], ["CONTEXTO PROFISSIONAL:"],
            [f"Prezado(a) {nome_aluno}, você foi designado para automatizar o relatório de vendas."],
            ["Sua missão é estruturar os dados para que a diretoria possa identificar a performance."],
            [""], ["DESAFIOS TÉCNICOS:"],
            ["1. CÁLCULO DE FATURAMENTO: Na coluna 'Venda Total', utilize operadores aritméticos."],
            ["   determine o montante total baseado no volume estocado e no valor unitário."],
            [""], ["2. ANÁLISE DE PERFORMANCE (LÓGICA CONDICIONAL): Na coluna 'Status', use a função 'SE'."],
            ["   - Se atingir ou superar 500, o status deve retornar 'META'."],
            ["   - Caso contrário, o sistema deve apontar a necessidade de 'REVISAR'."],
            [""], ["3. AUTOMAÇÃO (MACROS): Crie macros para ordenar por 'Produto' e 'Venda Total'."],
            ["   - Insira BOTÕES na planilha e atribua as macros correspondentes."],
            [""], ["REGRAS DE INTEGRIDADE:"],
            ["- Não altere a estrutura das colunas ou os nomes das abas."],
            ["- Salve como 'Pasta de Trabalho de Macro do Excel (.xlsm)'."],
            [""], ["Bom trabalho!"]
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

# --- INTERFACE ---
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
    turma = st.text_input("Identificação da Turma")
    email = st.text_input("E-mail Institucional/Pessoal")
    if st.button("Acessar Ambiente de Prova"):
        if nome and turma and email:
            st.session_state.aluno = {"nome": nome, "turma": turma.strip().upper(), "email": email}
            st.session_state.nome_esperado = f"Avaliacao_{nome.replace(' ','_')}.xlsx"
            st.session_state.excel_data = gerar_prova_excel(nome)
            st.session_state.etapa = 'prova'
            st.rerun()
else:
    st.title("Laboratório de Entrega")
    st.write(f"Candidato: **{st.session_state.aluno['nome']}** | Turma: **{st.session_state.aluno['turma']}**")
    st.download_button("📥 1. Baixar Caderno de Questões", st.session_state.excel_data, st.session_state.nome_esperado)
    st.divider()
    up_file = st.file_uploader("2. Enviar Solução (xlsx ou xlsm)", type=['xlsx', 'xlsm'])
    if st.button("🚀 3. Submeter para Correção"):
        if up_file:
            if up_file.name.split('.')[0] != st.session_state.nome_esperado.split('.')[0]:
                st.error("SISTEMA DE SEGURANÇA: Nome do arquivo divergente.")
            else:
                nota, feedback = calcular_nota(up_file)
                tutorial = gerar_feedback_pedagogico()
                novo_reg = pd.DataFrame([[st.session_state.aluno['nome'], st.session_state.aluno['turma'], nota]], columns=['Aluno', 'Turma', 'Nota'])
                novo_reg.to_csv("db_notas.csv", mode='a', header=not os.path.exists("db_notas.csv"), index=False)
                corpo = f"Aluno: {st.session_state.aluno['nome']}\nTurma: {st.session_state.aluno['turma']}\nNota: {nota}\n{feedback}\n\n{tutorial}"
                enviar_email(EMAIL_PROFESSOR, f"RESULTADO {nota}: {st.session_state.aluno['nome']}", corpo, up_file.getvalue(), up_file.name)
                enviar_email(st.session_state.aluno['email'], "Gabarito e Feedback - SENAI", corpo)
                st.success(f"Submissão realizada! Nota: {nota}")
                st.info(feedback)
                st.balloons()
    if st.button("Encerrar Sessão"):
        st.session_state.clear()
        st.rerun()

# --- PAINEL DE GESTÃO ---
st.divider()
with st.expander("👤 Painel de Controle (Professores / Gerência)"):
    tabs = st.tabs(["Acesso Professor", "Novo Cadastro", "Gerência Ricardo (ADM)"])
    
    with tabs[1]: # Cadastro
        st.write("### Cadastrar Novo Docente")
        p_nome = st.text_input("Nome do Professor")
        p_turma = st.text_input("Turma Autorizada", key="t_cad").strip().upper()
        p_senha = st.text_input("Definir Senha", type="password")
        if st.button("Finalizar Cadastro"):
            if p_nome and p_turma and p_senha:
                exists = False
                if os.path.exists("professores.csv"):
                    df_check = pd.read_csv("professores.csv")
                    if not df_check[(df_check['Professor'] == p_nome) & (df_check['Turma'] == p_turma)].empty:
                        exists = True
                
                if exists:
                    st.warning(f"O professor {p_nome} já está cadastrado para a turma {p_turma}.")
                else:
                    reg_p = pd.DataFrame([[p_nome, p_turma, p_senha]], columns=['Professor', 'Turma', 'Senha'])
                    reg_p.to_csv("professores.csv", mode='a', header=not os.path.exists("professores.csv"), index=False)
                    st.success(f"Professor {p_nome} vinculado à turma {p_turma}!")

    with tabs[0]: # Login Professor
        st.write("### Área do Docente")
        l_turma = st.text_input("Sua Turma", key="t_log").strip().upper()
        l_senha = st.text_input("Sua Senha", type="password", key="s_log")
        if st.button("Ver Minha Turma"):
            if os.path.exists("professores.csv"):
                profs = pd.read_csv("professores.csv", dtype={'Senha': str}) # Força leitura da senha como string
                # Validação exata de Turma e Senha
                valid_prof = profs[(profs['Turma'] == l_turma) & (profs['Senha'] == str(l_senha))]
                
                if not valid_prof.empty:
                    if os.path.exists("db_notas.csv"):
                        notas = pd.read_csv("db_notas.csv")
                        t_data = notas[notas['Turma'] == l_turma]
                        st.subheader(f"📊 Resultados - Turma {l_turma}")
                        st.metric("Alunos", len(t_data))
                        st.dataframe(t_data, use_container_width=True)
                    else: st.info("Nenhuma entrega registrada para sua turma.")
                else: st.error("Acesso Negado. Verifique a Turma e Senha.")
            else: st.error("Nenhum professor cadastrado no sistema.")

    with tabs[2]: # ADM Ricardo
        st.write("### Visão Geral do Sistema")
        m_senha = st.text_input("Senha de Gerência", type="password")
        if m_senha == "ricardoitmaster":
            if os.path.exists("db_notas.csv"):
                st.write("### Relatório Consolidado de Notas")
                st.dataframe(pd.read_csv("db_notas.csv"), use_container_width=True)
            if os.path.exists("professores.csv"):
                st.write("### Lista de Professores Ativos")
                st.dataframe(pd.read_csv("professores.csv"), use_container_width=True)
