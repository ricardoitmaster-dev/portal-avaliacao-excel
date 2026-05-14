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

# --- ESTILIZAÇÃO CSS AVANÇADA (Mantida íntegra) ---
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

# --- FUNÇÕES DE SUPORTE ---
def get_base64(file_path):
    try:
        with open(file_path, "rb") as f:
            return base64.b64encode(f.read()).decode()
    except: return ""

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
    itens = ["Notebook", "Mouse", "Teclado", "Monitor", "Impressora", "Cabo HDMI", "SSD 480GB", "Webcam"]
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
            ["Sua missão é estruturar os dados para que a diretoria identifique a performance."],
            [""],
            ["DESAFIOS TÉCNICOS:"],
            ["1. CÁLCULO DE FATURAMENTO: Na coluna 'Venda Total', use: Quantidade * Preço."],
            [""],
            ["2. ANÁLISE DE PERFORMANCE (LÓGICA CONDICIONAL): Na coluna 'Status', use SE."],
            ["   - Se atingir ou superar 500, o status deve ser 'META'."],
            ["   - Caso contrário, o status deve ser 'REVISAR'."],
            [""],
            ["3. AUTOMAÇÃO (MACROS):"],
            ["   - Crie uma Macro para ORDENAR a tabela pelo campo 'Produto' de A-Z."],
            ["   - Insira BOTÕES na planilha e atribua as macros correspondentes."],
            [""],
            ["REGRAS DE INTEGRIDADE:"],
            ["- Não altere a estrutura das colunas."],
            ["- Salve como 'Pasta de Trabalho de Macro do Excel (.xlsm)'."],
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
        feedback = f"Cálculos: {pv}/{total} | Lógica SE: {ps}/{total} | Macros: {'Detectadas' if tem_macro > 0 else 'Não detectadas'}"
        return nota, feedback
    except: return 0, "Erro: Certifique-se de preencher a aba 'Base_de_Dados' corretamente."

# --- INTERFACE STREAMLIT ---
if 'etapa' not in st.session_state: st.session_state.etapa = 'login'

col_logo, col_esp, col_assinatura = st.columns([1, 1, 1])
with col_logo: st.image("https://upload.wikimedia.org/wikipedia/commons/8/8c/SENAI_S%C3%A3o_Paulo_logo.png", width=120)
with col_assinatura:
    img_b64 = get_base64("Imagem para o app avaliação Excel_RicardoItmaster.png")
    st.markdown(f'<div style="text-align:right"><a href="https://www.youtube.com/channel/UCN8he2kZi8dhLs-1cIbfPzA" target="_blank"><img src="data:image/png;base64,{img_b64}" width="200" style="cursor:pointer;"></a></div>', unsafe_allow_html=True)

if st.session_state.etapa == 'login':
    st.title("Portal de Avaliação Profissional")
    nome = st.text_input("Nome Completo do Aluno")
    turma = st.text_input("Identificação da Turma")
    email = st.text_input("E-mail para Gabarito")
    if st.button("Acessar Ambiente de Prova"):
        if nome and turma and email:
            st.session_state.aluno = {"nome": nome, "turma": turma, "email": email}
            st.session_state.nome_esp = f"Avaliacao_{nome.replace(' ','_')}.xlsx"
            st.session_state.excel_data = gerar_prova_excel(nome)
            st.session_state.etapa = 'prova'
            st.rerun()
else:
    st.title("Laboratório de Entrega")
    st.write(f"Candidato: **{st.session_state.aluno['nome']}**")
    st.download_button("📥 1. Baixar Caderno de Questões", st.session_state.excel_data, st.session_state.nome_esp)
    st.divider()
    up = st.file_uploader("2. Enviar Solução Finalizada (.xlsx ou .xlsm)", type=['xlsx', 'xlsm'])
    if st.button("🚀 3. Submeter para Correção"):
        if up:
            if up.name.split('.')[0] != st.session_state.nome_esp.split('.')[0]:
                st.error("SISTEMA DE SEGURANÇA: Nome do arquivo divergente.")
            else:
                nota, feedback = calcular_nota(up)
                tutorial = gerar_feedback_pedagogico()
                corpo = f"Aluno: {st.session_state.aluno['nome']}\nNota: {nota}\n{feedback}\n\n{tutorial}"
                enviar_email(EMAIL_PROFESSOR, f"RESULTADO {nota}: {st.session_state.aluno['nome']}", corpo, up.getvalue(), up.name)
                enviar_email(st.session_state.aluno['email'], "Confirmação de Entrega - SENAI", corpo)
                
                # PERSISTÊNCIA PARA DASHBOARD
                novo_reg = pd.DataFrame([[st.session_state.aluno['nome'], st.session_state.aluno['turma'], nota]], columns=['Nome', 'Turma', 'Nota'])
                novo_reg.to_csv("resultados_senai.csv", mode='a', header=not os.path.exists("resultados_senai.csv"), index=False)
                
                st.success(f"Submissão realizada! Nota: {nota}")
                st.info(feedback)
                st.balloons()
        else: st.warning("Nenhum arquivo detectado.")

    if st.button("Encerrar Sessão"):
        st.session_state.clear()
        st.rerun()

# --- DASHBOARD DE GESTÃO (Final da página) ---
st.divider()
with st.expander("🔐 Área Administrativa (Acesso via Senha)"):
    pass_admin = st.text_input("Senha do Professor", type="password")
    if pass_admin == "senai122":
        st.subheader("📊 Performance Geral da Turma")
        try:
            df_dash = pd.read_csv("resultados_senai.csv")
            c1, c2, c3 = st.columns(3)
            c1.metric("Alunos Avaliados", len(df_dash))
            c2.metric("Média da Turma", round(df_dash['Nota'].mean(), 1))
            c3.metric("Aprovações (>=6)", len(df_dash[df_dash['Nota'] >= 6]))
            st.write("### Distribuição de Notas")
            st.bar_chart(df_dash['Nota'].value_counts())
            st.write("### Detalhamento por Aluno")
            st.dataframe(df_dash, use_container_width=True)
            csv_export = df_dash.to_csv(index=False).encode('utf-8')
            st.download_button("📥 Exportar Relatório Geral", csv_export, "relatorio_geral.csv", "text/csv")
        except: st.info("Aguardando as primeiras entregas para gerar o Dashboard.")
