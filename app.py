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

# --- ESTILIZAÇÃO CSS ---
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

# --- LÓGICA DE GERAÇÃO DE EXAMES ---
def gerar_prova_excel(nome_aluno):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        # Aba 1: Operações e Lógica (Base de Dados)
        itens = ["Monitor", "CPU", "Teclado", "Mouse", "Switch", "Roteador", "Cabo LAN"]
        dados = []
        for i in range(1, 21):
            qtd = random.randint(10, 100)
            preco = round(random.uniform(50, 500), 2)
            dados.append({
                "ID": i,
                "Produto": random.choice(itens),
                "Quantidade": qtd,
                "Preço Unit": preco,
                "Subtotal": 0,          # Desafio: Multiplicação
                "Desconto (5%)": 0,     # Desafio: Percentual
                "Total Líquido": 0,     # Desafio: Subtração
                "Status Estoque": "",   # Desafio: SE ou SES
                "Cod_Busca": random.randint(1, 20) # Para o PROCV
            })
        
        df_base = pd.DataFrame(dados)
        df_base.to_excel(writer, sheet_name='DESAFIO_PRATICO', index=False)
        
        # Aba 2: Tabela de Consulta (Para PROCV)
        df_ref = pd.DataFrame({
            "Cod": range(1, 21),
            "Categoria": [random.choice(["Hardware", "Redes", "Periféricos"]) for _ in range(20)]
        })
        df_ref.to_excel(writer, sheet_name='REFERENCIA', index=False)

        # Aba 3: Instruções Detalhadas
        instrucoes = [
            ["NOME DO ALUNO:", nome_aluno],
            ["DESAFIOS OBRIGATÓRIOS (VALE 100 PONTOS):"],
            ["1. FORMATAÇÃO:", "Transforme o intervalo da aba DESAFIO_PRATICO em TABELA."],
            ["2. ARITMÉTICA:", "Calcule Subtotal (Qtd * Preço), Desconto (Subtotal * 5%) e Total Líquido."],
            ["3. LÓGICA (SE/SES):", "Se Total Líquido > 5000, Status = 'CRÍTICO', senão 'NORMAL'."],
            ["4. BUSCA (PROCV):", "Na aba DESAFIO_PRATICO, crie uma coluna 'Categoria' e use PROCV buscando na aba REFERENCIA."],
            ["5. ALEATÓRIO:", "Use =ALEATÓRIOENTRE(1;100) em uma nova coluna de 'Sorteio'."],
            ["6. AUTOMAÇÃO:", "Crie 2 Macros de Ordenação (A-Z por Produto e Maior p/ Menor Total) e coloque BOTÕES."],
            ["7. SALVAMENTO:", "Salve obrigatoriamente como .XLSM (Pasta de Trabalho Habilitada para Macro)."]
        ]
        pd.DataFrame(instrucoes).to_excel(writer, sheet_name='INSTRUCOES', index=False, header=False)
        
    return output.getvalue()

# --- SISTEMA DE CORREÇÃO (0 A 100) ---
def calcular_nota_detalhada(arquivo_bytes):
    try:
        df = pd.read_excel(arquivo_bytes, sheet_name='DESAFIO_PRATICO', engine='openpyxl')
        wb = load_workbook(arquivo_bytes, keep_vba=True)
        
        pontos = 0
        feedbacks = []

        # 1. Validação Aritmética (20 pts)
        check_arit = True
        for i, r in df.iterrows():
            if round(r['Subtotal'], 2) != round(r['Quantidade'] * r['Preço Unit'], 2): check_arit = False
        if check_arit:
            pontos += 20
            feedbacks.append("✅ Aritmética (Soma/Mult/Sub): 20/20")
        else:
            feedbacks.append("❌ Aritmética: Erro nos cálculos. Correção: Subtotal = Qtd * Preço.")

        # 2. Lógica SE (20 pts)
        check_se = all(str(r['Status Estoque']).upper() == ("CRÍTICO" if r['Total Líquido'] > 5000 else "NORMAL") for _, r in df.iterrows())
        if check_se:
            pontos += 20
            feedbacks.append("✅ Lógica SE/SES: 20/20")
        else:
            feedbacks.append("❌ Lógica SE: Erro. Correção: =SE(Total>5000;\"CRÍTICO\";\"NORMAL\")")

        # 3. PROCV (20 pts)
        if 'Categoria' in df.columns and not df['Categoria'].isnull().any():
            pontos += 20
            feedbacks.append("✅ PROCV: 20/20")
        else:
            feedbacks.append("❌ PROCV: Não encontrado ou incompleto.")

        # 4. Macros e Botões (20 pts)
        if wb.vba_archive:
            pontos += 20
            feedbacks.append("✅ Macros e Automação: 20/20")
        else:
            feedbacks.append("❌ Macros: Arquivo não contém código VBA ou não é .XLSM.")

        # 5. Tabelas e Formatação (20 pts)
        # Nota: Checar se o objeto tabela existe no openpyxl
        ws = wb['DESAFIO_PRATICO']
        if ws.tables:
            pontos += 20
            feedbacks.append("✅ Formatação de Tabela: 20/20")
        else:
            feedbacks.append("❌ Formatação: O intervalo não foi convertido em Objeto Tabela.")

        return pontos, "\n".join(feedbacks)
    except Exception as e:
        return 0, f"Erro na correção: {str(e)}"

# --- INTERFACE ---
if 'etapa' not in st.session_state: st.session_state.etapa = 'login'

col_logo, _, col_ass = st.columns([1, 1, 1])
with col_logo: st.image("https://upload.wikimedia.org/wikipedia/commons/8/8c/SENAI_S%C3%A3o_Paulo_logo.png", width=120)
with col_ass: st.image("Imagem para o app avaliação Excel_RicardoItmaster.png", width=200)

if st.session_state.etapa == 'login':
    st.title("Portal de Avaliação Excel Profissional")
    nome = st.text_input("Nome do Aluno")
    turma = st.text_input("Turma").strip().upper()
    email = st.text_input("E-mail")
    if st.button("Iniciar Avaliação"):
        if nome and turma and email:
            st.session_state.aluno = {"nome": nome, "turma": turma, "email": email}
            st.session_state.excel_data = gerar_prova_excel(nome)
            st.session_state.etapa = 'prova'
            st.rerun()

elif st.session_state.etapa == 'prova':
    st.title("Laboratório de Testes")
    st.write(f"Aluno: {st.session_state.aluno['nome']}")
    st.download_button("📥 Baixar Planilha de Desafios", st.session_state.excel_data, f"Prova_{st.session_state.aluno['nome']}.xlsx")
    st.divider()
    up = st.file_uploader("Submeter Solução (.xlsm)", type=['xlsm'])
    if st.button("Finalizar e Corrigir"):
        if up:
            nota, relatorio = calcular_nota_detalhada(up)
            st.subheader(f"Sua Pontuação: {nota}/100")
            st.text(relatorio)
            
            # Registro
            pd.DataFrame([[st.session_state.aluno['nome'], st.session_state.aluno['turma'], nota]], columns=['Aluno', 'Turma', 'Nota']).to_csv("db_notas.csv", mode='a', header=not os.path.exists("db_notas.csv"), index=False)
            
            if nota >= 70: st.balloons()
            if st.button("Sair"):
                st.session_state.clear()
                st.rerun()

# --- PAINEL GESTÃO ---
with st.expander("👤 Gerência Ricardo (ADM)"):
    m_senha = st.text_input("Senha ADM", type="password")
    if m_senha == "ricardoitmaster":
        if os.path.exists("db_notas.csv"):
            st.dataframe(pd.read_csv("db_notas.csv"), use_container_width=True)
