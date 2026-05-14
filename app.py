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

# --- FUNÇÃO DE LIMPEZA TOTAL ---
def encerrar_sessao():
    for key in list(st.session_state.keys()):
        del st.session_state[key]
    st.rerun()

# --- GERADOR DE PLANILHA (COM ABA DE INSTRUÇÕES REINSERIDA) ---
def gerar_desafio_excel(nome_aluno):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        # 1. ABA DE DADOS
        itens = ["Monitor LED", "Teclado Mecânico", "Mouse Óptico", "SSD 1TB", "Placa de Vídeo", "Fonte 600W"]
        dados = [{"ID": i, "Produto": random.choice(itens), "Quantidade": random.randint(5, 50), "Preço Unit": round(random.uniform(100, 1500), 2), "Subtotal": 0, "IPI (10%)": 0, "Total Geral": 0, "Status": "", "Cod_Ref": random.randint(101, 106)} for i in range(1, 21)]
        df_base = pd.DataFrame(dados)
        df_base.to_excel(writer, sheet_name='DESAFIO_PRATICO', index=False)
        
        # 2. ABA DE REFERÊNCIA
        pd.DataFrame({"Cod": [101, 102, 103, 104, 105, 106], "Categoria": ["Periféricos", "Armazenamento", "Hardware", "Energia", "Vídeo", "Redes"]}).to_excel(writer, sheet_name='REFERENCIA', index=False)
        
        # 3. ABA DE INSTRUÇÕES (A QUE EU TINHA TIRADO)
        inst = [
            ["CANDIDATO:", nome_aluno.upper()],
            [""],
            ["REQUISITOS DA PROVA (0 A 100 PONTOS):"],
            ["1. TABELA:", "Converta o intervalo da aba DESAFIO_PRATICO em Objeto Tabela."],
            ["2. ARITMÉTICA:", "Calcule Subtotal (Qtd*Preço), IPI (Subtotal*0,10) e Total Geral (Soma)."],
            ["3. LÓGICA (SE):", "Se Total Geral > 5000 retornar 'ALTO CUSTO', senão 'NORMAL'."],
            ["4. BUSCA (PROCV):", "Crie coluna 'Categoria' e busque na aba REFERENCIA pelo Cod_Ref."],
            ["5. ALEATÓRIO:", "Use ALEATÓRIOENTRE(1;100) na coluna 'ID' para reordenar."],
            ["6. MACROS:", "Crie 2 macros de ordenação e insira BOTÕES para executá-las."],
            [""],
            ["⚠️ ATENÇÃO: Salve o arquivo mantendo seu nome original: 'Prova_" + nome_aluno.replace(" ","_") + "'"]
        ]
        pd.DataFrame(inst).to_excel(writer, sheet_name='INSTRUCOES', index=False, header=False)
    return output.getvalue()

# --- CORREÇÃO E SEGURANÇA DE NOME ---
def realizar_correcao_completa(lista_arquivos, nome_esperado):
    melhor_nota = 0
    relatorio = ""
    arquivos_validos = []
    
    # Filtrar apenas arquivos que contenham o nome do aluno
    for arq in lista_arquivos:
        if nome_esperado.lower() in arq.name.lower():
            arquivos_validos.append(arq)
        else:
            st.error(f"Arquivo REJEITADO: O arquivo '{arq.name}' não pertence ao aluno {nome_esperado}.")
            
    if not arquivos_validos:
        return -1, "Nenhum arquivo válido com seu nome foi detectado."

    for arq in arquivos_validos:
        try:
            df = pd.read_excel(arq, sheet_name='DESAFIO_PRATICO')
            wb = load_workbook(arq, keep_vba=True)
            p = 0
            res = [f"Análise: {arq.name}"]
            # Validações (25 pts cada bloco)
            if all(round(r['Subtotal'],1) == round(r['Quantidade']*r['Preço Unit'],1) for _,r in df.iterrows()): p += 25
            if 'Status' in df.columns and not df['Status'].isnull().all(): p += 25
            if 'Categoria' in df.columns: p += 25
            if arq.name.endswith('.xlsm') and wb.vba_archive: p += 25
            
            if p >= melhor_nota:
                melhor_nota = p
                relatorio = "\n".join(res) + f"\nPontuação calculada: {p}/100"
        except: continue
    return melhor_nota, relatorio

def enviar_emails(aluno_email, nome_aluno, nota, relato, arquivos):
    corpo = f"Avaliação SENAI\nAluno: {nome_aluno}\nNota: {nota}/100\n\n{relato}"
    try:
        for destinatario in [EMAIL_PROFESSOR, aluno_email]:
            msg = MIMEMultipart()
            msg['From'] = EMAIL_PROFESSOR
            msg['To'] = destinatario
            msg['Subject'] = f"Resultado SENAI - {nome_aluno}"
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
    n = st.text_input("Nome Completo", key="n_login")
    t = st.text_input("Turma", key="t_login").strip().upper()
    e = st.text_input("E-mail", key="e_login")
    if st.button("Acessar Laboratório"):
        if n and t and e:
            st.session_state.aluno = {"nome": n, "turma": t, "email": e}
            st.session_state.prova_origem = gerar_desafio_excel(n)
            st.session_state.nome_limpo = n.replace(" ","_")
            st.session_state.etapa = 'prova'
            st.rerun()

elif st.session_state.etapa == 'prova':
    st.title("Atividade Prática")
    st.write(f"Candidato: **{st.session_state.aluno['nome']}**")
    
    st.subheader("1. Download do Desafio")
    st.download_button("📥 Baixar Planilha", st.session_state.prova_origem, f"Prova_{st.session_state.nome_limpo}.xlsx")
    
    st.divider()
    
    st.subheader("2. Upload da Solução")
    st.warning(f"O arquivo deve conter seu nome: 'Prova_{st.session_state.nome_limpo}'")
    ups = st.file_uploader("Enviar (xlsx ou xlsm)", type=['xlsx', 'xlsm'], accept_multiple_files=True)
    
    if st.button("🚀 Finalizar e Enviar"):
        if ups:
            nota, relato = realizar_correcao_completa(ups, st.session_state.nome_limpo)
            if nota != -1:
                sucesso = enviar_emails(st.session_state.aluno['email'], st.session_state.aluno['nome'], nota, relato, ups)
                if sucesso:
                    st.success(f"Nota: {nota}/100. Relatório enviado para você e para o professor!")
                    st.text(relato)
            else:
                st.error("Erro: Nenhum dos arquivos enviados possui o seu nome.")

    st.divider()
    if st.button("🛑 Encerrar Sessão e Limpar Dados"):
        encerrar_sessao()
