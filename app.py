import streamlit as st
import pandas as pd
import random
import io
import smtplib
import os
import hashlib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from openpyxl import load_workbook

# --- CONFIGURAÇÕES DE AMBIENTE ---
EMAIL_PROFESSOR_PADRAO = "ricardoitmaster@gmail.com"
try:
    EMAIL_PROFESSOR = st.secrets.get("EMAIL_PROFESSOR", EMAIL_PROFESSOR_PADRAO)
    SENHA_APP_GOOGLE = st.secrets.get("SENHA_APP_GOOGLE", "ugjhusmwnbmgzspv")
except:
    EMAIL_PROFESSOR = EMAIL_PROFESSOR_PADRAO
    SENHA_APP_GOOGLE = "ugjhusmwnbmgzspv"

COR_AZUL_BMW = "#002366"
COR_PRETO_BRILHANTE = "#000000"
COR_DOURADO = "#D4AF37"
COR_TEXTO = "#FFFFFF"

st.set_page_config(page_title="Portal de Avaliação Excel SENAI 122", layout="wide")

# --- ESTILIZAÇÃO CSS ---
st.markdown(f"""
    <style>
        .stApp {{ background-color: {COR_PRETO_BRILHANTE} !important; }}
        h1, h2, h3 {{ color: {COR_DOURADO} !important; font-weight: bold; text-align: center; }}
        label, p, span {{ color: {COR_TEXTO} !important; }}
        .stButton>button {{ color: {COR_TEXTO}; background-color: {COR_AZUL_BMW}; border-radius: 10px; border: 1px solid {COR_DOURADO}; height: 4em; font-weight: bold; width: 100%; }}
        .stDownloadButton>button {{ color: {COR_PRETO_BRILHANTE} !important; background-color: {COR_DOURADO} !important; font-weight: bold; width: 100%; }}
        .btn-sair-link {{ display: flex; justify-content: center; align-items: center; text-decoration: none !important; background-color: #7B0000; color: white !important; font-weight: bold; border: 1px solid white; border-radius: 10px; height: 3.5em; width: 100%; margin-top: 30px; }}
    </style>
""", unsafe_allow_html=True)

# --- FUNÇÃO DE ENVIO DE E-MAIL COM LOG ---
def enviar_email_com_log(destinatario, assunto, corpo, arquivos=None):
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
        return True, f"E-mail enviado com sucesso para {destinatario}"
    except Exception as e:
        return False, f"Falha ao enviar para {destinatario}: {str(e)}"

# --- GERAÇÃO DE PROVA (30 REGISTROS) ---
def gerar_prova_30_registros(nome_aluno, turma):
    hash_id = int(hashlib.md5(f"{nome_aluno}{turma}".encode()).hexdigest(), 16)
    random.seed(hash_id % 2) # Apenas 2 variações por turma

    itens_lista = ["Notebook", "Mouse", "Teclado", "Monitor", "Cabo HDMI", "SSD 480GB", "Memória RAM", "Fonte ATX"]
    precos_lista = [3500.00, 85.50, 120.00, 850.00, 45.00, 280.00, 320.00, 250.00]
    
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        # Banco para PROCV
        df_banco = pd.DataFrame({"ID": range(1, 9), "PRODUTO": itens_lista, "PREÇO_BASE": precos_lista})
        df_banco.to_excel(writer, sheet_name='BANCO_DADOS', index=False)

        # Dados da Prova - 30 Registros
        dados = {
            "ID_ITEM": [random.randint(1, 8) for _ in range(30)],
            "DESCRIÇÃO": [""] * 30,
            "VALOR_UNIT": [0.0] * 30,
            "QTD": [random.randint(1, 15) for _ in range(30)],
            "SUBTOTAL": [0.0] * 30,
            "DESCONTO": [0.0] * 30,
            "TOTAL_FINAL": [0.0] * 30
        }
        pd.DataFrame(dados).to_excel(writer, sheet_name='RESOLUÇÃO', index=False)
        
        instr = [
            ["AVALIAÇÃO EXCEL - SENAI 122"], [f"ALUNO: {nome_aluno}"], [""],
            ["INSTRUÇÕES OBRIGATÓRIAS (VALOR 100 PONTOS):"],
            ["1. Use PROCV para preencher DESCRIÇÃO e VALOR_UNIT conforme o ID_ITEM."],
            ["2. SUBTOTAL = VALOR_UNIT * QTD."],
            ["3. SE: Se SUBTOTAL for maior que 1000, DESCONTO de 10% do subtotal, senão 0."],
            ["4. TOTAL_FINAL = SUBTOTAL - DESCONTO."],
            ["5. FORMATAÇÃO: Aplicar Moeda e converter em TABELA."],
            ["6. MACRO: Crie botão de Ordenação por Produto."]
        ]
        pd.DataFrame(instr).to_excel(writer, sheet_name='INSTRUÇÕES', index=False, header=False)
    return output.getvalue()

# --- MOTOR DE CORREÇÃO MINUCIOSA ---
def motor_correcao_senai(arquivo_submetido):
    try:
        # Carrega dados para conferência
        df_aluno = pd.read_excel(arquivo_submetido, sheet_name='RESOLUÇÃO')
        df_banco = pd.read_excel(arquivo_submetido, sheet_name='BANCO_DADOS')
        wb = load_workbook(arquivo_submetido, keep_vba=True)
        
        pontos = 0
        relatorio = []
        
        # 1. Verificação de Preenchimento (Não tirar 100 sem fazer nada)
        total_vazio = df_aluno['DESCRIÇÃO'].isna().sum() + df_aluno['TOTAL_FINAL'].sum()
        if total_vazio == 0 or df_aluno['TOTAL_FINAL'].sum() == 0:
            return 0, "PROVA EM BRANCO: O aluno não realizou os cálculos obrigatórios."

        # 2. Correção de Valores (30 itens)
        erros_calculo = 0
        for i, row in df_aluno.iterrows():
            # Busca valor real no banco
            try:
                base = df_banco[df_banco['ID'] == row['ID_ITEM']].iloc[0]
                val_esperado = base['PREÇO_BASE']
                sub_esperado = val_esperado * row['QTD']
                desc_esperado = (sub_esperado * 0.1) if sub_esperado > 1000 else 0
                tot_esperado = sub_esperado - desc_esperado
                
                # Compara com o que o aluno enviou (margem de erro de 0.01 para arredondamento)
                if abs(row['TOTAL_FINAL'] - tot_esperado) > 0.1:
                    erros_calculo += 1
            except: 
                erros_calculo += 1

        # Atribuição de Pontos
        if erros_calculo == 0: pontos += 70
        elif erros_calculo < 5: pontos += 50; relatorio.append(f"- Atenção: Você errou {erros_calculo} cálculos de 30.")
        else: pontos += 10; relatorio.append(f"- Erro grave: {erros_calculo} cálculos incorretos ou não realizados.")

        # 3. Verificação de Macro/VBA (30 pontos)
        if wb.vba_archive: 
            pontos += 30
        else: 
            relatorio.append("- Erro: Macro de ordenação não encontrada (salvou como .xlsm?).")

        feedback_final = "\n".join(relatorio) if relatorio else "Parabéns! Prova executada com perfeição."
        return pontos, feedback_final
    except Exception as e:
        return 0, f"Erro ao ler arquivo: {str(e)}. Certifique-se de não alterar as abas."

# --- INTERFACE ---
if 'perfil' not in st.session_state: st.session_state.perfil = None
if 'etapa' not in st.session_state: st.session_state.etapa = 'login'

c1, c2, c3 = st.columns([1, 2, 1])
with c1: st.image("https://upload.wikimedia.org/wikipedia/commons/8/8c/SENAI_S%C3%A3o_Paulo_logo.png", width=150)
with c3: st.image("Imagem para o app avaliação Excel_RicardoItmaster.png", width=200)

if st.session_state.perfil is None:
    st.title("Sistema de Avaliação Técnica")
    _, col1, col2, _ = st.columns([0.8, 1, 1, 0.8])
    with col1:
        if st.button("🎓 SOU ALUNO"): st.session_state.perfil = "aluno"; st.rerun()
    with col2:
        if st.button("👨‍🏫 PROFESSOR"): st.session_state.perfil = "admin"; st.rerun()
    st.markdown(f'<center><a href="https://guarulhos.sp.senai.br/" class="btn-sair-link">❌ SAIR DO SISTEMA</a></center>', unsafe_allow_html=True)

elif st.session_state.perfil == "aluno":
    if st.button("⬅️ Menu"): st.session_state.perfil = None; st.rerun()
    
    if st.session_state.etapa == 'login':
        with st.form("f_aluno"):
            n, t, e = st.text_input("Nome Completo"), st.text_input("Turma").upper(), st.text_input("E-mail")
            if st.form_submit_button("Gerar Avaliação (30 Itens)"):
                if n and t and e:
                    st.session_state.aluno = {"nome":n, "turma":t, "email":e}
                    st.session_state.prova = gerar_prova_30_registros(n, t)
                    st.session_state.etapa = 'prova'; st.rerun()
    
    elif st.session_state.etapa == 'prova':
        st.info(f"Aluno: {st.session_state.aluno['nome']} - Turma: {st.session_state.aluno['turma']}")
        st.download_button("📥 Baixar Planilha", st.session_state.prova, "Avaliacao_Excel_SENAI.xlsx")
        up = st.file_uploader("Upload da Prova Resolvida (.xlsm ou .xlsx)", type=['xlsm', 'xlsx'])
        if st.button("🚀 Finalizar Entrega"):
            if up:
                with st.spinner("Corrigindo prova e enviando e-mails..."):
                    nota, feedback = motor_correcao_senai(up)
                    relatorio = f"RELATÓRIO DE AVALIAÇÃO - SENAI 122\n\nAluno: {st.session_state.aluno['nome']}\nTurma: {st.session_state.aluno['turma']}\nNota: {nota}/100\n\nCorreção Técnica:\n{feedback}\n\nE-mail gerado automaticamente pelo sistema."
                    
                    # Envio para o Professor (Ricardo)
                    ok_p, msg_p = enviar_email_com_log(EMAIL_PROFESSOR, f"Entrega: {st.session_state.aluno['nome']}", relatorio, [(up.getvalue(), up.name)])
                    # Envio para o Aluno
                    ok_a, msg_a = enviar_email_com_log(st.session_state.aluno['email'], "Seu Resultado - Prova Excel", relatorio)
                    
                    st.success(f"Finalizado! Nota: {nota}")
                    st.write(f"### Detalhes da Correção:\n{feedback}")
                    
                    # Logs de confirmação de e-mail (Visíveis para o aluno e professor)
                    st.divider()
                    st.write("📫 **Status do Envio:**")
                    st.write(f"- {msg_p}")
                    st.write(f"- {msg_a}")
                    st.balloons()

elif st.session_state.perfil == "admin":
    if st.button("⬅️ Menu"): st.session_state.perfil = None; st.rerun()
    t1, t2 = st.tabs(["📊 Notas", "📝 Gestão"])
    with t1:
        st.write("Acesse com sua senha para ver o banco de dados.")
