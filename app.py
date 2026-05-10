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

# --- CONFIGURAÇÕES VIA SECRETS (COFRE DO STREAMLIT) ---
# O Streamlit buscará estas informações nas configurações avançadas que você preencherá depois
try:
    EMAIL_PROFESSOR = st.secrets["EMAIL_PROFESSOR"]
    SENHA_APP_GOOGLE = st.secrets["SENHA_APP_GOOGLE"]
except:
    # Caso rode fora da nuvem, usará seus dados padrão
    EMAIL_PROFESSOR = "ricardoitmaster@gmail.com"
    SENHA_APP_GOOGLE = "ugjhusmwnbmgzspv"

st.set_page_config(page_title="Avaliação Excel - Ricardo IT Master", page_icon="📊")

# Função de envio de e-mail SMTP Real
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
            ["AVALIAÇÃO PRÁTICA"], [f"ALUNO: {nome_aluno.upper()}"], [f"TURMA: {turma}"],
            [""], ["1. Calcule a Venda Total (Quantidade x Preço)"], ["2. Crie Status com a função SE"], ["3. Faça uma Tabela Dinâmica em nova aba"]
        ]
        pd.DataFrame(instrucoes).to_excel(writer, sheet_name='Instruções', index=False, header=False)
    return output.getvalue()

# --- Lógica de Navegação ---
if 'etapa' not in st.session_state:
    st.session_state.etapa = 'login'

if st.session_state.etapa == 'login':
    st.title("👨‍🏫 Portal de Avaliação Excel")
    st.write("Identifique-se para baixar sua prova personalizada.")
    nome = st.text_input("Nome Completo")
    turma = st.text_input("Sua Turma")
    email = st.text_input("Seu E-mail")
    
    if st.button("Acessar Avaliação"):
        if nome and turma and email:
            excel_data = gerar_prova_excel(nome, turma)
            nome_limpo = nome.replace(' ', '_')
            turma_limpa = turma.replace(' ', '_')
            st.session_state.nome_arquivo = f"Prova_Excel_{nome_limpo}_{turma_limpa}.xlsx"
            st.session_state.excel_data = excel_data
            st.session_state.aluno = {"nome": nome, "turma": turma, "email": email}
            st.session_state.etapa = 'prova'
            st.rerun()
        else:
            st.error("Por favor, preencha todos os campos.")

else:
    st.header(f"Área do Aluno: {st.session_state.aluno['nome']}")
    st.warning(f"⚠️ Salve seu arquivo exatamente como: **{st.session_state.nome_arquivo}**")
    
    st.download_button("📥 Baixar minha Prova Individual", st.session_state.excel_data, st.session_state.nome_arquivo)
    
    st.divider()
    
    st.subheader("Entrega da Prova")
    arquivo_upload = st.file_uploader("Anexe seu arquivo .xlsx resolvido", type=['xlsx'])
    
    col1, col2 = st.columns(2)
    
    with col1:
        if st.button("🚀 Enviar Prova Final"):
            if arquivo_upload:
                if arquivo_upload.name != st.session_state.nome_arquivo:
                    st.error(f"❌ NOME DO ARQUIVO INVÁLIDO! Envie como: {st.session_state.nome_arquivo}")
                else:
                    with st.spinner('Processando envio...'):
                        conteudo_arquivo = arquivo_upload.getvalue()
                        nome_do_arquivo = arquivo_upload.name
                        
                        # Envio para o Professor
                        corpo_prof = f"O aluno {st.session_state.aluno['nome']} (Turma {st.session_state.aluno['turma']}) entregou a prova."
                        p = enviar_email(EMAIL_PROFESSOR, f"PROVA RECEBIDA: {st.session_state.aluno['nome']}", corpo_prof, conteudo_arquivo, nome_do_arquivo)
                        
                        # Envio de comprovante com anexo para o Aluno
                        corpo_aluno = f"Olá {st.session_state.aluno['nome']}!\n\nConfirmamos o recebimento da sua prova de Excel.\nSegue em anexo o arquivo que você enviou como comprovante."
                        enviar_email(st.session_state.aluno['email'], "Comprovante de Entrega - Prova Excel", corpo_aluno, conteudo_arquivo, nome_do_arquivo)
                        
                        if p:
                            st.balloons()
                            st.success("Prova enviada com sucesso! Verifique seu e-mail.")
            else:
                st.error("Anexe o arquivo antes de enviar.")

    with col2:
        if st.button("🚪 Sair / Finalizar"):
            st.session_state.clear()
            st.rerun()
