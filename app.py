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
try:
    EMAIL_PROFESSOR = st.secrets["EMAIL_PROFESSOR"]
    SENHA_APP_GOOGLE = st.secrets["SENHA_APP_GOOGLE"]
except:
    EMAIL_PROFESSOR = "ricardoitmaster@gmail.com"
    SENHA_APP_GOOGLE = "ugjhusmwnbmgzspv"

# Cores Identidade Visual (BMW Portinari Blue, Black, Gold)
COR_AZUL_BMW = "#002366"
COR_PRETO_BRILHANTE = "#000000"
COR_DOURADO = "#D4AF37"
COR_TEXTO = "#FFFFFF"

st.set_page_config(page_title="Portal de Avaliação Excel", layout="wide")

# --- ESTILIZAÇÃO CSS ---
st.markdown(f"""
    <style>
        .stApp {{ background-color: {COR_PRETO_BRILHANTE} !important; }}
        h1, h2, h3 {{ color: {COR_DOURADO} !important; font-weight: bold; text-align: center; }}
        label, p, span {{ color: {COR_TEXTO} !important; }}
        .stButton>button {{ 
            color: {COR_TEXTO}; 
            background-color: {COR_AZUL_BMW}; 
            border-radius: 10px; 
            border: 1px solid {COR_DOURADO};
            height: 4em;
            font-weight: bold;
            width: 100%;
        }}
        .stButton>button:hover {{ border: 2px solid {COR_TEXTO}; color: {COR_DOURADO}; }}
        .stDownloadButton>button {{ 
            color: {COR_PRETO_BRILHANTE} !important; 
            background-color: {COR_DOURADO} !important; 
            font-weight: bold; 
            width: 100%; 
        }}
        .stTextInput input {{ background-color: #1A1A1A !important; color: white !important; border: 1px solid {COR_AZUL_BMW} !important; }}
        [data-testid="stHorizontalBlock"] {{ align-items: center !important; }}
        
        /* Estilo para o link que parece um botão de sair */
        .btn-sair-link {{
            display: flex;
            justify-content: center;
            align-items: center;
            text-decoration: none !important;
            background-color: #7B0000;
            color: white !important;
            font-weight: bold;
            border: 1px solid white;
            border-radius: 10px;
            height: 3.5em;
            width: 100%;
            margin-top: 30px;
            text-align: center;
        }}
        .btn-sair-link:hover {{
            background-color: #A00000;
            border: 2px solid white;
        }}
    </style>
""", unsafe_allow_html=True)

# --- FUNÇÕES DE APOIO ---
def enviar_email(destinatario, assunto, corpo, arquivos=None):
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
        return True
    except: return False

def gerar_prova_excel(nome_aluno):
    # Semente baseada no nome para garantir que a prova do aluno X seja sempre a mesma se ele baixar de novo
    seed = int(hashlib.md5(nome_aluno.encode()).hexdigest(), 16) % 10000
    random.seed(seed)
    
    itens = ["Notebook", "Mouse", "Teclado", "Monitor", "Impressora", "Cabo HDMI", "SSD 480GB"]
    categorias = ["Informática", "Periféricos", "Periféricos", "Informática", "Periféricos", "Acessórios", "Hardware"]
    
    # Aba de apoio para o ProcV
    df_apoio = pd.DataFrame({
        "ID": range(1, 8),
        "Produto": itens,
        "Categoria": categorias,
        "Preço Base": [3500.0, 80.0, 150.0, 900.0, 600.0, 45.0, 280.0]
    })
    
    dados = []
    # Obrigatoriamente 30 registros
    for i in range(1, 31):
        id_prod = random.randint(1, 7)
        qtd = random.randint(1, 20)
        dados.append({
            "Pedido": i, 
            "Data_Venda (AleatórioEntre)": "",
            "ID_Produto": id_prod, 
            "Produto (ProcV)": "", 
            "Quantidade": qtd, 
            "Preço Unit (ProcV)": "", 
            "Subtotal (Mult)": "", 
            "Desconto % (Divisão)": "",
            "Desconto R$ (Porcent)": "",
            "Total Final (Subtração)": "", 
            "Status Meta (SE)": ""
        })
        
    df = pd.DataFrame(dados)
    
    # Aba Resumo para Soma, SomaSE, Cont.SE
    resumo = pd.DataFrame({
        "Métricas e Funções": ["Soma Total de Vendas (SOMA)", "Média de Vendas (Soma/Divisão)", "Total Arrecadado Informática (SOMASE)", "Qtd Pedidos Informática (CONT.SE)"],
        "Valor Calculado": ["", "", "", ""]
    })
    
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, sheet_name='Base_de_Dados', index=False)
        df_apoio.to_excel(writer, sheet_name='Apoio_ProcV', index=False)
        resumo.to_excel(writer, sheet_name='Resumo_Analise', index=False)
        
        instrucoes = [
            ["AVALIAÇÃO PRÁTICA DE EXCEL - COMPLETA"], [f"Aluno: {nome_aluno}"], [""],
            ["INSTRUÇÕES OBRIGATÓRIAS:"], 
            ["1. MATEMÁTICA: Utilize +, -, *, / e % nos cálculos solicitados."],
            ["2. DATA: Na coluna Data_Venda, use =ALEATORIOENTRE() para gerar datas deste ano."],
            ["3. PROCV: Busque o Produto e o Preço Base na aba 'Apoio_ProcV' usando o ID_Produto."],
            ["4. MULTIPLICAÇÃO: Subtotal = Quantidade * Preço Unit."],
            ["5. PORCENTAGEM E DIVISÃO: Calcule a % de desconto dividindo o Preço Unit por 1000, e o valor R$ correspondente multiplicando pelo Subtotal."],
            ["6. SUBTRAÇÃO: Total Final = Subtotal - Desconto R$."],
            ["7. FUNÇÃO SE: Em Status, se Total Final >= 1000 escreva 'META', senão 'REVISAR'."],
            ["8. SOMASE E CONT.SE: Preencha a aba 'Resumo_Analise' utilizando as funções para a categoria 'Informática'."],
            ["9. GRÁFICOS: Crie 2 Gráficos na planilha."],
            ["10. TABELA DINÂMICA: Crie 1 Tabela Dinâmica em uma nova aba."],
            ["11. MACROS E BOTÕES: Crie 2 Macros de Ordenação (ex: A-Z e Z-A) e atribua a 2 Botões."],
            ["12. SALVAMENTO: Salve o arquivo final como .xlsm (habilitado para macros)"]
        ]
        pd.DataFrame(instrucoes).to_excel(writer, sheet_name='Instrucoes', index=False, header=False)
    return output.getvalue()

def calcular_nota(arquivo_bytes, nome_aluno):
    try:
        df = pd.read_excel(arquivo_bytes, sheet_name='Base_de_Dados', engine='openpyxl')
        df_apoio = pd.read_excel(arquivo_bytes, sheet_name='Apoio_ProcV', engine='openpyxl')
        
        acertos_calculo = 0
        total_verificacoes = 0
        erros_detalhes = []
        
        # Validando as lógicas matemáticas e lógicas linha a linha (30 registros)
        for i, row in df.iterrows():
            try:
                id_prod = row['ID_Produto']
                qtd = row['Quantidade']
                preco_esperado = float(df_apoio[df_apoio['ID'] == id_prod]['Preço Base'].values[0])
                
                # Checa PROCV
                if pd.notna(row['Preço Unit (ProcV)']) and abs(float(row['Preço Unit (ProcV)']) - preco_esperado) < 0.1: 
                    acertos_calculo += 1
                total_verificacoes += 1
                
                # Checa Multiplicação
                subt_esperado = qtd * preco_esperado
                if pd.notna(row['Subtotal (Mult)']) and abs(float(row['Subtotal (Mult)']) - subt_esperado) < 0.1: 
                    acertos_calculo += 1
                total_verificacoes += 1
                
                # Checa Subtração e Porcentagem (Aceita a lógica matemática aplicada se o final bater)
                desc = float(row['Desconto R$ (Porcent)']) if pd.notna(row['Desconto R$ (Porcent)']) else 0
                tot_esperado = subt_esperado - desc
                if pd.notna(row['Total Final (Subtração)']) and abs(float(row['Total Final (Subtração)']) - tot_esperado) < 0.1: 
                    acertos_calculo += 1
                total_verificacoes += 1
                
                # Checa SE
                status_esperado = "META" if tot_esperado >= 1000 else "REVISAR"
                if str(row['Status Meta (SE)']).strip().upper() == status_esperado: 
                    acertos_calculo += 1
                total_verificacoes += 1
                
            except:
                pass # Ignora erros de conversão se a linha estiver vazia/errada
        
        # Nota Cálculos vale 40% da prova
        pontos_calculo = (acertos_calculo / total_verificacoes) * 40 if total_verificacoes > 0 else 0
        
        try:
            wb = load_workbook(arquivo_bytes, keep_vba=True)
            
            # Valida Macros (.xlsm) - Vale 20% da prova
            pontos_macro = 20 if wb.vba_archive else 0
            if pontos_macro == 0: erros_detalhes.append("Macros não detectadas (Precisa ser .xlsm).")
            
            # Valida Gráficos e Tabela Dinâmica - Vale 40% da prova
            qtd_graficos = 0
            tem_td = False
            
            for sheet in wb.worksheets:
                if sheet._charts: qtd_graficos += len(sheet._charts)
                if hasattr(sheet, '_pivots') and sheet._pivots: tem_td = True
            
            pontos_grafico = 20 if qtd_graficos >= 2 else (10 if qtd_graficos == 1 else 0)
            pontos_td = 20 if tem_td else 0
            
            if qtd_graficos < 2: erros_detalhes.append(f"Encontrado {qtd_graficos} gráficos (Eram necessários 2).")
            if not tem_td: erros_detalhes.append("Tabela Dinâmica ausente.")
                
        except Exception as e: 
            pontos_macro = 0; pontos_grafico = 0; pontos_td = 0
            erros_detalhes.append("Erro ao ler abas/objetos (Não renomeie as abas base).")
            
        nota_final = round(pontos_calculo + pontos_macro + pontos_grafico + pontos_td, 1)
        
        resumo_fb = f"Fórmulas: {round(pontos_calculo,1)}/40 | Macros: {pontos_macro}/20 | Gráficos: {pontos_grafico}/20 | Dinâmica: {pontos_td}/20"
        if erros_detalhes:
            resumo_fb += "\nFaltou:\n- " + "\n- ".join(erros_detalhes)
            
        return nota_final, resumo_fb
    except Exception as e: return 0, f"Erro crítico na correção: {str(e)}"

# --- INICIALIZAÇÃO ---
if 'perfil' not in st.session_state: st.session_state.perfil = None
if 'etapa_aluno' not in st.session_state: st.session_state.etapa_aluno = 'login'

# --- CABEÇALHO ---
c1, c2, c3 = st.columns([1, 2, 1])
with c1: st.image("https://upload.wikimedia.org/wikipedia/commons/8/8c/SENAI_S%C3%A3o_Paulo_logo.png", width=150)
with c3: st.image("Imagem para o app avaliação Excel_RicardoItmaster.png", width=220)

# --- TELA INICIAL ---
if st.session_state.perfil is None:
    st.markdown("<br>", unsafe_allow_html=True)
    st.title("Sistema de Avaliação Técnica")
    st.write("### Bem-vindo! Selecione seu perfil de acesso:")
    
    # DISPOSIÇÃO HORIZONTAL
    _, col1, col2, _ = st.columns([0.8, 1, 1, 0.8])
    with col1:
        if st.button("🎓 SOU ALUNO"):
            st.session_state.perfil = "aluno"; st.rerun()
    with col2:
        if st.button("👨‍🏫 PROFESSOR / GESTOR"):
            st.session_state.perfil = "admin"; st.rerun()
    
    # SAIR COM LINK PARA SENAI GUARULHOS
    _, col_sair, _ = st.columns([1.5, 1, 1.5])
    with col_sair:
        st.markdown(
            '<a href="https://guarulhos.sp.senai.br/" target="_self" class="btn-sair-link">❌ SAIR DO SISTEMA</a>', 
            unsafe_allow_html=True
        )

# --- ÁREA DO ALUNO ---
elif st.session_state.perfil == "aluno":
    if st.button("⬅️ Voltar"): 
        st.session_state.clear() # Correção para garantir a limpeza total da sessão ao voltar
        st.rerun()
        
    if st.session_state.etapa_aluno == 'login':
        st.subheader("Identificação")
        with st.form("f_aluno"):
            n, t, e = st.text_input("Nome"), st.text_input("Turma").upper(), st.text_input("E-mail")
            if st.form_submit_button("Gerar Minha Prova"):
                if n and t and e:
                    st.session_state.aluno_dados, st.session_state.excel_data = {"nome":n,"turma":t,"email":e}, gerar_prova_excel(n)
                    st.session_state.etapa_aluno = 'prova'; st.rerun()
                    
    elif st.session_state.etapa_aluno == 'prova':
        nome_e = f"Avaliacao_{st.session_state.aluno_dados['nome'].replace(' ', '_')}"
        st.info(f"Envie o arquivo finalizado como: {nome_e}")
        st.download_button("📥 Baixar Prova", st.session_state.excel_data, f"{nome_e}.xlsx")
        
        up = st.file_uploader("Upload", type=['xlsx', 'xlsm'], accept_multiple_files=True)
        
        if st.button("🚀 Enviar Prova"):
            validos = [f for f in up if f.name.split('.')[0].lower() == nome_e.lower()]
            if validos:
                arq = next((f for f in validos if f.name.endswith('xlsm')), validos[0])
                
                # AQUI FOI ALTERADO: Passando o nome do aluno para o cálculo determinístico
                nota, info = calcular_nota(arq, st.session_state.aluno_dados['nome'])
                
                pd.DataFrame([[st.session_state.aluno_dados['nome'], st.session_state.aluno_dados['turma'], nota]], columns=['Aluno','Turma','Nota']).to_csv("db_notas.csv", mode='a', header=not os.path.exists("db_notas.csv"), index=False)
                
                corpo_email = f"Resultado SENAI Excel\n\nAluno: {st.session_state.aluno_dados['nome']}\nTurma: {st.session_state.aluno_dados['turma']}\nNota: {nota}\n\nDetalhes da Correção:\n{info}"
                
                enviar_email(EMAIL_PROFESSOR, f"Prova - {st.session_state.aluno_dados['nome']}", corpo_email, [(f.getvalue(), f.name) for f in up])
                enviar_email(st.session_state.aluno_dados['email'], "Seu Resultado na Prova de Excel", corpo_email)
                
                st.success(f"Enviado! Nota: {nota}")
                st.info(info)
                st.balloons()
            else:
                st.error("Nome do arquivo incorreto. Por favor, envie o arquivo que você baixou.")

# --- ÁREA ADMINISTRATIVA ---
elif st.session_state.perfil == "
