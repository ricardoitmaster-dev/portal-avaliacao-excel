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
    except:
        return False

def gerar_prova_excel(nome_aluno):
    # Semente pseudo-aleatória baseada no nome para gerar dados exclusivos por estudante
    seed = int(hashlib.md5(nome_aluno.encode()).hexdigest(), 16) % 10000
    random.seed(seed)
    
    itens = ["Notebook", "Mouse", "Teclado", "Monitor", "Impressora", "Cabo HDMI", "SSD 480GB"]
    categorias = ["Informática", "Periféricos", "Periféricos", "Informática", "Periféricos", "Acessórios", "Hardware"]
    
    # Tabela de Apoio para aplicação do PROCV
    df_apoio = pd.DataFrame({
        "ID_Prod": range(1, 8),
        "Produto": itens,
        "Categoria": categorias,
        "Preço Base": [3500.0, 80.0, 150.0, 900.0, 600.0, 45.0, 280.0]
    })
    
    dados = []
    # Criação estrita dos 30 registros solicitados
    for i in range(1, 31):
        id_prod = random.randint(1, 7)
        qtd = random.randint(2, 15)
        dados.append({
            "Nº Pedido": i,
            "Data Venda": "",  # Aluno deve preencher com ALEATORIOENTRE para datas
            "ID_Produto": id_prod,
            "Produto (ProcV)": "",
            "Quantidade": qtd,
            "Preço Unitário (ProcV)": "",
            "Subtotal (Multiplicação)": "",
            "Taxa Desconto % (Divisão)": "",
            "Valor Desconto R$ (Porcentagem)": "",
            "Total Líquido (Subtração)": "",
            "Status Meta (SE)": ""
        })
        
    df = pd.DataFrame(dados)
    
    # Aba Resumo para aplicação das fórmulas estatísticas e consolidadas
    df_resumo = pd.DataFrame({
        "Indicador Analítico": [
            "Faturamento Bruto Geral (SOMA)", 
            "Quantidade Média de Itens por Pedido", 
            "Faturamento Total da Categoria Informática (SOMASE)", 
            "Total de Pedidos Gerados para Informática (CONT.SE)"
        ],
        "Fórmula a ser Aplicada": [
            "Inserir a função SOMA na coluna ao lado",
            "Inserir cálculo de Média ou Soma/Divisão",
            "Inserir a função SOMASE baseada na tabela de apoio",
            "Inserir a função CONT.SE filtrando por Informática"
        ],
        "Resultado do Aluno": ["", "", "", ""]
    })
    
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, sheet_name='Base_de_Dados', index=False)
        df_apoio.to_excel(writer, sheet_name='Apoio_Matriz', index=False)
        df_resumo.to_excel(writer, sheet_name='Resumo_Gerencial', index=False)
        
        instrucoes = [
            ["AVALIAÇÃO PRÁTICA PROFISSIONAL DE EXCEL"],
            [f"Nome do Aluno(a): {nome_aluno}"],
            [""],
            ["CRITÉRIOS E REQUISITOS OBRIGATÓRIOS DA PROVA:"],
            ["1. DATAS COM ALEATORIOENTRE: Preencha a coluna 'Data Venda' utilizando =ALEATORIOENTRE() com números de série de datas deste ano."],
            ["2. BUSCA COM PROCV: Preencha as colunas 'Produto' e 'Preço Unitário' buscando os dados da aba 'Apoio_Matriz' pelo ID."],
            ["3. MULTIPLICAÇÃO: Calcule o 'Subtotal' multiplicando a Quantidade pelo Preço Unitário."],
            ["4. DIVISÃO: Na 'Taxa Desconto %', defina um critério de divisão (ex: dividir a Quantidade por 100 para achar um percentual)."],
            ["5. PORCENTAGEM: No 'Valor Desconto R$', calcule a porcentagem aplicando a taxa sobre o Subtotal."],
            ["6. SUBTRAÇÃO: O 'Total Líquido' deve ser calculado subtraindo o Desconto do Subtotal."],
            ["7. FUNÇÃO LÓGICA SE: No 'Status Meta', aplique a função SE (Se Total Líquido >= 1200 então 'META', caso contrário 'REVISAR')."],
            ["8. SOMA, SOMASE e CONT.SE: Preencha os campos vazios da aba 'Resumo_Gerencial' utilizando estritamente essas funções."],
            ["9. MACROS DE ORDENAÇÃO: Desenvolva 2 macros gravadas ou em VBA destinadas a ordenar os registros da base."],
            ["10. BOTÕES OPERACIONAIS: Crie 2 botões na planilha e atribua cada um deles a uma macro de ordenação desenvolvida."],
            ["11. COMPONENTES GRÁFICOS: Insira obrigatoriamente 2 gráficos dinâmicos ou estáticos que facilitem a leitura dos dados."],
            ["12. TABELA DINÂMICA: Desenvolva 1 Tabela Dinâmica consolidando as informações em uma nova aba exclusiva."],
            ["13. ENTREGA: Salve seu arquivo final impreterivelmente no formato .xlsm (Planilha Habilitada para Macro do Excel)."]
        ]
        pd.DataFrame(instrucoes).to_excel(writer, sheet_name='Instrucoes_Prova', index=False, header=False)
        
    return output.getvalue()

def calcular_nota(arquivo_bytes, nome_aluno):
    try:
        df = pd.read_excel(arquivo_bytes, sheet_name='Base_de_Dados', engine='openpyxl')
        df_apoio = pd.read_excel(arquivo_bytes, sheet_name='Apoio_Matriz', engine='openpyxl')
        
        acertos = 0
        total_testes = 0
        feedbacks = []
        
        # Validação do preenchimento e consistência lógica dos 30 registros
        for index, row in df.iterrows():
            try:
                id_p = row['ID_Produto']
                qtd = row['Quantidade']
                preco_ref = float(df_apoio[df_apoio['ID_Prod'] == id_p]['Preço Base'].values[0])
                
                # Teste PROCV
                if pd.notna(row['Preço Unitário (ProcV)']) and abs(float(row['Preço Unitário (ProcV)']) - preco_ref) < 0.1:
                    acertos += 1
                total_testes += 1
                
                # Teste Multiplicação
                sub_ref = qtd * preco_ref
                if pd.notna(row['Subtotal (Multiplicação)']) and abs(float(row['Subtotal (Multiplicação)']) - sub_ref) < 0.1:
                    acertos += 1
                total_testes += 1
                
                # Teste Subtração e Porcentagem
                desc_r = float(row['Valor Desconto R$ (Porcentagem)']) if pd.notna(row['Valor Desconto R$ (Porcentagem)']) else 0
                tot_ref = sub_ref - desc_r
                if pd.notna(row['Total Líquido (Subtração)']) and abs(float(row['Total Líquido (Subtração)']) - tot_ref) < 0.1:
                    acertos += 1
                total_testes += 1
