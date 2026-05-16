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
    seed = int(hashlib.md5(nome_aluno.encode()).hexdigest(), 16) % 10000
    random.seed(seed)
    
    itens = ["Notebook", "Mouse", "Teclado", "Monitor", "Impressora", "Cabo HDMI", "SSD 480GB"]
    categorias = ["Informática", "Periféricos", "Periféricos", "Informática", "Periféricos", "Acessórios", "Hardware"]
    
    df_apoio = pd.DataFrame({
        "ID_Prod": range(1, 8),
        "Produto": itens,
        "Categoria": categorias,
        "Preço Base": [3500.0, 80.0, 150.0, 900.0, 600.0, 45.0, 280.0]
    })
    
    dados = []
    for i in range(1, 31):
        id_prod = random.randint(1, 7)
        qtd = random.randint(2, 15)
        dados.append({
            "Nº Pedido": i,
            "Data Venda": "",  
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
            ["CRÍTÉRIOS E REQUISITOS OBRIGATÓRIOS DA PROVA:"],
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
        
        for index, row in df.iterrows():
            try:
                id_p = row['ID_Produto']
                qtd = row['Quantidade']
                preco_ref = float(df_apoio[df_apoio['ID_Prod'] == id_p]['Preço Base'].values[0])
                
                # Execução e incremento do bloco PROCV
                total_testes += 1
                if pd.notna(row['Preço Unitário (ProcV)']):
                    if abs(float(row['Preço Unitário (ProcV)']) - preco_ref) < 0.1:
                        acertos += 1
                
                # Execução e incremento do bloco Multiplicação
                sub_ref = qtd * preco_ref
                total_testes += 1
                if pd.notna(row['Subtotal (Multiplicação)']):
                    if abs(float(row['Subtotal (Multiplicação)']) - sub_ref) < 0.1:
                        acertos += 1
                
                # Execução e incremento do bloco Porcentagem e Cálculo Líquido
                desc_r = float(row['Valor Desconto R$ (Porcentagem)']) if pd.notna(row['Valor Desconto R$ (Porcentagem)']) else 0
                tot_ref = sub_ref - desc_r
                total_testes += 1
                if pd.notna(row['Total Líquido (Subtração)']):
                    if abs(float(row['Total Líquido (Subtração)']) - tot_ref) < 0.1:
                        acertos += 1
                
                # Execução e incremento do bloco da Função Lógica SE
                se_ref = "META" if tot_ref >= 1200 else "REVISAR"
                total_testes += 1
                if pd.notna(row['Status Meta (SE)']):
                    if str(row['Status Meta (SE)']).strip().upper() == se_ref:
                        acertos += 1
            except:
                pass
                
        pontos_formulas = (acertos / total_testes) * 4.0 if total_testes > 0 else 0.0
        
        # Teste estrutural de macros, gráficos e tabelas via openpyxl
        try:
            wb = load_workbook(arquivo_bytes, keep_vba=True)
            
            if wb.vba_archive:
                pontos_macro = 2.0
            else:
                pontos_macro = 0.0
                feedbacks.append("Arquivo não contém código VBA ou macros ativas (.xlsm ausente)")
                
            graficos_achados = 0
            tabela_dinamica_achada = False
            
            for planilha in wb.worksheets:
                if planilha._charts:
                    graficos_achados += len(planilha._charts)
                if hasattr(planilha, '_pivots') and planilha._pivots:
                    tabela_dinamica_achada = True
                    
            if graficos_achados >= 2:
                pontos_graficos = 2.0
            elif graficos_achados == 1:
                pontos_graficos = 1.0
                feedbacks.append("Apenas 1 gráfico foi detectado na estrutura")
            else:
                pontos_graficos = 0.0
                feedbacks.append("Nenhum gráfico válido foi construído")
                
            if tabela_dinamica_achada:
                pontos_dinamica = 2.0
            else:
                pontos_dinamica = 0.0
                feedbacks.append("A aba ou cache com a Tabela Dinâmica não foi localizada")
                
        except:
            pontos_macro = 0.0
            pontos_graficos = 0.0
            pontos_dinamica = 0.0
            feedbacks.append("Incapaz de extrair validações de objetos (verifique o formato do arquivo enviado)")
            
        nota_final = round(pontos_formulas + pontos_macro + pontos_graficos + pontos_dinamica, 1)
        detalhamento = f"Fórmulas/Matemática: {round(pontos_formulas, 1)}/4.0 | Macros/Botões: {pontos_macro}/2.0 | Gráficos: {pontos_graficos}/2.0 | Tabela Dinâmica: {pontos_dinamica}/2.0"
        
        if feedbacks:
            detalhamento += "\nInconformidades apontadas:\n- " + "\n- ".join(feedbacks)
            
        return nota_final, detalhamento
    except Exception as e:
        return 0, f"Erro operacional ao analisar estrutura da planilha: {str(e)}"

# --- INICIALIZAÇÃO DE ESTADOS ---
if 'perfil' not in st.session_state: 
    st.session_state.perfil = None
if 'etapa_aluno' not in st.session_state: 
    st.session_state.etapa_aluno = 'login'

# --- LAYOUT DE CABEÇALHO ---
c1, c2, c3 = st.columns([1, 2, 1])
with c1: 
    st.image("https://upload.wikimedia.org/wikipedia/commons/8/8c/SENAI_S%C3%A3o_Paulo_logo.png", width=150)
with c3: 
    st.image("Imagem para o app avaliação Excel_RicardoItmaster.png", width=220)

# --- FLUXO DE EXECUÇÃO DE TELAS ---
if st.session_state.perfil is None:
    st.markdown("<br>", unsafe_allow_html=True)
    st.title("Sistema de Avaliação Técnica")
    st.write("### Bem-vindo! Selecione seu perfil de acesso:")
    
    _, col1, col2, _ = st.columns([0.8, 1, 1, 0.8])
    with col1:
        if st.button("🎓 SOU ALUNO"):
            st.session_state.perfil = "aluno"
            st.rerun()
    with col2:
        if st.button("👨‍🏫 PROFESSOR / GESTOR"):
            st.session_state.perfil = "admin"
            st.rerun()
            
    _, col_sair, _ = st.columns([1.5, 1, 1.5])
    with col_sair:
        st.markdown(
            '<a href="https://guarulhos.sp.senai.br/" target="_self" class="btn-sair-link">❌ SAIR DO SISTEMA</a>', 
            unsafe_allow_html=True
        )

# --- PAINEL OPERACIONAL DO ESTUDANTE ---
elif st.session_state.perfil == "aluno":
    if st.button("⬅️ Voltar"): 
        st.session_state.clear()
        st.rerun()
        
    if st.session_state.etapa_aluno == 'login':
        st.subheader("Identificação")
        with st.form("f_aluno"):
            n = st.text_input("Nome")
            t = st.text_input("Turma").upper()
            e = st.text_input("E-mail")
            if st.form_submit_button("Gerar Minha Prova"):
                if n and t and e:
                    st.session_state.aluno_dados = {"nome": n, "turma": t, "email": e}
                    st.session_state.excel_data = gerar_prova_excel(n)
                    st.session_state.etapa_aluno = 'prova'
                    st.rerun()
                    
    elif st.session_state.etapa_aluno == 'prova':
        nome_e = f"Avaliacao_{st.session_state.aluno_dados['nome'].replace(' ', '_')}"
        st.info(f"Envie o arquivo finalizado como: {nome_e}")
        st.download_button("📥 Baixar Prova", st.session_state.excel_data, f"{nome_e}.xlsx")
        
        up = st.file_uploader("Upload", type=['xlsx', 'xlsm'], accept_multiple_files=True)
        
        if st.button("🚀 Enviar Prova"):
            validos = [f for f in up if f.name.split('.')[0].lower() == nome_e.lower()]
            if validos:
                arq = next((f for f in validos if f.name.endswith('xlsm')), validos[0])
                nota, info = calcular_nota(arq, st.session_state.aluno_dados['nome'])
                
                pd.DataFrame([[st.session_state.aluno_dados['nome'], st.session_state.aluno_dados['turma'], nota]], columns=['Aluno','Turma','Nota']).to_csv("db_notas.csv", mode='a', header=not os.path.exists("db_notas.csv"), index=False)
                
                corpo_email = f"Resultado SENAI Excel\n\nAluno: {st.session_state.aluno_dados['nome']}\nTurma: {st.session_state.aluno_dados['turma']}\nNota: {nota}\n\nDetalhes da Correção:\n{info}"
                
                enviar_email(EMAIL_PROFESSOR, f"Prova - {st.session_state.aluno_dados['nome']}", corpo_email, [(f.getvalue(), f.name) for f in up])
                enviar_email(st.session_state.aluno_dados['email'], "Seu Resultado na Prova de Excel", corpo_email)
                
                st.success(f"Enviado! Nota: {nota}")
                st.info(info)
                st.balloons()
            else:
                st.error("Nome do arquivo incorreto. Por favor, envie o arquivo original conforme as instruções.")

# --- PAINEL GESTOR / ADMINISTRATIVO ---
elif st.session_state.perfil == "admin":
    if st.button("⬅️ Voltar"): 
        st.session_state.clear()
        st.rerun()
        
    t1, t2, t3 = st.tabs(["📊 Notas", "📝 Professores", "🛡️ Área ADM"])
    with t1:
        if os.path.exists("professores.csv"):
            np = st.text_input("Professor")
            sp = st.text_input("Senha", type="password")
            if st.button("Consultar"):
                dfp = pd.read_csv("professores.csv")
                auth = dfp[(dfp['Professor'].str.lower() == np.lower()) & (dfp['Senha'] == str(sp))]
                if not auth.empty and os.path.exists("db_notas.csv"):
                    db = pd.read_csv("db_notas.csv")
                    st.dataframe(db[db['Turma'].isin(auth['Turma'].unique())], use_container_width=True)
    with t2:
        with st.form("c_p"):
            n_p = st.text_input("Nome")
            t_p = st.text_input("Turma")
            s_p = st.text_input("Senha")
            if st.form_submit_button("Cadastrar"):
                pd.DataFrame([[n_p, t_p, s_p]], columns=['Professor','Turma','Senha']).to_csv("professores.csv", mode='a', header=not os.path.exists("professores.csv"), index=False)
                st.success("OK!")
    with t3:
        if st.text_input("Senha ADM", type="password") == "Celina2610$$":
            if os.path.exists("db_notas.csv"):
                st.dataframe(pd.read_csv("db_notas.csv"), use_container_width=True)
                if st.button("Limpar Dados"): 
                    os.remove("db_notas.csv")
                    st.rerun()
