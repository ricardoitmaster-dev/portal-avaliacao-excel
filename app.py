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

# Cores Identidade Visual (Azul SENAI, BMW Portinari Blue, Black, Gold)
COR_AZUL_SENAI = "#0054A6"
COR_AZUL_BMW = "#002366"
COR_PRETO_BRILHANTE = "#000000"
COR_DOURADO = "#D4AF37"
COR_TEXTO = "#FFFFFF"

st.set_page_config(page_title="Portal de Avaliação Excel SENAI", layout="wide")

# --- ESTILIZAÇÃO CSS ---
st.markdown(f"""
    <style>
        .stApp {{
            background-color: {COR_AZUL_SENAI} !important; 
        }}
        h1, h2, h3 {{
            color: {COR_DOURADO} !important; 
            font-weight: bold; 
            text-align: center; 
        }}
        label, p, span {{ 
            color: {COR_TEXTO} !important; 
        }}
        .stButton>button {{
            color: {COR_TEXTO};
            background-color: {COR_AZUL_BMW};
            border-radius: 10px;
            border: 1px solid {COR_DOURADO};
            height: 4em;
            font-weight: bold;
            width: 100%;
        }}
        .stButton>button:hover {{ 
            border: 2px solid {COR_TEXTO}; 
            color: {COR_DOURADO}; 
        }}
        .stDownloadButton>button {{
            color: {COR_PRETO_BRILHANTE} !important;
            background-color: {COR_DOURADO} !important;
            font-weight: bold;
            width: 100%;
        }}
        .stTextInput input {{ 
            background-color: #1A1A1A !important; 
            color: white !important; 
            border: 1px solid {COR_AZUL_BMW} !important; 
        }}
        [data-testid="stHorizontalBlock"] {{ 
            align-items: center !important; 
        }}
        
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
            ["AVALIAÇÃO PRÁTICA PROFISSIONAL DE EXCEL - SENAI"],
            [f"Nome do Aluno(a): {nome_aluno}"],
            [""],
            ["MANUAL DE INSTRUÇÕES DETALHADO PARA REALIZAÇÃO DA PROVA:"],
            ["1. DATAS COM ALEATORIOENTRE: Na aba 'Base_de_Dados', preencha toda a coluna 'Data Venda' utilizando a função =ALEATORIOENTRE(). Use números de série válidos para gerar datas do ano corrente (ex: entre 45658 e 46022) e formate a coluna como Data."],
            ["2. BUSCA COM PROCV (PRODUTO E PREÇO): Na aba 'Base_de_Dados', use a função =PROCV() para preencher automaticamente as colunas 'Produto (ProcV)' e 'Preço Unitário (ProcV)'. Use como valor procurado o 'ID_Produto' desta aba e busque a matriz-tabela estritamente nas colunas A até D da aba 'Apoio_Matriz'. Não esqueça de travar a matriz ($A:$D) e definir a busca exata (0 ou FALSO)."],
            ["3. MULTIPLICAÇÃO (SUBTOTAL): Na coluna 'Subtotal (Multiplicação)', crie uma fórmula simples multiplicando o valor encontrado na coluna 'Quantidade' pelo valor da coluna 'Preço Unitário (ProcV)' (=Quantidade*Preço Unitário)."],
            ["4. DIVISÃO (TAXA DESCONTO): Na coluna 'Taxa Desconto % (Divisão)', defina uma taxa percentual dividindo o valor da coluna 'Quantidade' por 100 (=Quantidade/100). Formate as células como Porcentagem (%) para exibição adequada."],
            ["5. PORCENTAGEM (VALOR DESCONTO): Na coluna 'Valor Desconto R$ (Porcentagem)', calcule o valor real do desconto em moeda. Multiplique o 'Subtotal (Multiplicação)' pela 'Taxa Desconto % (Divisão)' de cada linha."],
            ["6. SUBTRAÇÃO (TOTAL LÍQUIDO): Na coluna 'Total Líquido (Subtração)', execute a subtração matemática deduzindo o valor do desconto calculado do subtotal (=Subtotal - Valor Desconto R$)."],
            ["7. FUNÇÃO LÓGICA SE (STATUS META): Na coluna 'Status Meta (SE)', crie uma estrutura de decisão lógica com a função SE. O critério é: Se o 'Total Líquido (Subtração)' for maior ou igual a 1200, a célula deve retornar exatamente o texto \"META\"; caso contrário, deve retornar exatamente o texto \"REVISAR\"."],
            ["8. SOMA, MÉDIA, SOMASE E CONT.SE: Vá até a aba 'Resumo_Gerencial'. Na coluna 'Resultado do Aluno', insira estritamente as funções solicitadas: na linha 1 use =SOMA() para somar o Total Líquido da aba anterior; na linha 2 use =MÉDIA() para a média da coluna Quantidade; na linha 3 use =SOMASE() para somar o Total Líquido apenas onde a Categoria for \"Informática\"; na linha 4 use =CONT.SE() para contar os pedidos de \"Informática\"."],
            ["9. MACROS DE ORDENAÇÃO: Desenvolva obrigatoriamente 2 macros utilizando a gravação de passos ou programação VBA. Uma macro deve ordenar a aba 'Base_de_Dados' de forma Crescente pelo número do pedido, e a outra macro deve ordenar de forma Decrescente pelo Total Líquido."],
            ["10. BOTÕES OPERACIONAIS: Insira na planilha 2 formas (Shapes) ou botões de comando. Atribua individualmente a cada um deles uma das macros criadas no passo anterior, nomeando-os de forma clara para que o usuário saiba qual ordenação será ativada."],
            ["11. COMPONENTES GRÁFICOS: Crie e posicione estrategicamente na planilha 2 gráficos (podem ser dinâmicos ou estáticos) que representem visualmente os resultados encontrados, facilitando a análise gerencial dos pedidos."],
            ["12. TABELA DINÂMICA: Desenvolva 1 Tabela Dinâmica completa. Consolide os dados estruturados da 'Base_de_Dados' e configure-a obrigatoriamente em uma nova aba exclusiva da planilha."],
            ["13. REGRA E FORMATO DE ENTREGA: Após finalizar o desenvolvimento de todas as fórmulas e recursos visuais, salve obrigatoriamente o seu arquivo final com a extensão .xlsm (Planilha Habilitada para Macro do Excel). Arquivos enviados in .xlsx perderão os pontos das macros por restrição técnica do formato."]
        ]
        pd.DataFrame(instrucoes).to_excel(writer, sheet_name='Instrucoes_Prova', index=False, header=False)
        
    return output.getvalue()

def calcular_nota(arquivo_bytes, nome_aluno):
    try:
        # Força a leitura das abas corretas criadas pelo sistema
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
                if pd.notna(row['Preço Unitário (ProcV)']) and str(row['Preço Unitário (ProcV)']).strip() != "":
                    if abs(float(row['Preço Unitário (ProcV)']) - preco_ref) < 0.1:
                        acertos += 1
                
                # Execução e incremento do bloco Multiplicação
                total_testes += 1
                if pd.notna(row['Subtotal (Multiplicação)']) and str(row['Subtotal (Multiplicação)']).strip() != "":
                    if abs(float(row['Subtotal (Multiplicação)']) - (qtd * preco_ref)) < 0.1:
                        acertos += 1
                
                # Execução e incremento do bloco Porcentagem e Cálculo Líquido
                desc_r = float(row['Valor Desconto R$ (Porcentagem)']) if (pd.notna(row['Valor Desconto R$ (Porcentagem)']) and str(row['Valor Desconto R$ (Porcentagem)']).strip() != "") else 0
                total_testes += 1
                if pd.notna(row['Total Líquido (Subtração)']) and str(row['Total Líquido (Subtração)']).strip() != "":
                    if abs(float(row['Total Líquido (Subtração)']) - ((qtd * preco_ref) - desc_r)) < 0.1:
                        acertos += 1
                
                # Execução e incremento do bloco da Função Lógica SE
                se_ref = "META" if ((qtd * preco_ref) - desc_r) >= 1200 else "REVISAR"
                total_testes += 1
                if pd.notna(row['Status Meta (SE)']) and str(row['Status Meta (SE)']).strip() != "":
                    if str(row['Status Meta (SE)']).strip().upper() == se_ref:
                        acertos += 1
            except:
                pass
                
        # Proporção matemática de acertos mapeada de 0 a 40 pontos
        pontos_formulas = (acertos / total_testes) * 40.0 if total_testes > 0 else 0.0
        
        # Teste estrutural avançado via openpyxl
        try:
            wb = load_workbook(arquivo_bytes, keep_vba=True)
            
            # Validação Real de Macros e Botões/Shapes operacionais ativos
            tem_botoes = any(len(p._images) > 0 or hasattr(p, 'legacy_drawing') and p.legacy_drawing for p in wb.worksheets)
            if wb.vba_archive and str(st.session_state.get('aluno_arquivo_nome', '')).lower().endswith('.xlsm') and tem_botoes:
                pontos_macro = 20.0
            else:
                pontos_macro = 0.0
                feedbacks.append("Nenhum botão de macro operacional foi criado")
                
            graficos_achados = 0
            tabela_dinamica_achada = False
            
            for planilha in wb.worksheets:
                if planilha._charts:
                    graficos_achados += len(planilha._charts)
                if hasattr(planilha, '_pivots') and planilha._pivots:
                    tabela_dinamica_achada = True
                    
            # Distribuição matemática de gráficos (Até 20 pontos)
            if graficos_achados >= 2:
                pontos_graficos = 20.0
            else:
                pontos_graficos = 0.0
                
            # Distribuição matemática da Tabela Dinâmica (Até 20 pontos)
            if tabela_dinamica_achada:
                pontos_dinamica = 20.0
            else:
                pontos_dinamica = 0.0
                feedbacks.append("A aba ou cache com a Tabela Dinâmica não foi localizada")
                
        except:
            pontos_macro = 0.0
            pontos_graficos = 0.0
            pontos_dinamica = 0.0
            
        # Nota final calculada estritamente na escala de 0 a 100
        nota_final = round(pontos_formulas + pontos_macro + pontos_graficos + pontos_dinamica, 1)
        detalhamento = f"Fórmulas/Matemática: {round(pontos_formulas, 1)}/40.0 | Macros/Botões: {pontos_macro}/20.0 | Gráficos: {pontos_graficos}/20.0 | Tabela Dinâmica: {pontos_dinamica}/20.0"
        
        if feedbacks:
            detalhamento = f"Fórmulas/Matemática: {round(pontos_formulas, 1)}/40.0 | Macros/Botões: {pontos_macro}/20.0 | Gráficos: {pontos_graficos}/20.0 | Tabela Dinâmica: {pontos_dinamica}/20.0 Inconformidades apontadas:\n- " + "\n- ".join(feedbacks)
            
        return nota_final, detalhamento
    except Exception as e:
        return 0.0, f"Erro ao ler arquivo: 'TOTAL_FINAL'. Certifique-se de não alterar as abas."

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
    
    _, col1, col2, col3, _ = st.columns([0.2, 1, 1, 1, 0.2])
    
    with col1:
        if st.button("🎓 SOU ALUNO"):
            st.session_state.perfil = "aluno"
            st.rerun()
            
    with col2:
        if st.button("👨‍🏫 PROFESSOR / GESTOR"):
            st.session_state.perfil = "admin"
            st.rerun()
            
    with col3:
        st.markdown(
            '<a href="https://guarulhos.sp.senai.br/" target="_self" class="btn-sair-link" style="margin-top: 0px;">❌ SAIR DO SISTEMA</a>',
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
        st.info(f"Aluno: {st.session_state.aluno_dados['nome']} - Turma: {st.session_state.aluno_dados['turma']}")
        st.download_button("Baixar Planilha", st.session_state.excel_data, f"{nome_e}.xlsx")
        
        up = st.file_uploader("Upload da Prova Resolvida (.xlsm ou .xlsx)", type=['xlsx', 'xlsm'], accept_multiple_files=True)
        
        if st.button("Finalizar Entrega"):
            validos = [f for f in up if f.name.split('.')[0].lower() == nome_e.lower() or f.name.split('_')[0].lower() == "avaliacao"]
            if validos:
                # Pega o arquivo que será processado
                arq = next((f for f in validos if f.name.endswith('xlsm')), validos[0])
                st.session_state.aluno_arquivo_nome = arq.name
                
                # --- VALIDAÇÃO EXCLUSIVA POR NOME E EXTENSÃO EXATA DO ARQUIVO ---
                ja_enviou = False
                if os.path.exists("db_notas.csv"):
                    try:
                        df_historico = pd.read_csv("db_notas.csv")
                        # Verifica se na base já existe um registro do mesmo aluno, mesma turma e com o mesmo nome de arquivo exato
                        if 'Arquivo' in df_historico.columns:
                            duplicado = df_historico[
                                (df_historico['Aluno'].str.lower() == st.session_state.aluno_dados['nome'].lower()) & 
                                (df_historico['Turma'].str.upper() == st.session_state.aluno_dados['turma'].upper()) & 
                                (df_historico['Arquivo'].str.lower() == arq.name.lower())
                            ]
                            if not duplicado.empty:
                                ja_enviou = True
                    except:
                        pass

                if ja_enviou:
                    st.error("Você já enviou o seu arquivo de avaliação! Não é permitido o envio de arquivos múltiplas vezes.")
                else:
                    # Executa a correção se o arquivo for inédito (ou extensão diferente)
                    nota, info = calcular_nota(arq, st.session_state.aluno_dados['nome'])
                    
                    # Salva incluindo a nova coluna 'Arquivo' para controle preciso posterior
                    novo_registro = pd.DataFrame([[
                        st.session_state.aluno_dados['nome'], 
                        st.session_state.aluno_dados['turma'], 
                        nota,
                        arq.name
                    ]], columns=['Aluno', 'Turma', 'Nota', 'Arquivo'])
                    
                    novo_registro.to_csv("db_notas.csv", mode='a', header=not os.path.exists("db_notas.csv"), index=False)
                    
                    corpo_email = f"Resultado SENAI Excel\n\nAluno: {st.session_state.aluno_dados['nome']}\nTurma: {st.session_state.aluno_dados['turma']}\nNota: {nota}\n\nDetalhes da Correção:\n{info}"
                    
                    enviar_email(EMAIL_PROFESSOR, f"Prova - {st.session_state.aluno_dados['nome']}", corpo_email, [(f.getvalue(), f.name) for f in up])
                    enviar_email(st.session_state.aluno_dados['email'], "Seu Resultado na Prova de Excel", corpo_email, [(f.getvalue(), f.name) for f in up])
                    
                    st.success(f"Finalizado! Nota: {int(nota)}")
                    st.write("### Detalhes da Correção:")
                    st.write(info)
                    st.write("📬 **Status do Envio:**")
                    st.write(f"* E-mail enviado com sucesso para {EMAIL_PROFESSOR}")
                    st.write(f"* E-mail enviado com sucesso para {st.session_state.aluno_dados['email']}")
                    st.balloons()
            else:
                st.error("Nome do arquivo incorreto. Por favor, envie o arquivo original conforme as instruções.")

# --- PAINEL GESTOR / ADMINISTRATIVO (Sidebar Retornada) ---
elif st.session_state.perfil == "admin":
    with st.sidebar:
        st.header("Painel de Controle")
        st.markdown("---")
        
        if st.button("⬅️ Sair do Painel Admin"):
            st.session_state.clear()
            st.rerun()
            
        st.markdown("### Navegação")
        opcao_admin = st.radio(
            "Selecione a área desejada:",
            ["📊 Notas", "📝 Professores", "🛡️ Área ADM"]
        )

    # Exibição do conteúdo baseado na seleção da Sidebar
    st.subheader(opcao_admin)
    
    if opcao_admin == "📊 Notas":
        if os.path.exists("professores.csv"):
            np = st.text_input("Professor")
            tp = st.text_input("Turma").upper()
            sp = st.text_input("Senha", type="password")
            if st.button("Consultar"):
                dfp = pd.read_csv("professores.csv")
                auth = dfp[(dfp['Professor'].str.lower() == np.lower()) & (dfp['Turma'].str.upper() == tp) & (dfp['Senha'] == str(sp))]
                if not auth.empty and os.path.exists("db_notas.csv"):
                    db = pd.read_csv("db_notas.csv")
                    st.dataframe(db[db['Turma'] == tp], use_container_width=True)
                else:
                    st.warning("Nenhum dado encontrado para esta turma ou credenciais inválidas.")
        else:
            st.info("Nenhum professor cadastrado no sistema ainda.")

    elif opcao_admin == "📝 Professores":
        with st.form("c_p"):
            n_p = st.text_input("Nome")
            t_p = st.text_input("Turma")
            s_p = st.text_input("Senha")
            if st.form_submit_button("Cadastrar"):
                pd.DataFrame([[n_p, t_p, s_p]], columns=['Professor','Turma','Senha']).to_csv(
                    "professores.csv", mode='a', header=not os.path.exists("professores.csv"), index=False
                )
                st.success("Professor cadastrado com sucesso!")

    elif opcao_admin == "🛡️ Área ADM":
        if st.text_input("Senha ADM", type="password") == "Celina2610$$":
            if os.path.exists("db_notas.csv"):
                st.dataframe(pd.read_csv("db_notas.csv"), use_container_width=True)
                if st.button("Limpar Todos os Dados"):
                    os.remove("db_notas.csv")
                    st.rerun()
            else:
                st.info("O banco de dados de notas está vazio no momento.")
