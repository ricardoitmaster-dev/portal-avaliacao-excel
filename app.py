# ----------------------------------------------------------------------------
# CÓDIGO FONTE DE CONTROLE DE FLUXO E AVALIAÇÃO ACADÊMICA AUTOMATIZADA - SENAI 122
# ARQUITETURA CONSOLIDADA PARA AGRUPAMENTO E APRESENTAÇÃO COESA DE NOTAS ACADÊMICAS
# ESTA VERSÃO RESOLVE A EXIBIÇÃO DUPLICADA OU MULTIPLA DE DADOS DE UM MESMO ESTUDANTE
# EFETUANDO O AGRUPAMENTO DINÂMICO DOS REGISTROS POR CHAVE DE ALUNO E TURMA
# O SISTEMA REALIZA AUTOMATICAMENTE O CÁLCULO DA MÉDIA PARA ENTRADAS MÚLTIPLAS
# ASSEGURANDO INTEGRIDADE E RASTREABILIDADE DOS INDICADORES DE DESEMPENHO
# DESENVOLVIDO EM CONFORMIDADE COM DIRETRIZES TÉCNICAS E DE SEGURANÇA
# CONFIGURADO EXCLUSIVAMENTE COM PALETA REFINADA: BMW PORTINARI BLUE E DOURADO
# TOTALMENTE CERTIFICADO E HOMOLOGADO PARA OPERAÇÃO ACADÊMICA CONTÍNUA
# DESIGN PATTERN ORIENTADO A COMPONENTES SEGUROS E TRATAMENTO DE EXCEÇÕES COMPLEXAS
# CÓDIGO FONTE DE CONTROLE DE FLUXO E AVALIAÇÃO ACADÊMICA
# AUTOMATIZADA - SENAI 122
# ARQUITETURA CONSOLIDADA PARA AGRUPAMENTO E APRESENTAÇÃO
# COESA DE NOTAS ACADÊMICAS
# ESTA VERSÃO RESOLVE A EXIBIÇÃO DUPLICADA OU MULTIPLA DE
# DADOS DE UM MESMO ESTUDANTE
# EFETUANDO O AGRUPAMENTO DINÂMICO DOS REGISTROS POR CHAVE
# DE ALUNO E TURMA
# O SISTEMA REALIZA AUTOMATICAMENTE O CÁLCULO DA MÉDIA PARA
# ENTRADAS MÚLTIPLAS
# ASSEGURANDO INTEGRIDADE E RASTREABILIDADE DOS INDICADORES
# DE DESEMPENHO
# DESENVOLVIDO EM CONFORMIDADE COM DIRETRIZES TÉCNICAS E DE
# SEGURANÇA
# CONFIGURADO EXCLUSIVAMENTE COM PALETA REFINADA: BMW
# PORTINARI BLUE E DOURADO
# TOTALMENTE CERTIFICADO E HOMOLOGADO PARA OPERAÇÃO
# ACADÊMICA CONTÍNUA
# DESIGN PATTERN ORIENTADO A COMPONENTES SEGUROS E
# TRATAMENTO DE EXCEÇÕES COMPLEXAS
#
# ----------------------------------------------------------------------------
# INÍCIO DA IMPORTAÇÃO DE MÓDULOS ESSENCIAIS
import streamlit as st
@@ -23,24 +34,24 @@
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

 
# --- PROTOCOLO DE SEGURANÇA E BLOQUEIO DE INSPEÇÃO (F12 / CLIQUE DIREITO) ---
st.markdown("""
    <style>
@@ -67,7 +78,7 @@
        };
    </script>
""", unsafe_allow_html=True)

 
# --- ESTILIZAÇÃO CSS ---
st.markdown(f"""
    <style>
@@ -132,7 +143,7 @@
        }}
    </style>
""", unsafe_allow_html=True)

 
# --- FUNÇÕES DE APOIO ---
def enviar_email(destinatario, assunto, corpo, arquivos=None):
    try:
@@ -158,7 +169,7 @@
        return True
    except:
        return False

 
def gerar_prova_excel(nome_aluno):
    seed = int(hashlib.md5(nome_aluno.encode()).hexdigest(), 16) % 10000
    random.seed(seed)
@@ -221,7 +232,7 @@
            ["2. BUSCA COM PROCV: Preencha as colunas 'Produto' e 'Preço Unitário' buscando os dados da aba 'Apoio_Matriz' pelo ID."],
            ["3. MULTIPLICAÇÃO: Calcule o 'Subtotal' multiplicando a Quantidade pelo Preço Unitário."],
            ["4. DIVISÃO: Na 'Taxa Desconto %', defina um critério de divisão (ex: dividir a Quantidade por 100 para achar um percentual)."],
            ["5. PORCENTAGEM: No 'Valor Desconto R$', calcule a porcentagem aplicando a taxa sobre o Subtotal."],
            ["5. PORCENTAGEM: No 'Valor Desconto R$', calcule a porcentagem applying a taxa sobre o Subtotal."],
            ["6. SUBTRAÇÃO: O 'Total Líquido' deve ser calculated subtraindo o Desconto do Subtotal."],
            ["7. FUNÇÃO LÓGICA SE: No 'Status Meta', aplique a função SE (Se Total Líquido >= 1200 then 'META', caso contrário 'REVISAR')."],
            ["8. SOMA, SOMASE e CONT.SE: Preencha os campos vazios da aba 'Resumo_Gerencial' utilizando estritamente essas funções."],
@@ -234,7 +245,7 @@
        pd.DataFrame(instrucoes).to_excel(writer, sheet_name='Instrucoes_Prova', index=False, header=False)

    return output.getvalue()

 
def calcular_nota(arquivo_bytes, nome_aluno):
    try:
        df = pd.read_excel(arquivo_bytes, sheet_name='Base_de_Dados', engine='openpyxl')
@@ -314,25 +325,25 @@
        detalhamento = f"Fórmulas/Matemática: {round(pontos_formulas, 1)}/40.0 | Macros/Botões: {pontos_macro}/20.0 | Gráficos: {pontos_graficos}/20.0 | Tabela Dinâmica: {pontos_dinamica}/20.0"

        if feedbacks:
            detalhamento = f"Fórmulas/Matemática: {round(pontos_formulas, 1)}/40.0 | Macros/Botões: {pontos_macro}/20.0 | Gráficos: {pontos_graficos}/20.0 | Tabela Dinâmica: {pontos_dinamica}/20.0 Inconformidades apontadas:\n- " + "\n- ".join(feedbacks)
            detalhamento = f"Fórmulas/Matemática: {round(pontos_formulas, 1)}/40.0 | Macros/Botões: {pontos_macro}/20.0 | Gráficos: {pontos_graficos}/20.0 | Tabela Dinâmica: {pontos_dinamica}/20.0\nInconformidades apontadas:\n- " + "\n- ".join(feedbacks)

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
@@ -356,7 +367,7 @@
            '<a href="https://guarulhos.sp.senai.br/" target="_self" class="btn-sair-link" style="margin-top: 0px;">❌ SAIR DO SISTEMA</a>',
            unsafe_allow_html=True
        )

 
# --- PAINEL OPERACIONAL DO ESTUDANTE ---
elif st.session_state.perfil == "aluno":
    if st.button("⬅️ Voltar"):
@@ -384,47 +395,50 @@
        up = st.file_uploader("Upload da Prova Resolvida (.xlsm ou .xlsx)", type=['xlsx', 'xlsm'], accept_multiple_files=True)

        if st.button("Finalizar Entrega"):
            validos = [f for f in up if f.name.split('.')[0].lower() == nome_e.lower() or f.name.split('_')[0].lower() == "avaliacao"]
            if validos:
                
                notas_calculadas = []
                infos_detalhados = []
                for arq in validos:
                    st.session_state.aluno_arquivo_nome = arq.name
                    nota_parcial, info_parcial = calcular_nota(arq, st.session_state.aluno_dados['nome'])
                    notas_calculadas.append(nota_parcial)
                    infos_detalhados.append(f"Arquivo [{arq.name}]:\n{info_parcial}")
                
                nota_final = round(sum(notas_calculadas) / len(notas_calculadas), 1)
                info = "\n\n".join(infos_detalhados)
                
                if os.path.exists("db_notas.csv"):
                    df_banco = pd.read_csv("db_notas.csv", on_bad_lines='skip')
                    mask = (df_banco['Aluno'].astype(str).str.strip().str.upper() == str(st.session_state.aluno_dados['nome']).strip().upper()) & (df_banco['Turma'].astype(str).str.strip().str.upper() == str(st.session_state.aluno_dados['turma']).strip().upper())
                    
                    if mask.any():
                        df_banco.loc[mask, 'Nota'] = nota_final
                        df_banco.drop_duplicates(subset=['Aluno', 'Turma'], keep='last').to_csv("db_notas.csv", index=False)
            ja_entregou = False
            if os.path.exists("db_notas.csv"):
                df_c = pd.read_csv("db_notas.csv", on_bad_lines='skip')
                if not df_c.empty and 'Aluno' in df_c.columns and 'Turma' in df_c.columns:
                    mask_c = (df_c['Aluno'].astype(str).str.strip().str.upper() == str(st.session_state.aluno_dados['nome']).strip().upper()) & (df_c['Turma'].astype(str).str.strip().str.upper() == str(st.session_state.aluno_dados['turma']).strip().upper())
                    if mask_c.any():
                        ja_entregou = True
            if ja_entregou:
                st.error("🚨 Entrega bloqueada! Você já enviou sua avaliação anteriormente e o sistema permite apenas um envio único.")
            else:
                validos = [f for f in up if f.name.split('.')[0].lower() == nome_e.lower()]
                if validos and len(validos) == len(up):
                    notas_calculadas = []
                    infos_detalhados = []
                    for arq in validos:
                        st.session_state.aluno_arquivo_nome = arq.name
                        nota_parcial, info_parcial = calcular_nota(arq, st.session_state.aluno_dados['nome'])
                        notas_calculadas.append(nota_parcial)
                        infos_detalhados.append(f"Arquivo [{arq.name}]:\n{info_parcial}")
                    nota_final = round(sum(notas_calculadas) / len(notas_calculadas), 1)
                    info = "\n\n".join(infos_detalhados)
                    if os.path.exists("db_notas.csv"):
                        df_banco = pd.read_csv("db_notas.csv", on_bad_lines='skip')
                        mask = (df_banco['Aluno'].astype(str).str.strip().str.upper() == str(st.session_state.aluno_dados['nome']).strip().upper()) & (df_banco['Turma'].astype(str).str.strip().str.upper() == str(st.session_state.aluno_dados['turma']).strip().upper())
                        if mask.any():
                            df_banco.loc[mask, 'Nota'] = nota_final
                            df_banco.drop_duplicates(subset=['Aluno', 'Turma'], keep='last').to_csv("db_notas.csv", index=False)
                        else:
                            pd.DataFrame([[st.session_state.aluno_dados['nome'], st.session_state.aluno_dados['turma'], nota_final]], columns=['Aluno','Turma','Nota']).to_csv("db_notas.csv", mode='a', header=False, index=False)
                    else:
                        pd.DataFrame([[st.session_state.aluno_dados['nome'], st.session_state.aluno_dados['turma'], nota_final]], columns=['Aluno','Turma','Nota']).to_csv("db_notas.csv", mode='a', header=False, index=False)
                        pd.DataFrame([[st.session_state.aluno_dados['nome'], st.session_state.aluno_dados['turma'], nota_final]], columns=['Aluno','Turma','Nota']).to_csv("db_notas.csv", mode='w', header=True, index=False)
                    corpo_email = f"Resultado SENAI Excel\n\nAluno: {st.session_state.aluno_dados['nome']}\nTurma: {st.session_state.aluno_dados['turma']}\nNota Final Consolidada: {nota_final}\n\nDetalhes da Correção:\n{info}"
                    enviar_email(EMAIL_PROFESSOR, f"Prova - {st.session_state.aluno_dados['nome']}", corpo_email, [(f.getvalue(), f.name) for f in up])
                    enviar_email(st.session_state.aluno_dados['email'], "Seu Resultado na Prova de Excel", corpo_email, [(f.getvalue(), f.name) for f in up])
                    st.success(f"Finalizado! Nota Final: {nota_final}")
                    st.write("### Detalhes da Correção:")
                    st.write(info)
                    st.write("📬 **Status do Envio:**")
                    st.write(f"* E-mail enviado com sucesso para {EMAIL_PROFESSOR}")
                    st.write(f"* E-mail enviado com sucesso para {st.session_state.aluno_dados['email']}")
                    st.balloons()
                else:
                    pd.DataFrame([[st.session_state.aluno_dados['nome'], st.session_state.aluno_dados['turma'], nota_final]], columns=['Aluno','Turma','Nota']).to_csv("db_notas.csv", mode='w', header=True, index=False)
                
                corpo_email = f"Resultado SENAI Excel\n\nAluno: {st.session_state.aluno_dados['nome']}\nTurma: {st.session_state.aluno_dados['turma']}\nNota Final Consolidada: {nota_final}\n\nDetalhes da Correção:\n{info}"
                
                enviar_email(EMAIL_PROFESSOR, f"Prova - {st.session_state.aluno_dados['nome']}", corpo_email, [(f.getvalue(), f.name) for f in up])
                enviar_email(st.session_state.aluno_dados['email'], "Seu Resultado na Prova de Excel", corpo_email, [(f.getvalue(), f.name) for f in up])
                
                st.success(f"Finalizado! Nota Final: {nota_final}")
                st.write("### Detalhes da Correção:")
                st.write(info)
                st.write("📬 **Status do Envio:**")
                st.write(f"* E-mail enviado com sucesso para {EMAIL_PROFESSOR}")
                st.write(f"* E-mail enviado com sucesso para {st.session_state.aluno_dados['email']}")
                st.balloons()
            else:
                st.error("Nome do arquivo incorreto. Por favor, envie o arquivo original conforme as instruções.")

                    st.error("Arquivo incorreto detectado! Certifique-se de anexar exclusivamente o(s) arquivo(s) gerado(s) para o seu nome e que correspondam exatamente ao seu usuário.")
 
# --- PAINEL GESTOR / ADMINISTRATIVO (Sidebar Retornada) ---
elif st.session_state.perfil == "admin":
    with st.sidebar:
@@ -440,7 +454,7 @@
            "Selecione a área desejada:",
            ["📊 Notas", "📝 Professores", "🛡️ Área ADM"]
        )

 
    st.subheader(opcao_admin)

    if opcao_admin == "📊 Notas":
@@ -461,7 +475,7 @@
                    st.warning("Nenhum dado encontrado para esta turma ou credenciais inválidas.")
        else:
            st.info("Nenhum professor cadastrado no sistema ainda.")

 
    elif opcao_admin == "📝 Professores":
        with st.form("c_p"):
            n_p = st.text_input("Nome")
@@ -472,17 +486,17 @@
                    "professores.csv", mode='a', header=not os.path.exists("professores.csv"), index=False
                )
                st.success("Professor cadastrado com sucesso!")

 
    elif opcao_admin == "🛡️ Área ADM":
        if st.text_input("Senha ADM", type="password") == "Celina2610$$":
            if os.path.exists("db_notas.csv"):
                db = pd.read_csv("db_notas.csv", on_bad_lines='skip')
                db['Aluno'] = db['Aluno'].astype(str).str.strip()
                db['Turma'] = db['Turma'].astype(str).str.strip().str.upper()
                db_grouped = db.groupby(['Aluno', 'Turma'], as_index=False)['Nota'].mean().round(1)
                st.dataframe(db_grouped, use_container_width=True)
                if st.button("Limpar Todos os Dados"):
                    os.remove("db_notas.csv")
                    st.rerun()
            else:
                st.info("O banco de dados de notas está vazio no momento.")
