import streamlit as st
import pandas as pd
from io import BytesIO

# --- Configurações da Página ---
st.set_page_config(
    page_title="Gerador de Cronograma de Desembolso",
    page_icon="📊",
    layout="centered"
)

# --- Funções de Processamento de Dados ---

def processar_equivalencia(df_equivalencia):
    """
    Prepara e transforma a planilha de equivalência.
    Converte o formato 'wide' para 'long'.
    """
    tarefas = df_equivalencia.copy()
    # Limpa e padroniza os nomes das colunas
    tarefas.columns = tarefas.columns.str.replace(" ", "_").str.replace(r'[\(\)]', '', regex=True)
    
    # Reorganiza colunas que começam com números (ex: '1_Item_Orç' para 'Item_Orç_1')
    # para funcionar com a função wide_to_long.
    novas_colunas = []
    for col in tarefas.columns:
        if "_" in col and col.split("_")[0].isdigit():
            partes = col.split("_")
            novo_nome = "_".join(partes[1:]) + "_" + partes[0]
            novas_colunas.append(novo_nome)
        else:
            novas_colunas.append(col)
    tarefas.columns = novas_colunas
    
    # Transforma a tabela de formato largo para longo
    try:
        tarefas = pd.wide_to_long(
            tarefas,
            stubnames=["Item_Orç", "Peso_Orç"],
            i=["EDT"],
            j="Num_Item",
            sep="_",
            suffix=r'\d+'
        ).reset_index()
    except ValueError as e:
        st.error(f"Erro ao reorganizar a planilha de equivalência. Verifique se as colunas seguem o padrão '1_Item_Orç', '1_Peso_Orç', etc.")
        st.error(f"Detalhe do erro do Pandas: {e}")
        return None

    tarefas = tarefas.rename(columns={
        "Item_Orç": "Codigo Serviço",
        "Peso_Orç": "Peso"
    })

    # Limpa os dados, removendo linhas sem peso ou com valores inválidos
    tarefas = tarefas.dropna(subset=["Peso"])
    tarefas = tarefas[~tarefas["Peso"].astype(str).str.upper().isin(["X", "(NA)"])]
    tarefas = tarefas.sort_values(by=["EDT", "Num_Item"]).reset_index(drop=True)
    
    # Seleciona colunas finais e ajusta tipos
    tarefas = tarefas[["EDT", "Codigo Serviço", "Peso"]]
    tarefas["EDT"] = tarefas["EDT"].astype(str)
    tarefas["Codigo Serviço"] = tarefas["Codigo Serviço"].astype(str)
    
    # --- CORREÇÃO ---
    # Converte a coluna 'Peso' para um tipo numérico (float).
    # O parâmetro errors='coerce' transformará qualquer valor que não seja numérico em NaN (Not a Number).
    # Em seguida, .fillna(0.0) substitui esses NaN por 0.0.
    # Isso garante que a coluna 'Peso' contenha apenas números, evitando o erro de multiplicação.
    tarefas["Peso"] = pd.to_numeric(tarefas["Peso"], errors='coerce').fillna(0.0)
    
    return tarefas

def processar_desembolso(df_desembolso):
    """
    Prepara e renomeia as colunas da planilha de desembolso, garantindo que os tipos de dados estejam corretos.
    """
    orcamento = df_desembolso.copy()
    # Renomeia as colunas de identificação
    orcamento = orcamento.rename(columns={
        "ITENS": "Codigo Serviço",
        "SERVIÇOS": "Serviços",
    })
    
    if "Codigo Serviço" not in orcamento.columns:
        st.error("A planilha de desembolso precisa ter uma coluna chamada 'ITENS'.")
        return None

    orcamento["Codigo Serviço"] = orcamento["Codigo Serviço"].astype(str)
    
    # --- CORREÇÃO ---
    # Itera sobre todas as colunas que não são de identificação para garantir que sejam numéricas.
    colunas_de_valor = [col for col in orcamento.columns if col not in ["Codigo Serviço", "Serviços"]]
    for col in colunas_de_valor:
        # Aplica a mesma lógica da função anterior para garantir que os dados de desembolso são numéricos.
        orcamento[col] = pd.to_numeric(orcamento[col], errors='coerce').fillna(0.0)
        
    return orcamento

def calcular_valores_finais(cronograma):
    """
    Calcula os valores finais multiplicando as colunas de desembolso pelo peso.
    """
    # Identifica as colunas de desembolso (que não são colunas de identificação)
    colunas_para_calcular = [
        col for col in cronograma.columns 
        if col not in ["EDT", "Codigo Serviço", "Peso", "Nome da Tarefa", "Serviços"]
    ]
    
    # --- CORREÇÃO PREVENTIVA ---
    # Após a junção (merge), algumas linhas podem ter valores nulos (NaN) se um serviço
    # existia na planilha de equivalência mas não na de desembolso.
    # Preenchemos esses valores nulos com 0.0 para garantir que os cálculos funcionem.
    cronograma[colunas_para_calcular] = cronograma[colunas_para_calcular].fillna(0.0)

    for col in colunas_para_calcular:
        # A multiplicação agora é segura, pois todas as colunas envolvidas são numéricas.
        cronograma[f"{col} final"] = cronograma["Peso"] * cronograma[col]
        
    return cronograma

def to_excel(df):
    """Converte um DataFrame para um objeto BytesIO em formato Excel."""
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Cronograma')
    # O método getvalue() é chamado após o bloco 'with' para garantir que tudo foi escrito.
    return output.getvalue()

# --- Interface do Streamlit ---

st.title("📊 Gerador de Cronograma de Desembolso")
st.write("Este aplicativo processa duas planilhas Excel e gera uma terceira planilha consolidada, pronta para download.")

st.markdown("---")

# Seção de Upload de Arquivos
col1, col2 = st.columns(2)
with col1:
    arquivo_equivalencia = st.file_uploader("1. Envie **equivalencia_eap_orcamento.xlsx**", type=["xlsx"])

with col2:
    arquivo_desembolso = st.file_uploader("2. Envie **desembolso.xlsx**", type=["xlsx"])

# Garante que os dataframes sejam lidos apenas uma vez
df_eq_original = None
df_des_original = None

if arquivo_equivalencia:
    df_eq_original = pd.read_excel(arquivo_equivalencia)
    with st.expander("🧐 Prévia do arquivo de Equivalência"):
        st.dataframe(df_eq_original.head())
        
if arquivo_desembolso:
    df_des_original = pd.read_excel(arquivo_desembolso)
    with st.expander("🧐 Prévia do arquivo de Desembolso"):
        st.dataframe(df_des_original.head())

# Lógica principal do aplicativo
if df_eq_original is not None and df_des_original is not None:
    if st.button("🚀 Gerar Cronograma", type="primary"):
        with st.spinner("Processando os dados... Por favor, aguarde."):
            try:
                # 1. Preparando Excel de Equalização
                tarefas = processar_equivalencia(df_eq_original)
                if tarefas is None:
                    # A função já exibiu o erro, então apenas paramos a execução
                    st.stop()

                # 2. Preparando tabela desembolso
                orcamento = processar_desembolso(df_des_original)
                if orcamento is None:
                    st.stop()

                # 3. Merge tarefas e orçamento
                cronograma = pd.merge(tarefas, orcamento, how="left", on="Codigo Serviço")

                # 4. Calculando quantidades finais
                cronograma = calcular_valores_finais(cronograma)

                # Selecionando colunas finais para o resultado
                colunas_finais = ["EDT"]
                colunas_finais += [col for col in cronograma.columns if col.endswith("final")]

                df_final = cronograma[colunas_finais].copy()
                
                # Agrupando e somando os resultados por EDT
                df_final = df_final.groupby("EDT", as_index=False).sum().sort_values(by="EDT").reset_index(drop=True)
                
                # Adicionando de volta o 'Nome da Tarefa' para referência
                if "Nome da Tarefa" in df_eq_original.columns:
                    edt = df_eq_original[["EDT", "Nome da Tarefa"]].drop_duplicates().copy()
                    edt["EDT"] = edt["EDT"].astype(str)
                    df_final = pd.merge(edt, df_final, how="left", on="EDT")
                else:
                    st.warning("Coluna 'Nome da Tarefa' não encontrada no arquivo de equivalência. O resultado será gerado sem ela.")

                # Renomeando colunas removendo " final"
                df_final = df_final.rename(columns=lambda x: x.replace(" final", "") if x.endswith(" final") else x)

                # 5. Exportando resultado
                excel_file = to_excel(df_final)
                
                st.success("✅ Arquivo gerado com sucesso!")
                
                st.download_button(
                    label="📥 Baixar Resultado (cronograma_desembolso.xlsx)",
                    data=excel_file,
                    file_name="cronograma_desembolso.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

            except Exception as e:
                st.error(f"❌ Ocorreu um erro inesperado durante o processamento:")
                # st.exception é melhor para depuração, pois mostra todos os detalhes do erro.
                st.exception(e) 
else:
    st.info("Por favor, envie ambos os arquivos para continuar.")
