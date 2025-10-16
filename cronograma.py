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
    tarefas = pd.wide_to_long(
        tarefas,
        stubnames=["Item_Orç", "Peso_Orç"],
        i=["EDT"],
        j="Num_Item",
        sep="_",
        suffix=r'\d+'
    ).reset_index()

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
    
    return tarefas

def processar_desembolso(df_desembolso):
    """
    Prepara e renomeia as colunas da planilha de desembolso.
    """
    orcamento = df_desembolso.rename(columns={
        "ITENS": "Codigo Serviço",
        "SERVIÇOS": "Serviços",
    })
    orcamento["Codigo Serviço"] = orcamento["Codigo Serviço"].astype(str)
    return orcamento

def calcular_valores_finais(cronograma):
    """
    Calcula os valores finais multiplicando as colunas pelo peso.
    """
    colunas_para_calcular = cronograma.columns.tolist()
    colunas_para_calcular.remove("EDT")
    colunas_para_calcular.remove("Codigo Serviço")
    colunas_para_calcular.remove("Peso")

    for col in colunas_para_calcular:
        cronograma[f"{col} final"] = cronograma["Peso"] * cronograma[col]
    return cronograma

def to_excel(df):
    """Converte um DataFrame para um objeto BytesIO em formato Excel."""
    output = BytesIO()
    # 'with' garante que o writer será fechado corretamente
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Cronograma')
    processed_data = output.getvalue()
    return processed_data

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

# Verifica se os arquivos foram carregados e exibe prévias
if arquivo_equivalencia:
    df_eq_original = pd.read_excel(arquivo_equivalencia)
    with st.expander("🧐 Prévia do arquivo de Equivalência"):
        st.dataframe(df_eq_original.head())
        
if arquivo_desembolso:
    df_des_original = pd.read_excel(arquivo_desembolso)
    with st.expander("🧐 Prévia do arquivo de Desembolso"):
        st.dataframe(df_des_original.head())

# Lógica principal do aplicativo
if arquivo_equivalencia is not None and arquivo_desembolso is not None:
    if st.button("🚀 Gerar Cronograma", type="primary"):
        with st.spinner("Processando os dados... Por favor, aguarde."):
            try:
                # 1. Preparando Excel de Equalização
                tarefas = processar_equivalencia(df_eq_original)

                # 2. Preparando tabela desembolso
                orcamento = processar_desembolso(df_des_original)

                # 3. Merge tarefas e orçamento
                cronograma = pd.merge(tarefas, orcamento, how="left", on="Codigo Serviço")

                # 4. Calculando quantidades finais
                cronograma = calcular_valores_finais(cronograma)

                # Selecionando colunas finais para o resultado
                # queremos apenas EDT e as colunas finais calculadas
                colunas_finais = ["EDT"]
                colunas_finais += [col for col in cronograma.columns if col.endswith("final")]

                df_final = cronograma[colunas_finais].copy()
                
                # Agrupando e somando os resultados por EDT
                df_final = df_final.groupby("EDT", as_index=False).sum().sort_values(by="EDT").reset_index(drop=True)
                
                # Adicionando de volta o 'Nome da Tarefa' para referência
                edt = df_eq_original[["EDT", "Nome da Tarefa"]].drop_duplicates().copy()
                edt["EDT"] = edt["EDT"].astype(str)
                df_final = pd.merge(edt, df_final, how="left", on="EDT")

                # Renomeando colunas removendo "final"
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
                st.error(f"❌ Ocorreu um erro durante o processamento:")
                st.warning(
                    "**Dica:** Um erro comum ocorre quando os nomes das colunas nos arquivos Excel não correspondem "
                    "exatamente ao esperado. **Verifique com atenção as colunas do arquivo `desembolso.xlsx`** "
                    "as colunas devem ser:"
                    "ITENS, SERVIÇOS, ORÇAMENTO ORIGINAL, DESEMBOLSOS REALIZADOS,"
                    "COMPROMETIDO TOTAL, ESTOQUE SIGNIFICATIVO/ADIANTAMENTO, OCS EM ABERTO,"
                    "SALDO DE CONTRATO EM ABERTO, ESTIMATIVA NO TERMINO (ENT)"
                )
                st.error(f"Detalhes do erro: {e}")
else:
    st.info("Por favor, envie ambos os arquivos para continuar.")

