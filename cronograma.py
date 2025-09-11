import streamlit as st
import pandas as pd

st.set_page_config(page_title="Gerador de Cronograma de Desembolso", page_icon="📊", layout="centered")

st.title("📊 Gerador de Cronograma de Desembolso")

st.write("Este aplicativo processa duas planilhas Excel e gera uma terceira planilha consolidada.")

# Pergunta 1
alteracao_edt = st.radio("Alguma alteração na EDT?", ("Sim", "Não"))

if alteracao_edt == "Sim":
    st.warning("⚠️ Altere a planilha **equivalencia_eap_orcamento.xlsx** antes de continuar.")
else:
    arquivo_equivalencia = st.file_uploader("📂 Envie a planilha **equivalencia_eap_orcamento.xlsx**", type=["xlsx"])

# Pergunta 2 (só aparece se respondeu NÃO na anterior)
if alteracao_edt == "Não" and arquivo_equivalencia is not None:
    atualizou_desembolso = st.radio("Atualizou a planilha 'desembolso.xlsx'?", ("Sim", "Não"))

    if atualizou_desembolso == "Não":
        st.warning("⚠️ Favor atualizar a planilha **desembolso.xlsx** antes de continuar.")
    else:
        arquivo_desembolso = st.file_uploader("📂 Envie a planilha **desembolso.xlsx**", type=["xlsx"])

        if arquivo_desembolso is not None:
            if st.button("🚀 Rodar"):
                try:
                    # ======================
                    # 1. Preparando Excel de Equalização
                    # ======================
                    dados = pd.read_excel(arquivo_equivalencia,
                                          header=0)

                    tarefas = dados.copy()
                    tarefas.columns = tarefas.columns.str.replace(" ", "_").str.replace("(", "", regex=False).str.replace(")", "", regex=False)
                    tarefas.columns = [
                        col if "_" not in col or col[0].isalpha() else "_".join(col.split("_")[1:] + [col.split("_")[0]])
                        for col in tarefas.columns
                    ]

                    tarefas = pd.wide_to_long(
                        tarefas,
                        stubnames=["Item_Orç", "Peso_Orç"],
                        i=["EDT"],
                        j="Num_Item",
                        sep="_",
                        suffix=r"\d+"
                    ).reset_index()

                    tarefas = tarefas.rename(columns={
                        "Item_Orç": "Codigo Serviço",
                        "Peso_Orç": "Peso"
                    })

                    tarefas = tarefas.dropna(subset=["Peso"]).reset_index(drop=True)
                    tarefas = tarefas[~tarefas["Peso"].astype(str).str.upper().isin(["X", "(NA)"])]
                    tarefas = tarefas.sort_values(by=["EDT","Num_Item"]).reset_index(drop=True)
                    tarefas = tarefas[["EDT", "Codigo Serviço", "Peso"]]
                    tarefas["EDT"] = tarefas["EDT"].astype(str)

                    # ======================
                    # 2. Preparando tabela desembolso
                    # ======================
                    orcamento = pd.read_excel(arquivo_desembolso)
                    orcamento = orcamento.rename(columns={"ITENS" : "Codigo Serviço",
                                    "SERVIÇOS" : "Serviços",
                                    "ORÇAMENTO" : "Orcamento",
                                    "DESEMBOLSOS REALIZADOS (R$)" : "Desembolso",
                                    "COMPROMETIDO" : "Comprometido",
                                    "ESTOQUE/ADIANTAMENTO" : "Estoque",
                                    "OCS EM ABERTO" : "OCs",
                                    "SALDO DE CONTRATO" : "Saldo",
                                    "ESTIMATIVA NO TERMINO" : "Estimativa"
                                    })

                    orcamento["Codigo Serviço"] = orcamento["Codigo Serviço"].astype(str)
                    # ======================
                    # 3. Merge tarefas e orçamento
                    # ======================
                    cronograma = pd.merge(tarefas, orcamento, how="left", on=["Codigo Serviço"])

                    # ======================
                    # 4. Calculando quantidades finais
                    # ======================
                    cronograma["Orcamento final"] = cronograma["Peso"] * cronograma["Orcamento"]
                    cronograma["Desembolso final"] = cronograma["Peso"] * cronograma["Desembolso"]
                    cronograma["Comprometido final"] = cronograma["Peso"] * cronograma["Comprometido"]
                    cronograma["Estoque final"] = cronograma["Peso"] * cronograma["Estoque"]
                    cronograma["OCs final"] = cronograma["Peso"] * cronograma["OCs"]
                    cronograma["Saldo final"] = cronograma["Peso"] * cronograma["Saldo"]
                    cronograma["Estimativa final"] = cronograma["Peso"] * cronograma["Estimativa"]

                    df_final = cronograma[["EDT","Orcamento final","Desembolso final", "Comprometido final","Estoque final","OCs final","Saldo final","Estimativa final"]].copy()
                    df_final = df_final.groupby("EDT", as_index=False).sum().sort_values(by="EDT").reset_index(drop=True)

                    edt = dados[["EDT", "Nome da Tarefa"]].copy()
                    df_final = pd.merge(edt, df_final, how="left", on=["EDT"])

                    df_final = df_final.rename(columns={
                                                        "Orcamento final":"ORÇAMENTO",
                                                        "Desembolso final":"DESEMBOLSOS",
                                                        "Comprometido final":"COMPROMETIDO",
                                                        "Estoque final":"ESTOQUE/ADIANTAMENTO",
                                                        "OCs final":"OCS EM ABERTO",
                                                        "Saldo final":"SALDO DE CONTRATO",
                                                        "Estimativa final":"ESTIMATIVA NO TERMINO"
                                                        })

                    # ======================
                    # 5. Exportando resultado
                    # ======================
                    output_filename = "cronograma_desembolso.xlsx"
                    df_final.to_excel(output_filename, index=False)

                    with open(output_filename, "rb") as f:
                        st.success("✅ Arquivo gerado com sucesso!")
                        st.download_button("📥 Baixar Resultado", f, file_name=output_filename)

                except Exception as e:
                    st.error(f"❌ Ocorreu um erro: {e}")
