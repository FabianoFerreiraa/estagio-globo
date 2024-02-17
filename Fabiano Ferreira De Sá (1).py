import pandas as pd

tabela = pd.read_excel("C:\\Users\\fabia\\Downloads\\pasta100.xlsx")

tabela.loc[tabela["Unnamed: 5"] == 10406, "Unnamed: 1"] = "A"
tabela.loc[tabela["Unnamed: 5"] == 10352, "Unnamed: 1"] = "B"
tabela.loc[tabela["Unnamed: 5"] == 10206, "Unnamed: 1"] = "C"
tabela.loc[tabela["Unnamed: 5"] == 10835, "Unnamed: 1"] = "D"

tabela.loc[tabela["Unnamed: 5"] == 10406, "Unnamed: 2"] = "novela"
tabela.loc[tabela["Unnamed: 5"] == 10352, "Unnamed: 2"] = "serie"
tabela.loc[tabela["Unnamed: 5"] == 10206, "Unnamed: 2"] = "novela"
tabela.loc[tabela["Unnamed: 5"] == 10835, "Unnamed: 2"] = "serie"

tabela["Unnamed: 6"] = pd.to_datetime(tabela["Unnamed: 6"], errors="coerce")
tabela["data_formatada"] = tabela["Unnamed: 6"].dt.strftime("%d/%m/%Y")
tabela = tabela.fillna("nenhum")
tabela = tabela.drop(["Unnamed: 6", "Unnamed: 3", "conteudo"], axis=1)
tabela = tabela.drop(0)
tabela.columns = ["conteudo", "categoria", "usuario", "id_conteudo", "horas_consumidas", "data"]
ordem_colunas = ["data","usuario","conteudo","id_conteudo", "categoria", "horas_consumidas"]
tabela = tabela[ordem_colunas]

mais_consumida = tabela.groupby("usuario")["categoria"].agg(lambda x: x.value_counts().idxmax()).reset_index()
mais_consumida.rename(columns={"categoria": "categoria_consumida"}, inplace=True)

tabela.sort_values(by=["usuario", "data"], ascending=[True, True], inplace=True)

primeiro_play_por_usuario = tabela.drop_duplicates(subset="usuario", keep="first")

resultado = primeiro_play_por_usuario.groupby("usuario")["conteudo"].agg(lambda x: x.value_counts().idxmax()).reset_index()
resultado.rename(columns={"conteudo": "conteudo_consumido"}, inplace=True)

tabela["horas_consumidas"] = tabela["horas_consumidas"].str.replace(',', '.', regex=True).astype(float)
resumo_categoria = tabela.groupby("categoria").agg({"horas_consumidas": "sum", "id_conteudo": "count"}).reset_index()
resumo_categoria.rename(columns={"horas_consumidas": "horas_consumidas", "id_conteudo": "total_plays"}, inplace=True)

novelas = tabela[tabela["categoria"] == "novela"]
novelas["data"] = pd.to_datetime(novelas["data"], errors="coerce")
novelas["ano"] = novelas["data"].dt.year
novelas["mes"] = novelas["data"].dt.month

ranking_novelas = novelas.groupby(["ano", "mes", "conteudo"])["horas_consumidas"].sum().reset_index()
ranking_novelas = ranking_novelas.sort_values(by=["ano", "mes", "horas_consumidas"], ascending=[True, True, False])

minutos_por_play = tabela.groupby("usuario")["horas_consumidas"].mean() * 60
minutos_por_play.rename("minutos_por_play", inplace=True)

print(tabela)
print("\nCategoria mais consumida por usuário:")
print(mais_consumida)
print("\nConteúdo mais consumido no primeiro play por usuário:")
print(resultado)
print("\nQuantidade de horas consumidas e plays por categoria:")
print(resumo_categoria)
print("\nRanking de novelas com mais horas consumidas por mês:")
print(ranking_novelas)
print("\nMinutos por play para cada usuário:")
print(minutos_por_play)