import pandas as pd
import requests
import time

# Função para consultar a API
def consultar_cnpj(cnpj):
    url = f"https://www.receitaws.com.br/v1/cnpj/{cnpj}"
    headers = {'User-Agent': 'Mozilla/5.0'}
    response = requests.get(url, headers=headers)
    
    if response.status_code == 200:
        data = response.json()
        if data.get("status") == "ERROR":
            return None
        return {
            "Razão Social": data.get("nome"),
            "Situação": data.get("situacao"),
            "Data Abertura": data.get("abertura"),
            "UF": data.get("uf"),
            "Atividade Principal": data.get("atividade_principal", [{}])[0].get("text")
        }
    else:
        return None

# Lê a planilha
df = pd.read_excel("cnpjs.xlsx")

# Cria colunas para os dados
df["Razão Social"] = ""
df["Situação"] = ""
df["Data Abertura"] = ""
df["UF"] = ""
df["Atividade Principal"] = ""

# Consulta cada CNPJ
for i, row in df.iterrows():
    cnpj = str(row["CNPJ"]).replace(".", "").replace("/", "").replace("-", "")
    print(f"Consultando {cnpj}...")
    dados = consultar_cnpj(cnpj)
    
    if dados:
        df.at[i, "Razão Social"] = dados["Razão Social"]
        df.at[i, "Situação"] = dados["Situação"]
        df.at[i, "Data Abertura"] = dados["Data Abertura"]
        df.at[i, "UF"] = dados["UF"]
        df.at[i, "Atividade Principal"] = dados["Atividade Principal"]
    else:
        print(f"Erro ao consultar CNPJ: {cnpj}")
    
    time.sleep(1.5)  # Evita bloqueio da API

# Salva o resultado
df.to_excel("cnpjs_atualizados.xlsx", index=False)
print("Consulta finalizada e planilha salva.")
