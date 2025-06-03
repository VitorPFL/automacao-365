import json
import requests
import msal
import pandas as pd
import os


# Função para obter o token de acesso
def obter_token(client_id, authority, client_secret, scope):
    app = msal.ConfidentialClientApplication(
        client_id, authority=authority, client_credential=client_secret
    )

    result = app.acquire_token_silent(scope, account=None)

    if not result:
        print("Gerando um novo token de acesso...")
        result = app.acquire_token_for_client(scopes=scope)

    if "access_token" in result:
        return result["access_token"]
    else:
        print("Erro ao obter token:", result.get("error"))
        print(result.get("error_description"))
        print(result.get("correlation_id"))
        return None


# Função para buscar todos os usuários do Microsoft 365
def obter_usuarios_365(token):
    url = ("https://graph.microsoft.com/v1.0/users?"
           "$top=999&$select=displayName,mail,userPrincipalName,accountEnabled,id,jobTitle,"
           "department,officeLocation,mobilePhone,employeeId,employeeHireDate,manager"
           "onPremisesSyncEnabled,onPremisesSamAccountName,onPremisesSecurityIdentifier,givenName,surname,employeeType")

    headers = {"Authorization": f"Bearer {token}"}

    usuarios = []

    while url:
        response = requests.get(url, headers=headers)
        if response.status_code == 200:
            data = response.json()
            usuarios.extend(data.get("value", []))
            url = data.get("@odata.nextLink", None)  # Pega a próxima página de resultados, se existir
        else:
            print("Erro ao buscar usuários:", response.status_code, response.text)
            return []

    return usuarios


# Função para salvar os dados em Excel
def salvar_em_excel(usuarios, caminho_saida):
    if not usuarios:
        print("Nenhum usuário encontrado. O arquivo Excel não será gerado.")
        return

    # Criar um DataFrame com as novas colunas
    df = pd.DataFrame(usuarios, columns=[
        "userPrincipalName", "onPremisesSamAccountName", "employeeId", "displayName","givenName","surname", "mail",
        "accountEnabled", "id","manager",
        "jobTitle", "department", "officeLocation", "mobilePhone",
        "employeeHireDate", "onPremisesSyncEnabled","employeeType"
    ])

    # Obter o nome do supervisor, se disponível
    df["manager"] = [user.get("manager", {}).get("displayName", "") for user in usuarios]

    # Renomear colunas para facilitar a leitura
    df.rename(columns={
        "displayName": "Nome",
        "givenName": "Primeiro Nome",
        "surname": "Sobrenome",
        "mail": "E-mail",
        "userPrincipalName": "Login",
        "accountEnabled": "Conta Ativa",
        "id": "ID do Usuário",
        "jobTitle": "Cargo",
        "department": "Departamento 365",
        "officeLocation": "Localização 365",
        "mobilePhone": "Celular 365",
        "manager": "Supervisor 365",
        "companyName": "Empresa 365",
        "employeeId": "RE",
        "employeeType": "Descrição",
        "onPremisesSyncEnabled": "Sincronizado",
        "onPremisesSamAccountName": "Login AD",
        "onPremisesSecurityIdentifier": "SID"
    }, inplace=True)

    # Salvar em Excel
    df.to_excel(caminho_saida, index=False)
    print(f"Arquivo salvo com sucesso em: {caminho_saida}")


# Carregar configurações do arquivo JSON
with open("../01_Inventario/parameters.json", "r") as file:
    config = json.load(file)

# Obter token de acesso
token = obter_token(
    config["client_id"], config["authority"], config["secret"], config["scope"]
)

if token:
    # Obter usuários do Microsoft 365
    usuarios = obter_usuarios_365(token)

    # Caminho de saída do arquivo Excel
    caminho_saida = r"Caminho\Usuarios_365.xlsx"

    # Salvar os usuários em Excel
    salvar_em_excel(usuarios, caminho_saida)
else:
    print("Falha ao obter token de acesso.")