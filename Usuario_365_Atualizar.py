import json
import requests
import msal
import pandas as pd
import os


# Função para obter token
def obter_token(client_id, authority, client_secret, scope):
    app = msal.ConfidentialClientApplication(
        client_id, authority=authority, client_credential=client_secret
    )
    result = app.acquire_token_silent(scope, account=None)
    if not result:
        result = app.acquire_token_for_client(scopes=scope)
    return result.get("access_token", None)


# Função para obter ID e se é sincronizado do AD
def buscar_usuario_por_email(email, token):
    url = f"https://graph.microsoft.com/v1.0/users/{email}"
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json"
    }
    resposta = requests.get(url, headers=headers)

    if resposta.status_code == 200:
        dados = resposta.json()
        user_id = dados["id"]
        syncado = dados.get("onPremisesSyncEnabled", False)
        return user_id, syncado

    elif resposta.status_code == 404:
        print(f"❌ Usuário {email} não encontrado no Microsoft 365.")
        return None, None

    else:
        print(f"❌ Erro ao buscar usuário {email}: {resposta.status_code} - {resposta.text}")
        return None, None


# Função para obter ID do gerente
def buscar_id_gerente(email_gerente, token):
    url = f"https://graph.microsoft.com/v1.0/users/{email_gerente}?$select=id"
    headers = {"Authorization": f"Bearer {token}"}
    response = requests.get(url, headers=headers)
    if response.status_code == 200:
        return response.json()["id"]
    else:
        print(f"❌ Erro ao buscar gerente {email_gerente}: {response.status_code} - {response.text}")
        return None


# Função para atualizar os campos permitidos
def atualizar_usuario(user_id, dados, token):
    url = f"https://graph.microsoft.com/v1.0/users/{user_id}"
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json"
    }
    body = json.dumps(dados)
    response = requests.patch(url, headers=headers, data=body)
    if response.status_code == 204:
        return True
    else:
        print(f"❌ Erro ao atualizar usuário {user_id}: {response.status_code} - {response.text}")
        return False


# Função para definir gerente
def definir_gerente(user_id, manager_id, token):
    url = f"https://graph.microsoft.com/v1.0/users/{user_id}/manager/$ref"
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json"
    }
    body = json.dumps({
        "@odata.id": f"https://graph.microsoft.com/v1.0/users/{manager_id}"
    })
    response = requests.put(url, headers=headers, data=body)
    if response.status_code in [204, 200]:
        return True
    else:
        print(f"❌ Erro ao definir gerente: {response.status_code} - {response.text}")
        return False


# --- INÍCIO DO SCRIPT PRINCIPAL ---

# Carregar configurações
with open("../01_Inventario/parameters.json", "r") as file:
    config = json.load(file)

# Obter token
token = obter_token(config["client_id"], config["authority"], config["secret"], config["scope"])

if not token:
    print("❌ Falha ao obter token.")
    exit()

# Carregar Excel
caminho_excel = os.path.expanduser("~/Downloads/Atualização 365.xlsx")
df = pd.read_excel(caminho_excel)

# Iterar pelos usuários
for index, row in df.iterrows():
    email_usuario = row["E-mail"]
    nome = row.get("Primeiro Nome Dcolab", "")
    sobrenome = row.get("Sobrenome Dcolab", "")
    cargo = row.get("Funcao", "")
    departamento = row.get("Secao", "")
    empresa = row.get("Empresa", "")
    supervisor = row.get("E-mail Super", "")
    RE = row.get("RE", "")
    descricao = row.get("Descrição", "")


    print(f"🔄 Atualizando: {email_usuario}")

    user_id, syncado = buscar_usuario_por_email(email_usuario, token)
    if not user_id:
        continue

    if syncado:
        print(f"⚠️ Usuário {email_usuario} é sincronizado com AD local. Ignorando atualização.")
        continue

    # Montar corpo de atualização
    dados_para_atualizar = {}
    # if nome: dados_para_atualizar["givenName"] = nome
    # if sobrenome: dados_para_atualizar["surname"] = sobrenome
    # if cargo: dados_para_atualizar["jobTitle"] = cargo
    # if departamento: dados_para_atualizar["department"] = departamento
    # if empresa: dados_para_atualizar["companyName"] = empresa
    if RE: dados_para_atualizar["employeeId"] = str(RE)
    if descricao: dados_para_atualizar["employeeType"] = descricao

    if dados_para_atualizar:
        sucesso = atualizar_usuario(user_id, dados_para_atualizar, token)
        if sucesso:
            campos = ", ".join([f"{k}='{v}'" for k, v in dados_para_atualizar.items()])
            print(f"✅ Usuário {email_usuario} atualizado com sucesso. Campos alterados: {campos}")

    # # Atualizar gerente
    # if supervisor and supervisor == supervisor:
    #     manager_id = buscar_id_gerente(supervisor, token)
    #     if manager_id:
    #         sucesso_gerente = definir_gerente(user_id, manager_id, token)
    #         if sucesso_gerente:
    #             print(f"✅ Gerente definido com sucesso para {email_usuario}")

    print("✅ Fim da atualização.")
