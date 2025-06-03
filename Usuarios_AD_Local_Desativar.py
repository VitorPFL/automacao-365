import json
import requests
import msal
import pandas as pd
import os

# Função para obter o token do Microsoft Graph
def obter_token(client_id, authority, client_secret, scope):
    app = msal.ConfidentialClientApplication(
        client_id, authority=authority, client_credential=client_secret
    )
    result = app.acquire_token_silent(scope, account=None)
    if not result:
        result = app.acquire_token_for_client(scopes=scope)
    return result.get("access_token", None)

# Buscar usuário e verificar se é sincronizado
def buscar_usuario_por_email(email, token):
    url = f"https://graph.microsoft.com/v1.0/users/{email}"
    headers = {"Authorization": f"Bearer {token}"}
    resposta = requests.get(url, headers=headers)
    if resposta.status_code == 200:
        dados = resposta.json()
        return dados["id"], dados.get("onPremisesSyncEnabled", False)
    elif resposta.status_code == 404:
        print(f"❌ Usuário {email} não encontrado.")
        return None, None
    else:
        print(f"❌ Erro ao buscar usuário {email}: {resposta.status_code} - {resposta.text}")
        return None, None

# Desabilitar o usuário no Azure AD
def desabilitar_usuario(user_id, token):
    url = f"https://graph.microsoft.com/v1.0/users/{user_id}"
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json"
    }
    body = json.dumps({"accountEnabled": False})
    resposta = requests.patch(url, headers=headers, data=body)
    if resposta.status_code == 204:
        print(f"✅ Usuário {user_id} desabilitado com sucesso.")
    else:
        print(f"❌ Falha ao desabilitar usuário {user_id}: {resposta.status_code} - {resposta.text}")

# Remover todas as licenças de um usuário
def remover_licencas_usuario(user_id, token):
    # Primeiro busca todas as licenças atribuídas
    url_get = f"https://graph.microsoft.com/v1.0/users/{user_id}/licenseDetails"
    headers = {"Authorization": f"Bearer {token}"}
    resposta = requests.get(url_get, headers=headers)

    if resposta.status_code != 200:
        print(f"❌ Erro ao buscar licenças do usuário {user_id}: {resposta.status_code}")
        return

    dados = resposta.json()
    licencas = [licenca["skuId"] for licenca in dados.get("value", [])]

    if not licencas:
        print(f"ℹ️ Nenhuma licença atribuída ao usuário {user_id}.")
        return

    url_remove = f"https://graph.microsoft.com/v1.0/users/{user_id}/assignLicense"
    body = {
        "addLicenses": [],
        "removeLicenses": licencas
    }

    resposta_remocao = requests.post(url_remove, headers=headers, json=body)
    if resposta_remocao.status_code == 200:
        print(f"🧹 Licenças removidas do usuário {user_id} com sucesso.")
    else:
        print(f"❌ Falha ao remover licenças do usuário {user_id}: {resposta_remocao.status_code} - {resposta_remocao.text}")

# === INÍCIO ===

# Carregar configurações
with open("../01_Inventario/parameters.json", "r") as file:
    config = json.load(file)

# Obter token
token = obter_token(config["client_id"], config["authority"], config["secret"], config["scope"])
if not token:
    print("❌ Falha ao obter token.")
    exit()

# Carregar planilha Excel
caminho_excel = os.path.expanduser("~/Downloads/Atualização 365.xlsx")
df = pd.read_excel(caminho_excel)

usuarios_processados = 0

# Iterar pelos usuários
for index, row in df.iterrows():
    email = str(row.get("E-mail", "")).strip()
    situacao = str(row.get("Situacao", "")).strip().lower()

    if not email or situacao != "demitido":
        continue

    print(f"\n🔍 Verificando usuário: {email}")
    user_id, syncado = buscar_usuario_por_email(email, token)

    if not user_id:
        continue

    if syncado:
        print(f"⚠️ Usuário {email} é sincronizado com AD local. Ignorando.")
        continue

    desabilitar_usuario(user_id, token)
    remover_licencas_usuario(user_id, token)
    usuarios_processados += 1

if usuarios_processados == 0:
    print("\n⚠️ Nenhum usuário foi desativado. Verifique os nomes das colunas e valores na planilha.")
else:
    print(f"\n✅ Total de usuários desativados e limpos: {usuarios_processados}")
