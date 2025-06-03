# 🛠️ **Automação de Gestão de Usuários no Microsoft 365**

Este projeto foi desenvolvido com o objetivo de **organizar, padronizar e automatizar** a gestão de usuários no ambiente Microsoft 365 de uma empresa. A automação surgiu de uma necessidade real da equipe de TI, que enfrentava sérias dificuldades para manter controle sobre contas, acessos e dados dos colaboradores, especialmente com inconsistências entre o AD local e o 365.

---

## 🎯 Problemas resolvidos

Antes desta solução, a empresa enfrentava:

- Contas de colaboradores demitidos ainda ativas no Microsoft 365
- Licenças atribuídas a ex-funcionários, impedindo novos acessos e gerando custos desnecessários
- Inconsistência entre usuários do AD local e do Microsoft 365
- Informações incompletas ou incorretas nos perfis dos usuários
- Falta de padronização no processo de criação, atualização e desativação de usuários


---

## ✅ Funcionalidades Atuais

- 📥 **Extração estruturada** de todos os usuários ativos no Microsoft 365 para planilha Excel
- 🔒 **Desativação automática** de contas de usuários desligados
- 🧾 **Atualização de atributos** como cargo, setor, empresa, nome, RE, CPF, e-mail, etc.
- 📁 **Revisão de licenças e acessos** conforme políticas internas
- 🔄 **Integração com base de dados corporativa (TOTVS)** para garantir fidelidade das informações
- 🔐 **Uso de variáveis de ambiente** para proteger credenciais sensíveis

---

## 🌱 Melhorias Futuras

- 🆕 Criação automática de usuários a partir de chamados (integração com API de chamados)
- 🤝 Integração 100% automática: TOTVS → Chamado → Criação no Microsoft 365
- 📦 Consolidação dos scripts em uma única aplicação com interface (CLI ou Web)
- 📊 Geração de logs e dashboards para rastreamento de alterações no ambiente 365

---

## 🧠 Tecnologias Utilizadas

- Python 3.x
- `msal` – autenticação segura no Microsoft 365 via Microsoft Graph API
- `pandas` – manipulação de planilhas
- `openpyxl` – leitura e escrita em arquivos Excel
- Microsoft 365
- Microsoft Graph API
	
---

## 🚀 Como Usar

1. Instale as dependências:
   ```bash
   pip install -r requirements.txt

---

## 👨‍💻 Autor

Vitor Pires Ferreira Leite
Pleno de Dados | Especialista em Automação de Processos, Integração de Sistemas e Soluções Corporativas

💡 Experiência com:

- Python, VBA, SQL
- Power BI (criação de dashboards interativos)
- SharePoint, Power Apps, APIs REST
- Integração entre plataformas corporativas (Metadados,TOTVS, AD, Microsoft 365)
- Azure, Databricks

Meu Linkedin:
- [LinkedIn](https://www.linkedin.com/in/vitor-ferreira-leite/)
