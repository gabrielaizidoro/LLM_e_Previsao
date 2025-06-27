# 📦 Layout_txt — Repositório de Projetos para Geração, Validação e Visualização de Arquivos Posicionais

Este repositório centraliza projetos Python voltados à **automação de layouts posicionais (.txt)**, **validação de arquivos** e **monitoramento visual com previsão**, com foco em aplicações contábeis e de seguros.

---

## 📁 Projetos Incluídos

### 1. **🔍 Separador de Chaves EMI**
> Extrai linhas específicas de arquivos `.txt` com base em chaves de referência definidas em planilha Excel.

- 📥 Lê vários arquivos `.txt` em um diretório
- 🔎 Localiza chaves entre a posição 3 a 15 da linha
- 📊 Compara com chaves de uma planilha `.xlsx`
- ✅ Gera um novo `.txt` com as linhas válidas
- ⚠️ Relatório de chaves não localizadas

📂 Pasta: `Separador de Chaves EMI`  
📄 [README completo](./Separador%20de%20Chaves%20EMI/README.md)

---

### 2. **🧾 Cria EMI — Gerador de Arquivo TXT Posicional**
> Gera arquivos `.txt` com layout posicional fixo, a partir de uma planilha com dados de apólices.

- 🧠 Validação de tipo e tamanho dos campos
- 🔄 Formatação dos dados em layout fixo
- 📄 Exporta arquivo `ARQEMISS.txt`
- ⚠️ Gera relatório de erros por linha

📂 Pasta: `Cria EMI`  
📄 [README completo](./Cria%20EMI/README.md)

---

### 3. **📊 Boletim de Envios — Verificação de Arquivos por Data**
> Automatiza a verificação de recebimento de arquivos por seguradora e gera um boletim diário.

- 📅 Verifica arquivos com base na data de modificação (D-1)
- 📈 Tabela com status por seguradora e data
- 📧 Geração de e-mail com status consolidado

📂 Pasta: `Envio_Boletim_Automatico`  
📄 [README completo](./Envio_Boletim_Automatico/README.md)

---

### 4. **📈 LLM_Previsao — Dashboard com Modelagem Preditiva**
> Cria gráficos interativos com dados reais e previsões de volume, usando regressão e visualização por produto e seguradora.

- 🤖 Modelos preditivos via `sklearn` (Linear/Random Forest)
- 📉 Gráficos reais e previstos por seguradora/produto
- 💡 Dash simples com `matplotlib` e previsão integrada

📂 Pasta: `LLM_Previsao`  
📄 [README completo](./LLM_Previsao/README.md)

---

### 5. **🧪 Valida Erro Layout — Validador de Arquivos Posicionais**
> Valida arquivos `.txt` conforme regras de tipo, tamanho e obrigatoriedade, com geração de relatório de erros.

- 🔍 Checa integridade de cada linha
- 🧾 Confere formatos esperados por campo
- 🛠️ Gera arquivo de erros com detalhes por linha e campo

📂 Pasta: `Valida Erro layout`  
📄 [README completo](./Valida%20Erro%20layout/README.md)

---

## 🛠️ Requisitos

- Python 3.8+
- Uso recomendado com `venv`
- Instalação de dependências via `requirements.txt` em cada projeto

---

## 🚀 Execução

Cada projeto possui seu próprio `README.md` com instruções para:

- Ativação do ambiente virtual
- Instalação das dependências
- Execução dos scripts principais

---

## 👩‍💻 Autora

**Gabriela Izidoro**  
Automação • Dados • Processos Contábeis  
[github.com/gabrielaizidoro](https://github.com/gabrielaizidoro)

---