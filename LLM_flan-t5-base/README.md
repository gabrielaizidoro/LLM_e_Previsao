# 📈 LLM_Previsao — Dashboard com Modelagem Preditiva para Valor de Comissão

Este projeto foca na **análise preditiva e visualização interativa** do valor de comissão a partir de dados históricos, utilizando modelos estatísticos e uma **Large Language Model (LLM) local** para explicações das métricas de desempenho.

---

## 🎯 Objetivo

O principal objetivo deste projeto é fornecer uma ferramenta para:
- Prever o valor de comissão para os próximos 30 dias por seguradora e produto.
- Visualizar o histórico e as previsões em gráficos interativos.
- Fornecer explicações compreensíveis sobre a performance dos modelos preditivos através de uma LLM.
- Gerar relatórios consolidados em formato Excel.

---

## 🛠️ Tecnologias e Bibliotecas

Este projeto utiliza as seguintes tecnologias:

-   **Python 3.8+**
-   **Pandas**: Manipulação e análise de dados.
-   **NumPy**: Suporte a operações numéricas.
-   **Scikit-learn**: Implementação de modelos de Machine Learning (Regressão Linear, Random Forest Regressor) e cálculo de métricas.
-   **Transformers (Hugging Face)**: Carregamento e uso do modelo LLM `google/flan-t5-base` para geração de texto.
-   **PyTorch**: Framework subjacente para a LLM.
-   **Matplotlib**: Geração de gráficos de histórico e previsão.
-   **Openpyxl**: Manipulação avançada de arquivos Excel para formatação de saída.
-   **Argparse**: Para argumentos de linha de comando.

---

## 🚀 Como Executar

Siga os passos abaixo para configurar e executar o projeto:

1.  **Clone o Repositório Principal:**
    ```bash
    git clone [https://github.com/gabrielaizidoro/LLM_e_Pevisao.git](https://github.com/gabrielaizidoro/LLM_e_Pevisao.git)
    cd Layout_txt/LLM_Previsao
    ```
    
2.  **Crie e Ative o Ambiente Virtual:**
    É altamente recomendado usar um ambiente virtual para isolar as dependências do projeto.
    ```bash
    python -m venv venv
    ```
    *No Windows:*
    ```bash
    .\venv\Scripts\activate
    ```
    *No macOS/Linux:*
    ```bash
    source venv/bin/activate
    ```

3.  **Instale as Dependências:**
    Com o ambiente virtual ativado, instale todas as bibliotecas necessárias:
    ```bash
    pip install pandas scikit-learn transformers torch openpyxl matplotlib
    ```
    *(Se preferir, crie um `requirements.txt` com essas dependências e use `pip install -r requirements.txt`)*

4.  **Prepare os Dados de Entrada:**
    - Crie uma pasta `Input/1_dia` dentro da pasta `LLM_Previsao` (ou ajuste o `INPUT_DIR` no script).
    - Coloque seus arquivos Excel (`.xlsx`) com os dados históricos de apólices nesta pasta.
    - **Certifique-se de que os arquivos Excel contêm as colunas:**
        - `Data de Emissão`
        - `Valor Comissão`
        - `Seguradora`
        - `Produto`
    - Os dados de "Data de Emissão" devem ser formatados como datas.

5.  **Configure os Diretórios (Opcional):**
    Abra `previsao_comissao.py` e ajuste `INPUT_DIR` e `OUTPUT_DIR` se necessário. Por padrão, ele espera `Input/1_dia` e criará `Output/previsao_inicial`.

6.  **Execute o Script:**
    A partir da pasta `LLM_Previsao`, execute o script principal:
    ```bash
    python previsao_comissao.py
    ```
    Você também pode passar os diretórios de entrada e saída como argumentos:
    ```bash
    python previsao_comissao.py -i "C:/caminho/para/seus/dados" -o "C:/caminho/para/seus/resultados"
    ```

---

## 📊 Saída Gerada

Ao final da execução, os resultados serão salvos no `OUTPUT_DIR` especificado:

-   **`previsao_llm_30dias.xlsx`**: Um arquivo Excel contendo:
    -   **`Previsao_30_Dias`**: Aba com as previsões de comissão para os próximos 30 dias.
    -   **`Resumo_Explicativo`**: Aba com as métricas de performance dos modelos (MAE, RMSE, R²) e suas interpretações geradas pela LLM.
-   **`graficos/`**: Um subdiretório contendo imagens (`.png`) dos gráficos de histórico e previsão para cada combinação de seguradora e produto.

---