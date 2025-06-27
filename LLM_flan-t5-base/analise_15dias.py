# analise_15dias.py

# Análise de Aderência das Previsões de Comissão

"""
Este script compara as previsões de comissão geradas anteriormente (pelo script 'previsao_comissao.py')
com os dados reais atualizados dos primeiros 15 dias do mês.

Etapas:
1. Carregar as previsões dos 30 dias do mês anterior.
2. Carregar a base de dados histórica atualizada com os novos 15 dias.
3. Alinhar e comparar previsões vs. real para o período dos 15 dias.
4. Calcular métricas de aderência (MAE, MAPE).
5. Gerar interpretações com LLM sobre a performance da previsão.
6. Adicionar nova aba ao arquivo Excel existente com a análise.

Requisitos:
- pip install pandas scikit-learn openpyxl transformers torch
"""

# === CONFIGURAÇÃO INICIAL =====================================================
INPUT_DIR_DADOS_ATUALIZADOS = r'INSIRA_AQUI_O_CAMINHO_DOS_DADOS_ATUALIZADOS'
OUTPUT_DIR_PREVISAO = r'INSIRA_AQUI_O_CAMINHO_DO_DIRETORIO_DE_SAIDA'
NOME_ARQUIVO_PREVISAO = "previsao_30dias.xlsx"
MODEL_NAME = "google/flan-t5-base" # Modelo LLM para explicações

# =============================================================================
import os
import sys
import time
import getpass
import platform
import warnings
from glob import glob
from datetime import datetime, timedelta
import argparse

import pandas as pd
import numpy as np
from sklearn.metrics import mean_absolute_error, mean_squared_error, r2_score, root_mean_squared_error
from transformers import AutoTokenizer, AutoModelForSeq2SeqLM
import torch
from openpyxl import load_workbook

warnings.filterwarnings("ignore", category=UserWarning)

parser = argparse.ArgumentParser(add_help=False)
parser.add_argument("-i", "--input_atualizado", dest="cli_input_atualizado")
parser.add_argument("-o", "--output_previsao", dest="cli_output_previsao")
args, _ = parser.parse_known_args()

if args.cli_input_atualizado:
    INPUT_DIR_DADOS_ATUALIZADOS = args.cli_input_atualizado
if args.cli_output_previsao:
    OUTPUT_DIR_PREVISAO = args.cli_output_previsao

# Caminho completo para o arquivo de previsão gerado anteriormente
caminho_excel_previsao = os.path.join(OUTPUT_DIR_PREVISAO, NOME_ARQUIVO_PREVISAO)

if not os.path.isdir(INPUT_DIR_DADOS_ATUALIZADOS):
    raise FileNotFoundError(f"Pasta de entrada de dados atualizados não encontrada: {INPUT_DIR_DADOS_ATUALIZADOS}")
if not os.path.isfile(caminho_excel_previsao):
    raise FileNotFoundError(f"Arquivo de previsão não encontrado: {caminho_excel_previsao}. Execute 'previsao_comissao.py' primeiro.")

print("\n🧑 Usuário:", getpass.getuser())
print("💻 Máquina:", platform.node())
print("🐍 Python :", sys.version.split()[0])
print("🗕️ Início  :", datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
print("\nInput Dados Atualizados:", INPUT_DIR_DADOS_ATUALIZADOS)
print("Output Previsão        :", OUTPUT_DIR_PREVISAO)
print("Arquivo de Previsão    :", NOME_ARQUIVO_PREVISAO)
print("="*60, "\n")

# === Carregar LLM para Geração de Explicações ================================
print(f"🧠 Carregando LLM ({MODEL_NAME}) para geração de explicações...")
tokenizer = AutoTokenizer.from_pretrained(MODEL_NAME)
model_llm = AutoModelForSeq2SeqLM.from_pretrained(MODEL_NAME)
if torch.cuda.is_available():
    model_llm.to('cuda')
    print("   • LLM carregada para GPU.")
else:
    print("   • LLM carregada para CPU.")

def generate_llm_explanation(prompt_text):
    inputs = tokenizer(prompt_text, return_tensors="pt", max_length=512, truncation=True)
    if torch.cuda.is_available():
        inputs = {k: v.to('cuda') for k, v in inputs.items()}
    with torch.no_grad():
        outputs = model_llm.generate(
            **inputs,
            max_new_tokens=150, # Aumentei um pouco para explicações mais completas
            num_beams=5,
            early_stopping=True
        )
    return tokenizer.decode(outputs[0], skip_special_tokens=True)

# === 1. Carregar Previsões e Dados Reais Atualizados =========================
print("🔍 Carregando previsões anteriores…")
try:
    df_previsoes_anteriores = pd.read_excel(caminho_excel_previsao, sheet_name="Previsao_30_Dias")
    df_previsoes_anteriores["Data"] = pd.to_datetime(df_previsoes_anteriores["Data"])
    # Convertendo 'Previsão Comissão' de volta para float para cálculos
    df_previsoes_anteriores["Previsão Comissão"] = df_previsoes_anteriores["Previsão Comissão"]\
        .str.replace(r"R\$", "", regex=True)\
        .str.replace(".", "", regex=False)\
        .str.replace(",", ".", regex=False)\
        .astype(float)
    print(f"✅ Previsões anteriores carregadas. {len(df_previsoes_anteriores)} registros.")
except Exception as e:
    raise FileNotFoundError(f"Erro ao carregar ou processar a aba 'Previsao_30_Dias' do arquivo {caminho_excel_previsao}: {e}")


print("🔍 Varredura de dados históricos atualizados…")
arquivos_xlsx_atualizados = glob(os.path.join(INPUT_DIR_DADOS_ATUALIZADOS, "*.xlsx"))
if not arquivos_xlsx_atualizados:
    raise FileNotFoundError("Nenhum .xlsx encontrado na pasta de entrada de dados atualizados.")

lista_df_atualizados = [pd.read_excel(arq) for arq in arquivos_xlsx_atualizados]
df_real_atualizado = pd.concat(lista_df_atualizados, ignore_index=True)

# Validar e pré-processar dados atualizados
colunas_obrigatorias = ["Data de Emissão", "Valor Comissão", "Seguradora", "Produto"]
if not all(col in df_real_atualizado.columns for col in colunas_obrigatorias):
    raise ValueError(f"As colunas {colunas_obrigatorias} são obrigatórias nos dados atualizados.")

df_real_atualizado["Data de Emissão"] = pd.to_datetime(df_real_atualizado["Data de Emissão"])
df_real_atualizado["Seguradora"] = df_real_atualizado["Seguradora"].astype(str)
df_real_atualizado["Produto"] = df_real_atualizado["Produto"].astype(str)
df_real_atualizado["Valor Comissão"] = pd.to_numeric(df_real_atualizado["Valor Comissão"], errors='coerce')
df_real_atualizado.dropna(subset=["Valor Comissão"], inplace=True)

print(f"✅ Dados reais atualizados carregados. Total de {len(df_real_atualizado)} registros.")

# === 2. Alinhar e Comparar Previsões vs. Real (Primeiros 15 Dias) ============
print("\n🔄 Realizando comparação dos primeiros 15 dias…")

# Determinar o período dos 15 dias do mês corrente
# A data de execução do script (hoje) é usada como referência para os "15 dias"
data_hoje = datetime.now().date()
data_inicio_mes_atual = data_hoje.replace(day=1)
data_fim_comparacao = data_hoje # Assumimos que 'hoje' é o 15º dia ou o dia em que se roda a análise

# Filtrar dados reais para o período
df_reais_periodo = df_real_atualizado[
    (df_real_atualizado["Data de Emissão"].dt.date >= data_inicio_mes_atual) &
    (df_real_atualizado["Data de Emissão"].dt.date <= data_fim_comparacao)
].groupby(["Seguradora", "Produto", df_real_atualizado["Data de Emissão"].dt.date.rename("Data_Real")])["Valor Comissão"].sum().reset_index()

# Filtrar previsões para o mesmo período
df_prev_periodo = df_previsoes_anteriores[
    (df_previsoes_anteriores["Data"] >= data_inicio_mes_atual) &
    (df_previsoes_anteriores["Data"] <= data_fim_comparacao)
]

# Unir os DataFrames para comparação
df_comparacao = pd.merge(
    df_reais_periodo,
    df_prev_periodo,
    left_on=["Seguradora", "Produto", "Data_Real"],
    right_on=["Seguradora", "Produto", "Data"],
    how="inner" # Apenas seguradoras/produtos/datas que têm dados nos dois
)

if df_comparacao.empty:
    print("⚠️ Nenhuma data coincidente encontrada para comparação nos primeiros 15 dias. Verifique os dados e o período.")
    df_analise_15dias = pd.DataFrame(columns=["Seguradora", "Produto", "MAE (15 dias)", "MAPE (15 dias)", "Total Real (R$)", "Total Previsto (R$)", "Interpretação da LLM", "Status"])
else:
    # Calcular métricas de erro para o período de comparação
    df_comparacao["Erro Absoluto"] = abs(df_comparacao["Valor Comissão"] - df_comparacao["Previsão Comissão"])
    # MAPE: Mean Absolute Percentage Error (evitar divisão por zero)
    df_comparacao["Erro Percentual"] = np.where(
        df_comparacao["Valor Comissão"] != 0,
        (df_comparacao["Erro Absoluto"] / df_comparacao["Valor Comissão"]) * 100,
        0 # Se o valor real for 0, o erro percentual é 0 se a previsão também for 0, ou alto se não for.
          # Para simplificar, estamos tratando como 0 aqui. Uma abordagem mais robusta seria log(x+1).
    )

    analise_15dias_resultados = []
    for (seg, prod), grupo_comp in df_comparacao.groupby(["Seguradora", "Produto"]):
        mae_15dias = grupo_comp["Erro Absoluto"].mean()
        mape_15dias = grupo_comp["Erro Percentual"].mean()
        total_real = grupo_comp["Valor Comissão"].sum()
        total_previsto = grupo_comp["Previsão Comissão"].sum()

        # Determinar status baseado no MAPE (exemplo de regra)
        if mape_15dias < 10:
            status = "Excelente"
        elif mape_15dias < 20:
            status = "Bom"
        elif mape_15dias < 35:
            status = "Moderado, requer atenção"
        else:
            status = "Previsão distante, requer investigação"

        # Gerar interpretação com LLM
        prompt_llm_15dias = (
            f"Análise de Desempenho da Previsão para os primeiros dias do mês atual:\n"
            f"Seguradora: {seg}, Produto: {prod}\n"
            f"Valor Real Total (período): R${total_real:,.2f}\n"
            f"Valor Previsto Total (período): R${total_previsto:,.2f}\n"
            f"Erro Absoluto Médio (MAE): R${mae_15dias:,.2f}\n"
            f"Erro Percentual Absoluto Médio (MAPE): {mape_15dias:.2f}%\n\n"
            f"Com base nessas métricas, forneça uma análise concisa sobre a aderência da previsão aos dados reais para este período."
            f"Indique se a previsão foi satisfatória ou se há desvios significativos e possíveis razões."
        )
        interpretação_llm_15dias = generate_llm_explanation(prompt_llm_15dias)

        analise_15dias_resultados.append({
            "Seguradora": seg,
            "Produto": prod,
            "MAE (15 dias)": round(mae_15dias, 2),
            "MAPE (15 dias)": round(mape_15dias, 2),
            "Total Real (R$)": round(total_real, 2),
            "Total Previsto (R$)": round(total_previsto, 2),
            "Interpretação da LLM": interpretação_llm_15dias,
            "Status": status
        })

    df_analise_15dias = pd.DataFrame(analise_15dias_resultados)

# === 3. Salvar Análise no Excel ==============================================
print("\n📄 Salvando aba 'Analise_15_Dias' no Excel…")

with pd.ExcelWriter(caminho_excel_previsao, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
    df_analise_15dias.to_excel(writer, sheet_name="Analise_15_Dias", index=False)

# === FINAL ===================================================================
print("\n✅ PROCESSAMENTO CONCLUÍDO: Análise dos 15 Dias")
print("📁 Excel atualizado:", caminho_excel_previsao)
print("⏱️ Tempo total     : {:.2f}s".format(time.time() - data_inicio_exec))
print("="*60, "\n")