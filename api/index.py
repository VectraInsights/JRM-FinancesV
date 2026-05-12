from fastapi import FastAPI, Query
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import JSONResponse
import pandas as pd
import requests
from io import BytesIO
import time
from datetime import datetime
from zoneinfo import ZoneInfo
from typing import List, Optional

app = FastAPI()

# Configuração de CORS para permitir que seu frontend acesse a API
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

# =========================================================
# CONSTANTES E CONFIGURAÇÕES (Preservadas do original)
# =========================================================
SHEET_ID_2026 = "1Bh2Xb1t5m7Si3HXgRVHkrQ5ALtsBRV_wTTU1G-EQo0A"
SHEET_ID_2025 = "1JEruYIxHbwPlQS45oTFH60l3i4hqF_3idM2Vu8sFPHs"
TIMEZONE_BRASILIA = ZoneInfo("America/Sao_Paulo")

MESES = [
    "Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho",
    "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"
]
MESES_MAP = {mes: i + 1 for i, mes in enumerate(MESES)}
BLOCOS_DESPESA_EXATOS = {"DIVERSOS", "DESPESAS DIVERSAS", "ACESSÓRIOS", "ACESSORIOS", "INVESTIMENTOS"}
BLOCOS_PARADA_GERAIS = {"RESULTADO", "RECEITAS", "TOTAL", "FATURAMENTO"}

# =========================================================
# FUNÇÕES DE TRATAMENTO DE DADOS (Sua Lógica Original)
# =========================================================
def texto(v): return str(v).strip() if pd.notna(v) else ""
def numero(v): return pd.to_numeric(v, errors="coerce")

def limpar_prefixo_despesas(nome):
    t = texto(nome).strip()
    tu = t.upper()
    mapa = {
        "DESPESAS ADMINISTRATIVAS": "Administrativas",
        "DESPESAS DA OFICINA MECÂNICA": "Mecânica",
        "DESPESAS DE BORRACHARIA": "Borracharia",
        "DESPESAS DA ELÉTRICA": "Elétrica",
        "DESPESAS DIVERSAS": "Abastecimento/Descargas",
        "INVESTIMENTOS": "Financiamentos/Imobilizados",
    }
    if tu in mapa: return mapa[tu]
    return t

def eh_bloco_despesa(valor):
    t = texto(valor).upper()
    return t.startswith("DESPESAS") or t in BLOCOS_DESPESA_EXATOS

def processar_planilha(xl):
    receitas, despesas_res, despesas_det = [], [], []

    for aba in xl.sheet_names:
        if aba not in MESES: continue
        df = pd.read_excel(xl, sheet_name=aba, usecols="A:D", header=None, names=["ColA", "ColB", "ColC", "ColD"])
        df["ColA"], df["ColB"] = df["ColA"].apply(texto), df["ColB"].apply(texto)
        df["ColC"], df["ColD"] = df["ColC"].apply(numero), df["ColD"].apply(numero)

        for i in df.index:
            if not eh_bloco_despesa(df.loc[i, "ColA"]): continue
            setor = limpar_prefixo_despesas(df.loc[i, "ColA"])
            total_setor = df.loc[i + 1, "ColD"] if (i + 1) in df.index else 0
            if pd.notna(total_setor) and total_setor != 0:
                despesas_res.append({"Mes": aba, "Setor": setor, "Valor": abs(float(total_setor))})

            for j in range(i + 1, len(df)):
                bloco_j = texto(df.loc[j, "ColA"])
                if j > i + 1 and (eh_bloco_despesa(bloco_j) or bloco_j.upper() in BLOCOS_PARADA_GERAIS): break
                cat, val = limpar_prefixo_despesas(df.loc[j, "ColB"]), df.loc[j, "ColC"]
                if cat and pd.notna(val) and val != 0:
                    despesas_det.append({"Mes": aba, "Setor": setor, "Categoria": cat, "Valor": abs(float(val))})

        idx_fat = df[df["ColA"].str.upper() == "FATURAMENTO"].index
        if not idx_fat.empty:
            for k in range(idx_fat[0] + 1, len(df)):
                if texto(df.loc[k, "ColA"]).upper() in BLOCOS_PARADA_GERAIS or eh_bloco_despesa(df.loc[k, "ColA"]): break
                r_nome, r_val = texto(df.loc[k, "ColB"]), df.loc[k, "ColC"]
                if r_nome and pd.notna(r_val) and r_val > 0:
                    receitas.append({"Mes": aba, "Receita": r_nome, "Valor": float(r_val)})

    return pd.DataFrame(receitas), pd.DataFrame(despesas_res), pd.DataFrame(despesas_det)

def carregar_dados(sheet_id):
    ts = int(time.time())
    url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=xlsx&t={ts}"
    response = requests.get(url, timeout=60)
    response.raise_for_status()
    xl = pd.ExcelFile(BytesIO(response.content))
    return processar_planilha(xl)

# =========================================================
# ENDPOINTS DA API
# =========================================================

@app.get("/api/data")
async def get_financial_data(ano: str = "2026"):
    try:
        sheet_id = SHEET_ID_2026 if ano == "2026" else SHEET_ID_2025
        df_r, df_dr, df_dd = carregar_dados(sheet_id)

        # Dados para os filtros do frontend
        meses_disponiveis = df_r["Mes"].unique().tolist()
        
        # Enviamos os dados brutos para o frontend filtrar (conforme o Streamlit fazia)
        return {
            "status": "success",
            "meses_disponiveis": meses_disponiveis,
            "raw_receitas": df_r.to_dict(orient="records"),
            "raw_despesas_res": df_dr.to_dict(orient="records"),
            "raw_despesas_det": df_dd.to_dict(orient="records"),
            "ultima_atualizacao": datetime.now(TIMEZONE_BRASILIA).strftime("%d/%m/%Y %H:%M:%S")
        }
    except Exception as e:
        return JSONResponse(status_code=500, content={"status": "error", "message": str(e)})
