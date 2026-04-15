"""
api.py — Mapa de Contagem de Tráfego
Fotos servidas via Google Drive API (pastas públicas)
"""

import os, json, urllib.request, urllib.parse
from pathlib import Path
from datetime import datetime
from typing import Optional
from functools import lru_cache
import time

from fastapi import FastAPI, HTTPException, UploadFile, File
from fastapi.responses import JSONResponse, HTMLResponse
from fastapi.middleware.cors import CORSMiddleware

try:
    import pandas as pd
except ImportError:
    raise SystemExit("Instale: pip install pandas openpyxl")

# ── Configuração ──────────────────────────────────────────────────────────────
BASE_DIR  = Path(__file__).parent
HTML_PATH = BASE_DIR / "mapa.html"

def encontrar_excel():
    candidatos = [
        "PONTOS_DE_CONTAGEM_DE_TR\u00c1FEGO.xlsx",
        "PONTOS DE CONTAGEM DE TR\u00c1FEGO.xlsx",
        "PONTOS_DE_CONTAGEM_DE_TRAFEGO.xlsx",
        "PONTOS DE CONTAGEM DE TRAFEGO.xlsx",
    ]
    for nome in candidatos:
        p = BASE_DIR / nome
        if p.exists():
            return p
    xlsx = list(BASE_DIR.glob("*.xlsx"))
    if xlsx:
        return xlsx[0]
    return BASE_DIR / "PONTOS_DE_CONTAGEM_DE_TR\u00c1FEGO.xlsx"

EXCEL_PATH = encontrar_excel()

ABA_PONTOS      = "LISTA DE PONTOS"
ABA_DADOS       = "DADOS COLETADOS"

# Google Drive
DRIVE_API_KEY       = os.environ.get("DRIVE_API_KEY", "AIzaSyATY3YHzzMJTB92heZcIxIJBGpK5Y2Ff6A")
DRIVE_ROOT_FOLDER   = os.environ.get("DRIVE_ROOT_FOLDER", "1yrBYApnWeKEe7YEJT7I6Clc5JgeDFpCj")
DRIVE_BASE_URL      = "https://www.googleapis.com/drive/v3"
DRIVE_VIEW_URL      = "https://lh3.googleusercontent.com/d/"

# Cache de 10 minutos para não bater na API a cada clique
_cache_fotos: dict = {}
_cache_subpastas: dict = {}
_cache_timestamp: dict = {}
CACHE_TTL = 600  # segundos

CLASSES = ['MOTO','PASS','UTIL','2C/2CB','3C/3CB',
           '2S1','2S2','2S3','3S1','3S2','3S3',
           '4C','4CD','2C2','2C3','3C2','3C3','2I3','3I3','BIT','ROD/TRIT']

EXTS_FOTO = {'image/jpeg','image/png','image/webp','image/jpg'}

# ── App ───────────────────────────────────────────────────────────────────────
app = FastAPI(title="Mapa de Contagem de Tráfego", version="2.0")
app.add_middleware(CORSMiddleware, allow_origins=["*"], allow_methods=["*"], allow_headers=["*"])

# ── Google Drive helpers ──────────────────────────────────────────────────────
def drive_list(folder_id: str) -> list:
    """Lista arquivos/pastas dentro de uma pasta do Drive."""
    q = urllib.parse.quote(f"'{folder_id}' in parents and trashed=false")
    fields = "files(id,name,mimeType)"
    url = f"{DRIVE_BASE_URL}/files?q={q}&key={DRIVE_API_KEY}&fields={fields}&pageSize=100"
    try:
        resp = urllib.request.urlopen(url, timeout=10)
        data = json.loads(resp.read())
        return data.get("files", [])
    except Exception as e:
        print(f"Drive API erro: {e}")
        return []

def encontrar_pasta_id(parent_id: str, nome_busca: str) -> Optional[str]:
    """Encontra ID de uma subpasta pelo nome (case-insensitive, aceita variações)."""
    items = drive_list(parent_id)
    # Tenta match exato primeiro
    for item in items:
        if item["mimeType"] == "application/vnd.google-apps.folder":
            if item["name"].upper() == nome_busca.upper():
                return item["id"]
    # Tenta match parcial com "instal"
    for item in items:
        if item["mimeType"] == "application/vnd.google-apps.folder":
            if "instal" in item["name"].lower():
                return item["id"]
    # Se só tiver uma subpasta, usa ela
    pastas = [i for i in items if i["mimeType"] == "application/vnd.google-apps.folder"]
    if len(pastas) == 1:
        return pastas[0]["id"]
    return None

def listar_fotos_drive(pid: str) -> list:
    """Retorna lista de URLs diretas das fotos do ponto no Drive."""
    now = time.time()

    # Cache válido?
    if pid in _cache_fotos and (now - _cache_timestamp.get(pid, 0)) < CACHE_TTL:
        return _cache_fotos[pid]

    # 1. Encontra pasta do ponto (ex: "E 05") dentro da raiz
    if "root_items" not in _cache_subpastas or (now - _cache_timestamp.get("root", 0)) > CACHE_TTL:
        items = drive_list(DRIVE_ROOT_FOLDER)
        _cache_subpastas["root_items"] = {
            i["name"]: i["id"] for i in items
            if i["mimeType"] == "application/vnd.google-apps.folder"
        }
        _cache_timestamp["root"] = now

    pasta_ponto_id = _cache_subpastas["root_items"].get(pid)
    if not pasta_ponto_id:
        _cache_fotos[pid] = []
        _cache_timestamp[pid] = now
        return []

    # 2. Encontra subpasta FOTOS - INSTALAÇÃO
    sub_id = encontrar_pasta_id(pasta_ponto_id, "FOTOS - INSTALAÇÃO")
    if not sub_id:
        sub_id = pasta_ponto_id  # fotos direto na pasta do ponto

    # 3. Lista arquivos de imagem
    items = drive_list(sub_id)
    fotos = sorted(
        [i for i in items if i.get("mimeType","") in EXTS_FOTO],
        key=lambda x: x["name"].lower()
    )
    urls = [DRIVE_VIEW_URL + f["id"] for f in fotos]

    _cache_fotos[pid] = urls
    _cache_timestamp[pid] = now
    return urls

# ── Helpers Excel ─────────────────────────────────────────────────────────────
def limpar(v):
    return "" if (v is None or str(v).strip() == "nan") else str(v).strip()

def coord(v):
    try: return float(str(v).replace(",", ".").strip())
    except: return None

def fmt_data(v):
    if pd.isna(v): return None
    try:
        ts = pd.Timestamp(v)
        return None if ts.year < 1950 else ts.strftime("%d/%m/%Y")
    except: return None

def ler_excel():
    if not EXCEL_PATH.exists():
        raise HTTPException(404, f"Excel não encontrado: {EXCEL_PATH.name}")

    df = pd.read_excel(EXCEL_PATH, sheet_name=ABA_PONTOS, dtype=str)
    df.dropna(how="all", inplace=True)

    col_map = {c.upper().strip(): c for c in df.columns}
    def col(cands):
        for c in cands:
            if c.upper() in col_map: return col_map[c.upper()]
        return None

    c_id  = col(["ID"])
    c_lat = col(["LATITUDE","LAT"])
    c_lng = col(["LONGITUDE","LNG","LONG"])
    c_rod = col(["RODOVIA"])
    c_tre = col(["TRECHO"])
    c_ini = col(["DESCRIÇÃO INÍCIO","DESCRICAO INICIO","INICIO"])
    c_fim = col(["DESCRIÇÃO FINAL","DESCRICAO FINAL","FIM"])
    c_sug = col(["SUGESTÃO","SUGESTAO","RESPONSAVEL","RESPONSÁVEL"])
    c_sta = col(["STATUS DE EXECUÇÃO","STATUS DE EXECUCAO","STATUS"])
    c_pis = col(["SITUAÇÃO PISTA","SITUACAO PISTA","PISTA"])

    # Dados coletados
    dados_dc = {}
    try:
        dc = pd.read_excel(EXCEL_PATH, sheet_name=ABA_DADOS, header=None)
        for i in range(4, len(dc)):
            row = dc.iloc[i]
            pid = limpar(str(row[0])) if pd.notna(row[0]) else ""
            if not pid or pid == "nan": continue
            total = int(row[28]) if pd.notna(row[28]) and str(row[28]) not in ("nan","") else 0
            contagens = {}
            for j, cls in enumerate(CLASSES):
                val = row[7+j]
                contagens[cls] = int(val) if pd.notna(val) and str(val) not in ("nan","") else 0
            dados_dc[pid] = {
                "periodo_inicio": fmt_data(row[5]),
                "periodo_fim":    fmt_data(row[6]),
                "total": total,
                "contagens": contagens,
            }
    except Exception as e:
        print(f"Aviso dados coletados: {e}")

    pontos = []
    for _, row in df.iterrows():
        pid = limpar(row.get(c_id, "")) if c_id else ""
        if not pid: continue
        lat = coord(row.get(c_lat, "")) if c_lat else None
        lng = coord(row.get(c_lng, "")) if c_lng else None
        if lat is None or lng is None: continue

        def g(c): return limpar(row.get(c, "")) if c else ""
        dc = dados_dc.get(pid, {})

        pontos.append({
            "id":       pid,
            "lat":      lat,
            "lng":      lng,
            "rodovia":  g(c_rod),
            "trecho":   g(c_tre),
            "inicio":   g(c_ini),
            "fim":      g(c_fim),
            "sugestao": g(c_sug) or "-",
            "status":   g(c_sta) or "-",
            "pista":    g(c_pis) or "-",
            "fotos":    listar_fotos_drive(pid),
            "total":           dc.get("total", 0),
            "periodo_inicio":  dc.get("periodo_inicio"),
            "periodo_fim":     dc.get("periodo_fim"),
            "contagens":       dc.get("contagens", {}),
        })

    return pontos

# ── Rotas ─────────────────────────────────────────────────────────────────────
@app.get("/", response_class=HTMLResponse)
def raiz():
    if HTML_PATH.exists():
        return HTMLResponse(HTML_PATH.read_text(encoding="utf-8"))
    return HTMLResponse("<h2>mapa.html não encontrado</h2>")

@app.get("/api/pontos")
def get_pontos():
    pontos = ler_excel()
    return JSONResponse({
        "total": len(pontos),
        "com_contagem": sum(1 for p in pontos if p["total"] > 0),
        "com_fotos": sum(1 for p in pontos if p["fotos"]),
        "atualizado_em": datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
        "pontos": pontos,
    })

@app.get("/api/ponto/{pid}")
def get_ponto(pid: str):
    pontos = ler_excel()
    for p in pontos:
        if p["id"] == pid: return JSONResponse(p)
    raise HTTPException(404, f"Ponto '{pid}' não encontrado")

@app.post("/api/upload-excel")
async def upload_excel(file: UploadFile = File(...)):
    if not file.filename.endswith((".xlsx",".xls")):
        raise HTTPException(400, "Apenas .xlsx aceito")
    EXCEL_PATH.write_bytes(await file.read())
    # Limpa cache de fotos
    _cache_fotos.clear()
    _cache_subpastas.clear()
    _cache_timestamp.clear()
    return {"ok": True, "atualizado_em": datetime.now().strftime("%d/%m/%Y %H:%M:%S")}

@app.get("/api/limpar-cache")
def limpar_cache():
    """Força releitura das fotos do Drive"""
    _cache_fotos.clear()
    _cache_subpastas.clear()
    _cache_timestamp.clear()
    return {"ok": True, "msg": "Cache limpo"}

@app.get("/api/status")
def status():
    return {
        "ok": True,
        "excel_existe": EXCEL_PATH.exists(),
        "excel_modificado": datetime.fromtimestamp(EXCEL_PATH.stat().st_mtime).strftime("%d/%m/%Y %H:%M:%S") if EXCEL_PATH.exists() else None,
        "drive_root": DRIVE_ROOT_FOLDER,
        "cache_pontos": len(_cache_fotos),
        "servidor_hora": datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
    }
