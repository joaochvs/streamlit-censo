import sys
import subprocess
import math
import os
import re
import logging
import urllib3
import time
from datetime import datetime
import concurrent.futures
from itertools import combinations

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%H:%M:%S",
)
log = logging.getLogger("orquestrador_rapido")

# =============================================================================
# AUTO-INSTALAÇÃO DE DEPENDÊNCIAS
# =============================================================================
def instalar_ferramentas() -> None:
    log.info("=" * 70)
    log.info("DEEP ORQUESTRADOR V12.4 (RÁPIDO) — verificando dependências")
    log.info("=" * 70)
    pacotes = ["pandas", "openpyxl", "requests", "Pillow", "imagehash", "scikit-learn", "numpy"]
    python_exe = sys.executable.replace("pythonw.exe", "python.exe")
    for pacote in pacotes:
        try:
            mod = "PIL" if pacote == "Pillow" else ("sklearn" if pacote == "scikit-learn" else pacote)
            __import__(mod)
        except ImportError:
            log.info("Instalando %s...", pacote)
            subprocess.check_call([python_exe, "-m", "pip", "install", pacote, "--quiet"])

# =============================================================================
# 1. CONFIGURAÇÕES PRINCIPAIS
# =============================================================================
RAIO_BUSCA_VALIDACAO: float = 250.0   # metros
LIMITE_MESMO_LOCAL:   float = 50.0    # metros

PREENCHER_ENDERECO_OSM: bool = False
SIM_FOTO_INTRUSA: float = 20.0          
SIM_FOTO_OK_MESMO_CODIGO: float = 35.0  

SIM_DESMEMBRAMENTO: float = 55.0   
DIST_DESMEMBRAMENTO: float = 50.0  
SIM_DUPLICATA: float = 70.0        
DIST_DUPLICATA: float = 50.0       

PRIORIDADE_STATUS: dict[str, int] = {
    "✅ OK": 0,
    "⚠️ AVISO": 1,
    "⚠️ AVISO - FOTO OK, SEM GPS": 1,
    "📋 INCOMPLETO": 2,
    "❌ ERRO DE LINK": 2,
    "⚠️ ALERTA": 3,
    "👀 DUPLICATA - REVISÃO - FOTO PARECIDA": 4,
    "⚠️ ALERTA DE PROXIMIDADE": 4,
    "🚨 DESMEMBRAMENTO - 1ª VISITA": 5,
    "🚨 DUPLICATA - MESMO IMÓVEL EM OUTRO CÓDIGO": 5,
    "🚨 DUPLICATA — 1ª VISITA (FOTO IGUAL)": 6,
    "🚨 DUPLICATA - FOTO REAPROVEITADA": 6,
    "🚨 DESMEMBRAMENTO - ERRO GRAVE": 6,
    "🚨 DESMEMBRAMENTO - FOTO INTRUSA": 6,
    "🚨 DESMEMBRAMENTO - INCONSISTÊNCIA CRÍTICA": 7,
    "🚨 DESNEMBRAMENTO - FOTOS DISTINTAS NO MESMO LOCAL": 7,
    "🚨 DESMEMBRAMENTO - OUTRO IMÓVEL NO MESMO CÓDIGO": 8,
}

# =============================================================================
# 2. FUNÇÕES AUXILIARES
# =============================================================================
_VAZIOS = {"", "nan", "<na>", "n/a", "none", "nao", "null"}

def limpar_coordenada(val, tipo: str = "lat") -> float | None:
    if val is None:
        return None
    s = str(val).strip().lower()
    if s in _VAZIOS:
        return None
    try:
        v = float(s.replace(",", "."))
        if math.isnan(v):
            return None
        if abs(v) > 1000:
            v /= 1_000_000.0
        if tipo == "lat" and v > 5.0:
            v = -v
        if tipo == "lat" and not (-90.0 <= v <= 90.0):
            return None
        if tipo == "lon" and not (-180.0 <= v <= 180.0):
            return None
        return v
    except Exception:
        return None

def calcular_distancia(lat1, lon1, lat2, lon2) -> float | None:
    if any(v is None or (isinstance(v, float) and math.isnan(v)) for v in (lat1, lon1, lat2, lon2)):
        return None
    try:
        R = 6_371_000.0
        phi1, phi2 = math.radians(lat1), math.radians(lat2)
        dphi = math.radians(lat2 - lat1)
        dlambda = math.radians(lon2 - lon1)
        a = math.sin(dphi / 2) ** 2 + math.cos(phi1) * math.cos(phi2) * math.sin(dlambda / 2) ** 2
        return float(R * 2 * math.atan2(math.sqrt(a), math.sqrt(1 - a)))
    except Exception:
        return None

def calcular_similaridade(ha, hb) -> float:
    if ha is None or hb is None:
        return -2.0
    if isinstance(ha, str) or isinstance(hb, str):
        return -1.0
    if not isinstance(ha, tuple) or not isinstance(hb, tuple):
        return -1.0
    similares: list[float] = []
    for i in range(3):
        try:
            similares.append(1.0 - ((ha[i] - hb[i]) / 64.0))
        except Exception:
            similares.append(0.0)
    return max(0.0, round(max(similares) * 100, 2))

def aplicar_status(resultado: dict, idx: int, novo_status: str, detalhe: str, cod_suspeito: str = "",
                   similaridade: float | None = None, desmembramento: str = "") -> None:
    if similaridade is not None:
        atual_sim = resultado[idx].get("similaridade", -1.0)
        if similaridade > atual_sim:
            resultado[idx]["similaridade"] = round(float(similaridade), 2)
    if PRIORIDADE_STATUS.get(novo_status, 0) >= PRIORIDADE_STATUS.get(resultado[idx]["status"], 0):
        resultado[idx]["status"] = novo_status
        resultado[idx]["detalhe"] = detalhe
        if cod_suspeito:
            resultado[idx]["codigo_suspeito"] = cod_suspeito
        if desmembramento:
            resultado[idx]["desmembramento"] = desmembramento

def detectar_coluna(df, termos: list[str]) -> str | None:
    for c in df.columns:
        if any(t in str(c).lower().strip() for t in termos):
            return c
    return None

def achar_col_gps(df, termo: str) -> str | None:
    exato = [c for c in df.columns if str(c).strip().lower() == termo]
    if exato:
        return exato[0]
    filtrado = [c for c in df.columns if termo in str(c).lower() and "1.3" not in str(c) and "2.2" not in str(c)]
    if filtrado:
        return filtrado[-1]
    qualquer = [c for c in df.columns if termo in str(c).lower()]
    return qualquer[-1] if qualquer else None

# =============================================================================
# 3. EXTRAÇÃO DE FOTOS E BACKOFFICE (VETORIZADA)
# =============================================================================
def extrair_fotos_vetorizado(df):
    import pandas as pd
    keywords = ["foto", "midia", "arquivo", "link", "imagem", "url", "anexo", "photo", "media", "file"]
    target_cols = [c for c in df.columns if any(k in str(c).lower() for k in keywords)]
    if not target_cols:
        target_cols = df.select_dtypes(include=["object"]).columns.tolist()
    
    combined = df[target_cols].fillna('').astype(str).agg(' '.join, axis=1)
    url_regex = r"(https?://[^\s\'\"<>\[\],]+)"
    all_matches = combined.str.extractall(url_regex)[0]
    if all_matches.empty:
        return pd.Series([[] for _ in range(len(df))], index=df.index)
    
    matches_lower = all_matches.str.lower()
    is_image = matches_lower.str.contains(r"\.(?:jpg|jpeg|png|webp)") | matches_lower.str.contains("deepessoas")
    valid = all_matches[is_image]
    links_per_row = valid.groupby(level=0).apply(lambda x: list(dict.fromkeys(x)))
    return links_per_row.reindex(df.index, fill_value=[])

def extrair_url_backoffice(df):
    import pandas as pd
    col_back = next((c for c in df.columns if "backoffice" in str(c).lower().replace(" ", "")), None)
    if not col_back:
        return pd.Series("", index=df.index)
    def _primeira_url(texto: str) -> str:
        urls = re.findall(r"https?://[^\s\'\"<>\[\],]+", str(texto))
        return urls[0] if urls else ""
    return df[col_back].apply(_primeira_url)

# =============================================================================
# 4. DOWNLOAD E HASH DE IMAGENS
# =============================================================================
def baixar_hash(url: str):
    if not url:
        return url, None
    import requests
    from PIL import Image
    from io import BytesIO
    import imagehash
    headers = {"User-Agent": "Mozilla/5.0"}
    try:
        res = requests.get(url, timeout=15, headers=headers, verify=False)
        if res.status_code == 200:
            try:
                img = Image.open(BytesIO(res.content)).convert("RGB")
                return url, (imagehash.phash(img), imagehash.average_hash(img), imagehash.dhash(img))
            except Exception:
                return url, "ERRO_LINK"
        return url, "ERRO_LINK"
    except Exception:
        return url, "ERRO_LINK"

def pre_carregar_hashes(urls: set, cache: dict) -> None:
    urls_novas = {u for u in urls if u and str(u).strip() and u not in cache}
    if not urls_novas:
        return
    log.info("Baixando %d fotos exclusivas...", len(urls_novas))
    with concurrent.futures.ThreadPoolExecutor(max_workers=100) as ex:
        for url, h in ex.map(baixar_hash, urls_novas):
            cache[url] = h

# =============================================================================
# 5. IDENTIFICAÇÃO DE DUPLICATAS
# =============================================================================
def identificar_duplicatas(df):
    import pandas as pd
    import numpy as np
    df = df.copy()
    df["Duplicata"] = ""
    return df.reset_index(drop=True)

# =============================================================================
# 6. PROCESSAMENTO DE PAR (VALIDAÇÃO VISUAL)
# =============================================================================
def _processar_par(args: tuple) -> list[tuple]:
    (ia, ib, dados_ia, dados_ib, col_agente_flag) = args
    updates: list[tuple] = []
    row_A, row_B = dados_ia, dados_ib
    
    hp_A, hi_A = row_A["hash_prop"], row_A["hash_imov"]
    hp_B, hi_B = row_B["hash_prop"], row_B["hash_imov"]
    
    sims = [s for s in [
        calcular_similaridade(hp_A, hp_B),
        calcular_similaridade(hi_A, hi_B),
        calcular_similaridade(hp_A, hi_B),
        calcular_similaridade(hi_A, hp_B),
    ] if s >= 0.0]
    
    sim_max = max(sims) if sims else 0.0
    sim_min = min(sims) if sims else 0.0
    dist = calcular_distancia(row_A["lat"], row_A["lon"], row_B["lat"], row_B["lon"])
    mesmo_codigo = row_A["id_base"] == row_B["id_base"]
    mesmo_agente = (
        col_agente_flag
        and row_A["agente"] is not None
        and row_B["agente"] is not None
        and str(row_A["agente"]).strip() == str(row_B["agente"]).strip()
    )
    
    inconsistencia_clara = inconsistencia_parcial = False
    if dist is None:
        inconsistencia_clara, inconsistencia_parcial = sim_max >= 90.0, sim_max >= 75.0
    elif dist <= 10.0:
        inconsistencia_clara, inconsistencia_parcial = sim_max >= 55.0, sim_max >= 45.0
    elif dist <= LIMITE_MESMO_LOCAL:
        inconsistencia_clara, inconsistencia_parcial = sim_max >= 75.0, sim_max >= 60.0
    else:
        if mesmo_agente:
            inconsistencia_clara, inconsistencia_parcial = sim_max >= 80.0, sim_max >= 45.0
        else:
            inconsistencia_clara, inconsistencia_parcial = sim_max >= 85.0, sim_max >= 70.0
            
    # Regra Antifraude de Backoffice
    url_back_a = row_A.get("url_backoffice", "")
    url_back_b = row_B.get("url_backoffice", "")
    if url_back_a and url_back_b and url_back_a == url_back_b and sim_max < 40.0:
        if dist is not None and dist <= LIMITE_MESMO_LOCAL:
            msg = f"Link backoffice idêntico, fotos distintas (sim={sim_max:.1f}%, dist={int(dist)}m)."
            updates.append((ia, "🚨 FRAUDE DE BACKOFFICE - FOTOS DISTINTAS NO MESMO LOCAL", msg, row_B["id_completo"], sim_max, ""))
            updates.append((ib, "🚨 FRAUDE DE BACKOFFICE - FOTOS DISTINTAS NO MESMO LOCAL", msg, row_A["id_completo"], sim_max, ""))
        elif dist is not None and dist > LIMITE_MESMO_LOCAL:
            msg = f"Mesmo link backoffice, locais e fotos totalmente diferentes (sim={sim_max:.1f}%, dist={int(dist)}m)."
            updates.append((ia, "🚨 FRAUDE GRAVE - OUTRO IMÓVEL NO MESMO CÓDIGO", msg, row_B["id_completo"], sim_max, ""))
            updates.append((ib, "🚨 FRAUDE GRAVE - OUTRO IMÓVEL NO MESMO CÓDIGO", msg, row_A["id_completo"], sim_max, ""))
            
    for idx_p, row_p, row_o in [(ia, row_A, row_B), (ib, row_B, row_A)]:
        eh_revisita = "-" in str(row_p["id_completo"]) or "-" in str(row_o["id_completo"])
        cod_o = str(row_o["id_completo"])
        if mesmo_codigo:
            imagem_diferente = (len(sims) > 0) and (sim_max < SIM_DESMEMBRAMENTO)
            coord_diferente = (dist is not None) and (dist > DIST_DESMEMBRAMENTO)

            if imagem_diferente or coord_diferente:
                motivos = []
                if imagem_diferente:
                    motivos.append(f"imagem diferente ({sim_max:.1f}% sim)")
                if coord_diferente:
                    motivos.append(f"coordenada diferente ({int(dist)}m)")
                detalhe = "Mesmo Code Deep com " + " e ".join(motivos) + f" vs visita {cod_o}."
                updates.append((idx_p, "🚨 DESMEMBRAMENTO - FOTO INTRUSA",
                                detalhe, cod_o, sim_max, cod_o))
        else:
            coord_proxima = (dist is not None) and (dist <= DIST_DUPLICATA)
            imagem_parecida = (len(sims) > 0) and (sim_max >= SIM_DUPLICATA)

            if coord_proxima and imagem_parecida:
                updates.append((idx_p, "🚨 DUPLICATA - MESMO IMÓVEL EM OUTRO CÓDIGO",
                                f"Mesmo imóvel do código {cod_o} "
                                f"({sim_max:.1f}% sim, {int(dist)}m).", cod_o, sim_max, ""))
            elif inconsistencia_parcial:
                updates.append((idx_p, "👀 REVISÃO - FOTO PARECIDA", f"Possível ângulo diferente: {cod_o} ({sim_max:.1f}% sim).", cod_o, sim_max, ""))
            elif dist is not None and dist <= 10:
                updates.append((idx_p, "⚠️ ALERTA DE PROXIMIDADE", f"GPS muito colado ({int(dist)}m) a outro cadastro.", cod_o, sim_max, ""))
                
    return updates

# =============================================================================
# 6.1 GEOCODIFICAÇÃO REVERSA (Nominatim / OpenStreetMap — gratuito)
# =============================================================================
_GEOCACHE: dict = {}

def geocodificar_reverso(lat, lon, session=None):
    import time
    if lat is None or lon is None:
        return "", ""
    chave = (round(float(lat), 5), round(float(lon), 5))
    if chave in _GEOCACHE:
        return _GEOCACHE[chave]
    try:
        import requests
        if session is None:
            session = requests.Session()
            session.headers.update({"User-Agent": "DeepAuditoria/1.0 (contato@deepessoas.com.br)"})
        url = "https://nominatim.openstreetmap.org/reverse"
        params = {"lat": lat, "lon": lon, "format": "json",
                  "accept-language": "pt-BR", "zoom": 18, "addressdetails": 1}
        resp = session.get(url, params=params, timeout=10)
        time.sleep(1.05)
        if resp.status_code == 200:
            addr = resp.json().get("address", {})
            bairro = (addr.get("suburb") or addr.get("neighbourhood")
                      or addr.get("city_district") or addr.get("quarter") or "").strip()
            logradouro = (addr.get("road") or addr.get("pedestrian")
                          or addr.get("footway") or addr.get("path") or "").strip()
            _GEOCACHE[chave] = (bairro, logradouro)
            return bairro, logradouro
    except Exception as e:
        log.debug("Falha geocodificação (%s,%s): %s", lat, lon, e)
    _GEOCACHE[chave] = ("", "")
    return "", ""

# =============================================================================
# 6.2 REGRAS DE NEGÓCIO DO DOCUMENTO WORD (LIMPEZA / PREENCHIMENTO)
# =============================================================================
def _achar(df, *termos):
    for c in df.columns:
        cl = str(c).lower().strip()
        if all(t.lower() in cl for t in termos):
            return c
    return None

def _norm_txt(serie):
    import pandas as pd
    return serie.fillna("").astype(str).str.strip()

def aplicar_regras_word(df):
    import pandas as pd
    import numpy as np
    df = df.copy()
    metricas: dict = {}

    col_agente_nome = _achar(df, "nome do agente") or _achar(df, "nome", "agente")
    if col_agente_nome:
        nomes = _norm_txt(df[col_agente_nome])
        low = nomes.str.lower()
        mask_escritorio = low.str.contains("gabriel dias", na=False) | low.str.fullmatch(
            r"ana(\s*\d+)?", na=False
        )
        df[col_agente_nome] = np.where(mask_escritorio, "Escritório", df[col_agente_nome])
        log.info("R1: %d nomes de agente padronizados para 'Escritório'.", int(mask_escritorio.sum()))

    col_data_reg = _achar(df, "data do registro") or _achar(df, "data", "registro")
    col_data_vis = _achar(df, "data da visita") or _achar(df, "data", "visita")
    if col_data_reg and col_data_vis:
        df[col_data_vis] = df[col_data_reg]
        log.info("R2: 'Data da Visita' copiada de 'Data do Registro'.")

    cols_foto = [c for c in df.columns if "foto" in str(c).lower()]
    if cols_foto:
        tem_foto = df[cols_foto].apply(
            lambda col: _norm_txt(col).str.lower().isin(_VAZIOS) == False
        ).any(axis=1)
        metricas["visitas_sem_foto"] = int((~tem_foto).sum())
        if "ID_Base" in df.columns and cols_foto:
            primeira_foto = (
                df.assign(_f=df[cols_foto].apply(
                    lambda r: next((str(v) for v in r if str(v).strip().lower() not in _VAZIOS), ""), axis=1))
                  .groupby("ID_Base")["_f"]
                  .transform(lambda s: next((v for v in s if v), ""))
            )
            for c in cols_foto:
                vazio = _norm_txt(df[c]).str.lower().isin(_VAZIOS)
                pode = vazio & (primeira_foto != "")
                df.loc[pode, c] = primeira_foto[pode]
        log.info("R3: %d visitas sem nenhuma foto.", metricas["visitas_sem_foto"])

    pares_coord = [
        ("latitude", "longitude"),
        ("1.3.1 coordenada porteira - captura 1 latitude", "1.3.1 coordenada porteira - captura 1 longitude"),
        ("2.2.2 coordenada do imóvel - captura 1 latitude", "2.2.2 coordenada do imóvel - captura 1 longitude"),
    ]
    achou_alguma = pd.Series(False, index=df.index)
    for lat_t, lon_t in pares_coord:
        cl = next((c for c in df.columns if str(c).lower().strip() == lat_t), None)
        co = next((c for c in df.columns if str(c).lower().strip() == lon_t), None)
        if cl and co:
            latv = df[cl].apply(lambda v: limpar_coordenada(v, "lat"))
            lonv = df[co].apply(lambda v: limpar_coordenada(v, "lon"))
            achou_alguma = achou_alguma | (latv.notna() & lonv.notna())
    metricas["visitas_sem_coordenada"] = int((~achou_alguma).sum())
    log.info("R4: %d visitas sem nenhuma coordenada (nas 3 fontes).", metricas["visitas_sem_coordenada"])

    for perg_t, val_t in [("1.7 possui número", "1.7.1 número"),
                          ("1.8 possui complemento", "1.8.1 complemento")]:
        cp = _achar(df, perg_t)
        cv = _achar(df, val_t)
        if cp and cv:
            df[cp] = _norm_txt(df[cp])
            val_preenchido = ~_norm_txt(df[cv]).str.lower().isin(_VAZIOS)
            perg_lower = df[cp].str.lower()
            perg_vazia = perg_lower.isin(_VAZIOS)
            perg_nao = perg_lower.isin({"nao", "não"})
            corrigir = val_preenchido & (perg_vazia | perg_nao)
            df.loc[corrigir, cp] = "Sim"
            if corrigir.any():
                log.info("R5: %d '%s' ajustadas para 'Sim' (valor preenchido ao lado).",
                         int(corrigir.sum()), cp)

    col_est_uso    = _achar(df, "2.1 estrutura está em uso") or _achar(df, "estrutura está em uso")
    col_est_status = _achar(df, "2.1.1 status da estrutura") or _achar(df, "status da estrutura")
    col_encontrado = _achar(df, "alguém foi encontrado") or _achar(df, "alguem foi encontrado")
    col_motivo     = _achar(df, "qual motivo de alguém não ter sido contactado") or _achar(df, "motivo", "contactado")
    if col_est_uso and col_est_status:
        df[col_est_uso] = _norm_txt(df[col_est_uso])
        df[col_est_status] = _norm_txt(df[col_est_status])
        uso = df[col_est_uso].str.lower()
        mask_sim = uso == "sim"
        df.loc[mask_sim, col_est_status] = "Em uso"
        motivos_uso = {
            "edificação longe da porteira", "edificacao longe da porteira",
            "não foi encontrada porteira", "nao foi encontrada porteira",
            "não foi localizado", "nao foi localizado",
            "propriedade totalmente produtiva",
        }
        if col_motivo:
            motivo_norm = _norm_txt(df[col_motivo]).str.lower()
            branco_am = _norm_txt(df[col_est_uso]).str.lower().isin(_VAZIOS)
            mask_motivo = branco_am & motivo_norm.isin(motivos_uso)
            df.loc[mask_motivo, col_est_uso] = "Sim"
            df.loc[mask_motivo, col_est_status] = "Em uso"
        if col_encontrado:
            enc = _norm_txt(df[col_encontrado]).str.lower()
            mask_ab = enc == "abandonado"
            df.loc[mask_ab, col_est_uso] = "Não"
            df.loc[mask_ab, col_est_status] = "Abandonado"
        log.info("R7: coerência de estrutura (AM/AN) aplicada.")

    def _fmt_cpf(s):
        d = re.sub(r"\D", "", str(s))
        return f"{d[:3]}.{d[3:6]}.{d[6:9]}-{d[9:11]}" if len(d) == 11 else str(s)

    def _fmt_cnpj(s):
        d = re.sub(r"\D", "", str(s))
        return f"{d[:2]}.{d[2:5]}.{d[5:8]}/{d[8:12]}-{d[12:14]}" if len(d) == 14 else str(s)

    col_cpf  = _achar(df, "cpf do entrevistado") or _achar(df, "3.3.1.1 cpf")
    col_cnpj = _achar(df, "cnpj do entrevistado") or _achar(df, "3.3.1.1 cnpj")
    if col_cpf:
        df[col_cpf] = df[col_cpf].apply(lambda v: _fmt_cpf(v) if str(v).strip().lower() not in _VAZIOS else v)
    if col_cnpj:
        df[col_cnpj] = df[col_cnpj].apply(lambda v: _fmt_cnpj(v) if str(v).strip().lower() not in _VAZIOS else v)
    if col_cpf or col_cnpj:
        log.info("R8: CPF/CNPJ padronizados.")

    col_cel = _achar(df, "telefone do entrevistado") or _achar(df, "celular") or _achar(df, "telefone")
    if col_cel:
        def _fmt_cel(s):
            d = re.sub(r"\D", "", str(s))
            if len(d) == 11:
                return f"{d[:2]} - {d[2:7]} - {d[7:]}"
            if len(d) == 10:
                return f"{d[:2]} - {d[2:6]} - {d[6:]}"
            return str(s)
        df[col_cel] = df[col_cel].apply(lambda v: _fmt_cel(v) if str(v).strip().lower() not in _VAZIOS else v)
        log.info("R9: telefones formatados.")

    if col_motivo and col_encontrado:
        df[col_motivo] = _norm_txt(df[col_motivo])
        enc_norm = _norm_txt(df[col_encontrado]).str.lower()
        manter = enc_norm == "não"
        esvaziados = int((~manter & (_norm_txt(df[col_motivo]) != "")).sum())
        df.loc[~manter, col_motivo] = ""
        log.info("R10: motivo (coluna X) esvaziado em %d linhas onde 'Alguém foi encontrado?' != 'Não'.", esvaziados)

    return df, metricas

def preencher_bairro_logradouro(df, col_lat, col_lon):
    import pandas as pd
    col_bairro = _achar(df, "1.5 bairro") or _achar(df, "bairro")
    col_logra  = _achar(df, "1.6 logradouro") or _achar(df, "logradouro")
    if not (col_bairro or col_logra) or not (col_lat and col_lon):
        return df

    df = df.copy()
    if col_bairro:
        df[col_bairro] = _norm_txt(df[col_bairro])
    if col_logra:
        df[col_logra] = _norm_txt(df[col_logra])
    import requests
    sess = requests.Session()
    sess.headers.update({"User-Agent": "DeepAuditoria/1.0 (contato@deepessoas.com.br)"})

    preenchidos = 0
    for i in df.index:
        b_vazio = (not col_bairro) or (str(df.at[i, col_bairro]).strip().lower() in _VAZIOS)
        l_vazio = (not col_logra)  or (str(df.at[i, col_logra]).strip().lower() in _VAZIOS)
        if not (b_vazio or l_vazio):
            continue
        lat = limpar_coordenada(df.at[i, col_lat], "lat") if col_lat else None
        lon = limpar_coordenada(df.at[i, col_lon], "lon") if col_lon else None
        if lat is None or lon is None:
            continue
        bairro, logra = geocodificar_reverso(lat, lon, sess)
        if col_bairro and b_vazio and bairro:
            df.at[i, col_bairro] = bairro
            preenchidos += 1
        if col_logra and l_vazio and logra:
            df.at[i, col_logra] = logra
            preenchidos += 1
    log.info("R6: %d campos de Bairro/Logradouro preenchidos via geocodificação.", preenchidos)
    return df

# =============================================================================
# 7. EXECUÇÃO PRINCIPAL
# =============================================================================


def processar_desmembramentos(arquivo_entrada, progresso=None):
    """Executa a lógica original usando CSV/XLSX e devolve a base final."""
    import pandas as pd
    import numpy as np
    from io import BytesIO
    from sklearn.neighbors import BallTree

    def atualizar(percentual, etapa, detalhe=""):
        if progresso:
            progresso(percentual, etapa, detalhe)

    atualizar(0.02, "Lendo a base", "Identificando CSV ou Excel")
    try:
        posicao = arquivo_entrada.tell() if hasattr(arquivo_entrada, "tell") else None
        assinatura = arquivo_entrada.read(4) if hasattr(arquivo_entrada, "read") else b""
        if posicao is not None:
            arquivo_entrada.seek(posicao)
        nome = str(getattr(arquivo_entrada, "name", arquivo_entrada)).lower()
        eh_excel = (
            assinatura.startswith(b"PK")
            or assinatura == b"\xd0\xcf\x11\xe0"
            or nome.endswith((".xlsx", ".xls"))
        )
        if eh_excel:
            df = pd.read_excel(arquivo_entrada, dtype=str)
        else:
            conteudo = arquivo_entrada.read() if hasattr(arquivo_entrada, "read") else None
            ultimo_erro = None
            for codificacao in ("utf-8-sig", "utf-8", "latin-1"):
                try:
                    fonte = BytesIO(conteudo) if conteudo is not None else arquivo_entrada
                    df = pd.read_csv(fonte, dtype=str, sep=None, engine="python", encoding=codificacao)
                    break
                except UnicodeDecodeError as erro:
                    ultimo_erro = erro
            else:
                raise ultimo_erro or ValueError("Codificação do CSV não reconhecida.")
    except Exception as erro:
        raise ValueError(f"Não foi possível ler o arquivo: {erro}") from erro

    df = df.astype(str)
    df = identificar_duplicatas(df)
    atualizar(0.12, "Preparando a base", f"{len(df):,} registros carregados".replace(",", "."))
    
    col_codigo = detectar_coluna(df, ["codigo_unico", "code deep", "código"])
    col_agente = detectar_coluna(df, ["agente", "pesquisador", "entrevistador"])
    col_nome_agente = next((c for c in df.columns if "nome do agente" in str(c).lower() or ("nome" in str(c).lower() and "agente" in str(c).lower())), None)
    
    if not col_agente and col_nome_agente:
        col_agente = col_nome_agente
    if not col_codigo:
        log.error("Coluna de Código Único não encontrada.")
        raise SystemExit(1)
        
    col_lat = achar_col_gps(df, "latitude")
    col_lon = achar_col_gps(df, "longitude")
    
    atualizar(0.22, "Localizando imagens", "Procurando links nas colunas da base")
    log.info("Extraindo links de imagem...")
    df["Links_Detectados"] = extrair_fotos_vetorizado(df)
    df["URL_Prop"] = df["Links_Detectados"].apply(lambda x: x[0] if len(x) > 0 else "")
    df["URL_Imov"] = df["Links_Detectados"].apply(lambda x: x[1] if len(x) > 1 else "")
    df["URL_Backoffice"] = extrair_url_backoffice(df)
    
    total_links = df["Links_Detectados"].apply(len).sum()
    log.info("%d links válidos detectados.", total_links)
    
    df["ID_Completo"] = df[col_codigo].astype(str).str.strip().str.upper()
    df["ID_Base"] = df["ID_Completo"].str.split("-").str[0]
    df["Lat_Num"] = df[col_lat].apply(lambda v: limpar_coordenada(v, "lat")) if col_lat else np.nan
    df["Lon_Num"] = df[col_lon].apply(lambda v: limpar_coordenada(v, "lon")) if col_lon else np.nan
    
    log.info("Construindo pares de análise...")
    pares: set[tuple[int, int]] = set()
    for id_base, grupo in df.groupby("ID_Base"):
        if str(id_base).upper().startswith("SEM_ID") or str(id_base) in ("", "NAN"):
            continue
        idxs = grupo.index.tolist()
        for ia, ib in combinations(idxs, 2):
            pares.add((ia, ib))
            
    df_gps = df.dropna(subset=["Lat_Num", "Lon_Num"]).copy()
    if not df_gps.empty:
        log.info("BallTree com raio de %.0fm...", RAIO_BUSCA_VALIDACAO)
        coords = np.radians(df_gps[["Lat_Num", "Lon_Num"]].values)
        tree = BallTree(coords, metric="haversine")
        raio_rad = RAIO_BUSCA_VALIDACAO / 6_371_000.0
        indices_vizinhos = tree.query_radius(coords, r=raio_rad)
        for i, vizinhos in enumerate(indices_vizinhos):
            idx_A = df_gps.index[i]
            for j in vizinhos:
                if i < j:
                    idx_B = df_gps.index[j]
                    if df.at[idx_A, "ID_Base"] != df.at[idx_B, "ID_Base"]:
                        pares.add((idx_A, idx_B))
                        
    log.info("Total de pares: %d", len(pares))
    
    urls_dl = set(df["URL_Prop"].dropna().tolist() + df["URL_Imov"].dropna().tolist())
    urls_dl.discard("")
    cache_hash: dict = {}
    atualizar(0.40, "Analisando imagens", f"{len(urls_dl):,} links exclusivos encontrados".replace(",", "."))
    pre_carregar_hashes(urls_dl, cache_hash)
    atualizar(0.70, "Imagens analisadas", "Preparando comparações entre registros")
    
    dados_rows: dict[int, dict] = {}
    for i in df.index:
        row = df.loc[i]
        dados_rows[i] = {
            "id_completo": str(row["ID_Completo"]),
            "id_base": str(row["ID_Base"]),
            "agente": str(row[col_agente]) if col_agente else None,
            "lat": row["Lat_Num"],
            "lon": row["Lon_Num"],
            "hash_prop": cache_hash.get(row["URL_Prop"]),
            "hash_imov": cache_hash.get(row["URL_Imov"]),
            "url_backoffice": str(row["URL_Backoffice"]).strip() if pd.notna(row["URL_Backoffice"]) else "",
        }
        
    args_pares = (
        (ia, ib, dados_rows[ia], dados_rows[ib], bool(col_agente))
        for ia, ib in pares
    )
    
    resultado: dict[int, dict] = {
        i: {"status": "✅ OK", "detalhe": "", "codigo_suspeito": "",
            "similaridade": -1.0, "desmembramento": ""}
        for i in df.index
    }
    
    atualizar(0.74, "Comparando registros", f"{len(pares):,} pares para analisar".replace(",", "."))
    log.info("Analisando pares visualmente (CPU/Processos)...")
    if pares:
        with concurrent.futures.ProcessPoolExecutor() as executor:
            for updates in executor.map(_processar_par, args_pares, chunksize=2000):
                for idx_p, novo_status, detalhe, cod, sim, desmemb in updates:
                    aplicar_status(resultado, idx_p, novo_status, detalhe, cod,
                                   similaridade=sim, desmembramento=desmemb)
                
    for i, row in df.iterrows():
        if resultado[i]["status"] != "✅ OK":
            continue
        hp = cache_hash.get(row["URL_Prop"])
        hi = cache_hash.get(row["URL_Imov"])
        sem_foto_valida = (hp is None or isinstance(hp, str)) and (hi is None or isinstance(hi, str))
        sem_gps = pd.isna(row["Lat_Num"])
        if sem_foto_valida:
            if (row["URL_Prop"] or row["URL_Imov"]) and (isinstance(hp, str) or isinstance(hi, str)):
                aplicar_status(resultado, i, "❌ ERRO DE LINK", "Bloqueio de servidor (SSL/Permissão).")
            elif sem_gps:
                aplicar_status(resultado, i, "📋 INCOMPLETO", "Sem fotos e sem GPS.")
            else:
                aplicar_status(resultado, i, "⚠️ AVISO", "Scanner não encontrou links de imagem.")
        elif not sem_foto_valida and sem_gps:
            aplicar_status(resultado, i, "⚠️ AVISO - FOTO OK, SEM GPS", "Sem coordenadas GPS.")
            
    df["Status_Validacao"] = [resultado[i]["status"] for i in df.index]
    df["Codigo_Suspeito"] = [resultado[i]["codigo_suspeito"] for i in df.index]
    df["Detalhe_Inconsistencia"] = [resultado[i]["detalhe"] for i in df.index]

    df["Percentual_Similaridade"] = [
        (f"{resultado[i]['similaridade']:.2f}%" if resultado[i]["similaridade"] >= 0 else "")
        for i in df.index
    ]

    def _desmembramento_linha(i):
        desm = resultado[i].get("desmembramento", "")
        if not desm:
            return ""
        return f"Desmembramento: imóvel/coordenada diferente vs visita {desm}"

    df["Apontamento_Desmembramento"] = [_desmembramento_linha(i) for i in df.index]

    atualizar(0.92, "Aplicando regras finais", "Consolidando status e apontamentos")
    df, metricas_word = aplicar_regras_word(df)

    if PREENCHER_ENDERECO_OSM:
        df = preencher_bairro_logradouro(df, col_lat, col_lon)
    else:
        log.info("R6: preenchimento de Bairro/Logradouro via OSM DESLIGADO "
                 "(PREENCHER_ENDERECO_OSM=False).")

    _drop = {"Lat_Num", "Lon_Num", "URL_Prop", "URL_Imov", "ID_Base", "Links_Detectados"}
    df_saida = df[[c for c in df.columns if c not in _drop]]

    def classificar_suave(status):
        status_lower = str(status).lower().strip()
        if 'reprovado' in status_lower or 'erro' in status_lower or '🚨' in status_lower:
            return 'Revisão Recomendada'
        elif 'suspeito' in status_lower or 'fraude' in status_lower:
            return 'Verificação de Detalhes'
        elif 'pendente' in status_lower:
            return 'Aguardando Validação'
        elif 'ok' in status_lower:
            return 'Em Conformidade'
        else:
            return 'Em Conformidade'

    df_saida = df_saida.copy()
    df_saida['apontamento_ia'] = df_saida['Status_Validacao'].apply(classificar_suave)
    # Na origem, campos vazios foram convertidos para texto para preservar os
    # tipos durante as regras. Antes da exportação, restaura esses marcadores
    # para que o CSV tenha células realmente vazias.
    df_saida = df_saida.replace(r'^\s*(?:nan|none|<na>|null)\s*$', '', regex=True)
    atualizar(1.0, 'Processamento concluído', f'{len(df_saida):,} registros preparados'.replace(',', '.'))
    return df_saida
