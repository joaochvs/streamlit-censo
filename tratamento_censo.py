import sys
import subprocess
import csv
import math
import os
import re
import urllib3
from datetime import datetime
import concurrent.futures
from itertools import combinations

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(errors="replace")

# =============================================================================
# AUTO-INSTALAÇÃO DE DEPENDÊNCIAS E AVISOS
# =============================================================================
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

def instalar_ferramentas():
    print("\n" + "="*80)
    print("🧠 DEEP ORQUESTRADOR V12.1 — MALHA DE MUNICÍPIOS E VALIDAÇÃO VISUAL (OTIMIZADO)")
    print("="*80)
    pacotes = ['pandas', 'openpyxl', 'requests', 'Pillow', 'imagehash', 'scikit-learn', 'numpy']
    python_exe = sys.executable.replace("pythonw.exe", "python.exe")
    for pacote in pacotes:
        try:
            if pacote == 'Pillow': __import__('PIL')
            elif pacote == 'scikit-learn': __import__('sklearn')
            else: __import__(pacote)
        except ImportError:
            print(f"   📦 Instalando biblioteca ausente: {pacote}...")
            subprocess.check_call([python_exe, "-m", "pip", "install", pacote, "--quiet"])

import pandas as pd
import numpy as np
import requests
from PIL import Image
from io import BytesIO
import imagehash
from sklearn.neighbors import BallTree


def ler_csv_flexivel(conteudo):
    """Lê CSVs comuns e tolera aspas malformadas exportadas por outros sistemas."""
    ultimo_erro = None
    for codificacao in ("utf-8-sig", "utf-8", "latin-1"):
        try:
            return pd.read_csv(
                BytesIO(conteudo),
                dtype=str,
                sep=None,
                engine="python",
                encoding=codificacao,
            )
        except UnicodeDecodeError as erro:
            ultimo_erro = erro
            continue
        except (pd.errors.ParserError, csv.Error) as erro:
            ultimo_erro = erro

        # Alguns sistemas geram campos com aspas abertas ou fechadas no meio
        # do texto. Nesse caso, detecta o separador sem interpretar aspas.
        try:
            texto = conteudo.decode(codificacao)
            linhas = [linha for linha in texto.splitlines() if linha.strip()][:20]
            separadores = (";", ",", "\t", "|")
            separador = max(
                separadores,
                key=lambda candidato: sum(linha.count(candidato) for linha in linhas),
            )
            if not linhas or not any(separador in linha for linha in linhas):
                continue
            return pd.read_csv(
                BytesIO(conteudo),
                dtype=str,
                sep=separador,
                engine="python",
                encoding=codificacao,
                quoting=csv.QUOTE_NONE,
            )
        except (UnicodeDecodeError, pd.errors.ParserError, csv.Error) as erro:
            ultimo_erro = erro

    raise ultimo_erro or ValueError("Não foi possível identificar o formato do CSV.")

# =============================================================================
# ✏️ NOVO: REPARO DE ENCODING (corrige "NÃ£o" -> "Não")
# =============================================================================
def reparar_mojibake(df):
    """
    Corrige "mojibake" (dupla codificação UTF-8), que faz "Não" virar "NÃ£o".
    Isso ocorre quando texto UTF-8 foi lido como Latin-1/CP1252 em um passo
    anterior (tipicamente um CSV salvo/aberto no Excel com codificação errada)
    e o .xlsx de entrada já chegou corrompido. Aqui revertemos célula a célula:
    re-encodamos em Latin-1 e decodificamos de volta em UTF-8. Só aplicamos a
    correção quando ela melhora o texto (presença dos marcadores Ã/Â), para
    nunca estragar textos que já estão corretos.
    """
    marcadores = ("Ã", "Â", "â\x80", "Ã\x83")

    def _conserta(valor):
        if not isinstance(valor, str) or not valor:
            return valor
        if not any(m in valor for m in ("Ã", "Â")):
            return valor
        try:
            corrigido = valor.encode("latin-1").decode("utf-8")
        except (UnicodeEncodeError, UnicodeDecodeError):
            return valor
        # Só aceita se de fato removeu os marcadores de mojibake.
        if any(m in corrigido for m in ("Ã", "Â")) and not any(m in valor.replace(corrigido, "") for m in marcadores):
            # se ainda restou Ã/Â, tenta um segundo passe (casos de tripla codificação)
            try:
                corrigido2 = corrigido.encode("latin-1").decode("utf-8")
                if not any(m in corrigido2 for m in ("Ã", "Â")):
                    return corrigido2
            except (UnicodeEncodeError, UnicodeDecodeError):
                pass
        return corrigido

    # Corrige os NOMES das colunas e o CONTEÚDO das colunas de texto.
    df = df.rename(columns={c: _conserta(c) if isinstance(c, str) else c for c in df.columns})
    for col in df.select_dtypes(include=["object"]).columns:
        df[col] = df[col].map(_conserta)
    return df


# =============================================================================
# 1. CONFIGURAÇÕES PRINCIPAIS E LIMITES DO SISTEMA
# =============================================================================

# 🔧 CONTROLE PRINCIPAL DE VELOCIDADE
VALIDAR_BAIRRO_OSM = False   # ← Mude para True se quiser validar bairros no OpenStreetMap (MUITO lento)

ARQUIVO_ENTRADA = r"C:\Users\Edulm\OneDrive - DEEPESSOAS\Área de Trabalho\Planilha com os dois códigos\Planilha de teste dos dois códigos.xlsx"
agora_str = datetime.now().strftime("%Hh%Mm%Ss")
ARQUIVO_SAIDA = os.path.join(os.path.dirname(ARQUIVO_ENTRADA), f'Relatorio_Master_V12.1_{agora_str}.xlsx')

# 🎯 OTIMIZAÇÃO: Raio reduzido de 2500m para 250m. Evita criar milhões de cruzamentos desnecessários
RAIO_BUSCA_VALIDACAO = 250.0  
LIMITE_MESMO_LOCAL  = 50.0    
LIMITE_TRIANGULACAO = 500.0   

PRIORIDADE_STATUS = {
    '✅ OK': 0, '⚠️ AVISO': 1, '⚠️ AVISO - FOTO OK, SEM GPS': 1, '📋 INCOMPLETO': 2, '❌ ERRO DE LINK': 2,
    '⚠️ ALERTA': 3, '👀 REVISÃO - FOTO PARECIDA': 4, '⚠️ ALERTA DE PROXIMIDADE': 4, 
    '🚨 DESMEMBRAMENTO - 1ª VISITA': 5,
    # ✏️ NOVO: duplicata pela definição correta (mesmo imóvel, códigos diferentes).
    '🚨 DUPLICATA - MESMO IMÓVEL EM OUTRO CÓDIGO': 5,
    '🚨 ALERTA — 1ª VISITA (FOTO IGUAL)': 6, '🚨 DESMEMBRAMENTO COM FOTO REAPROVEITADA': 6,
    # ✏️ NOVO (Regra 1/2): apontamento dedicado de foto intrusa.
    '🚨 DESMEMBRAMENTO - FOTO INTRUSA': 6,
    '🚨 ERRO GRAVE': 6, '🚨 INCONSISTÊNCIA CRÍTICA': 7,
}

# ✏️ NOVO (Regra 1 e 4): limiares de tolerância de imagem dentro do MESMO código.
# Fotos iguais/parecidas (sim alta) NÃO geram apontamento; só a foto INTRUSA
# (extremamente diferente, sim mínima abaixo do limiar) é apontada.
SIM_FOTO_INTRUSA = 20.0          # < 20% de similaridade = foto intrusa
SIM_FOTO_OK_MESMO_CODIGO = 35.0  # >= 35% = imóvel compatível, silêncio

# ✏️ NOVO: limiares das DEFINIÇÕES CORRETAS de Desmembramento e Duplicata.
#   DESMEMBRAMENTO (mesmo Code Deep): imóveis diferentes no mesmo código.
#     Basta UMA divergência: imagem (sim < SIM_DESMEMBRAMENTO) OU
#     coordenada (dist > DIST_DESMEMBRAMENTO m).
#   DUPLICATA (Code Deeps DIFERENTES): mesmo imóvel em dois blocos.
#     Exige AS DUAS: coordenada próxima (dist <= DIST_DUPLICATA) E
#     imagem parecida (sim >= SIM_DUPLICATA).
SIM_DESMEMBRAMENTO = 55.0
DIST_DESMEMBRAMENTO = 50.0
SIM_DUPLICATA = 70.0
DIST_DUPLICATA = 50.0

# ⚡ DICIONÁRIO OFICIAL DE CIDADES
CENTROIDES_MUNICIPIOS = {
    'CARAÇA': [(-23.4388, -47.0636)], 'CBADA':  [(-23.6272, -49.5622)], 
    'CAPAO':  [(-24.0063, -48.3494)], 'CCGUA':  [(-23.6197, -45.4127)], 
    'CMACE':  [(-23.6288, -49.3141)], 'CEMBU':  [(-23.6490, -46.8520)], 
    'CGUAPI': [(-24.1852, -48.5327)], 'CITAI':  [(-23.4172, -49.0905)], 
    'CITNHA': [(-24.1838, -46.7888)], 'CRIBE':  [(-24.6558, -49.0063)], 
    'CEDRA':  [(-24.5211, -48.1072)], 'CFART':  [(-23.6455, -49.5316)], 
    'CIGUAP': [(-24.7083, -47.5552)], 'CITBR':  [(-23.8619, -49.1363)], 
    'CITPGA': [(-23.7086, -49.4900)], 'CJQUÁ':  [(-24.3208, -47.6352)], 
    'CMAIPÃ': [(-23.3186, -46.5866)], 'CNVCP':  [(-24.1202, -48.9036)], 
    'CPIAÇU': [(-24.7141, -47.8816)], 'CRIBR':  [(-24.2216, -48.7663)], 
    'CSAOP':  [(-23.5505, -46.6333)], 'CBUR':   [(-23.7958, -48.5927)], 
    'CAJAMA': [(-23.3555, -46.8777)], 'CMA':    [(-23.5936, -48.4755)], 
    'CILB':   [(-23.7780, -45.3558)], 'CSJDC':  [(-23.2237, -45.9009)], 
    'CSTB':   [(-24.3872, -47.9255)], 'CIPRGA': [(-24.0058, -46.4028)],
    'CITAPE': [(-23.7135, -46.8491), (-23.9822, -48.8755)], 
    'CTAQUA': [(-23.9238, -48.6947), (-23.5319, -49.2458)], 
    'CMCATU': [(-24.5855, -48.5927), (-24.2811, -47.4580), (-23.4502, -49.4083)], 
    'CGUA':   [(-23.9930, -46.2563), (-23.6197, -45.4127)], 
}

# =============================================================================
# 2. FUNÇÕES AUXILIARES DE LIMPEZA E LÓGICA
# =============================================================================

def limpar_coordenada(val, tipo='lat'):
    _VAZIOS = {"", "nan", "<na>", "n/a", "none", "nao", "null"}
    if pd.isna(val) or str(val).strip().lower() in _VAZIOS: return None
    try:
        v = float(str(val).replace(',', '.').strip())
        if math.isnan(v): return None
        if abs(v) > 1000: v = v / 1_000_000.0
        if tipo == 'lat' and v > 5.0: v = -v
        if tipo == 'lat' and not (-90.0 <= v <= 90.0): return None
        if tipo == 'lon' and not (-180.0 <= v <= 180.0): return None
        return v
    except Exception: return None

def calcular_distancia(lat1, lon1, lat2, lon2):
    if any(x is None or (isinstance(x, float) and math.isnan(x)) for x in (lat1, lon1, lat2, lon2)): return None
    try:
        R = 6_371_000 
        phi1, phi2 = math.radians(lat1), math.radians(lat2)
        dphi, dlambda = math.radians(lat2 - lat1), math.radians(lon2 - lon1)
        a = math.sin(dphi/2)**2 + math.cos(phi1) * math.cos(phi2) * math.sin(dlambda/2)**2
        return float(R * 2 * math.atan2(math.sqrt(a), math.sqrt(1 - a)))
    except Exception: return None

def calcular_similaridade(ha, hb):
    if isinstance(ha, str) or isinstance(hb, str): return -1.0
    if ha is None or hb is None: return -2.0
    if not isinstance(ha, tuple) or not isinstance(hb, tuple): return -1.0
    similares = []
    for i in range(3):
        try: similares.append(1.0 - ((ha[i] - hb[i]) / 64.0))
        except Exception: similares.append(0.0)
    return max(0.0, round(max(similares) * 100, 2))

def aplicar_status(resultado, idx, novo_status, detalhe, cod_suspeito=""):
    if PRIORIDADE_STATUS.get(novo_status, 0) >= PRIORIDADE_STATUS.get(resultado[idx]['status'], 0):
        resultado[idx]['status'] = novo_status
        resultado[idx]['detalhe'] = detalhe
        if cod_suspeito: resultado[idx]['codigo_suspeito'] = cod_suspeito

def detectar_coluna(df, termos):
    for c in df.columns:
        if any(t in str(c).lower().strip() for t in termos): return c
    return None

def achar_col_gps(df, termo):
    exato = [c for c in df.columns if str(c).strip().lower() == termo]
    if exato: return exato[0]
    filtrado = [c for c in df.columns if termo in str(c).lower() and '1.3' not in str(c) and '2.2' not in str(c)]
    if filtrado: return filtrado[-1]
    qualquer = [c for c in df.columns if termo in str(c).lower()]
    if qualquer: return qualquer[-1]
    return None

def extrair_fotos_da_linha(row):
    urls = []
    for val in row.values:
        val_str = str(val).strip()
        if 'http' in val_str:
            for link in re.findall(r'(https?://[^\s\'"<>\[\],]+)', val_str):
                link_l = link.lower()
                if (any(ext in link_l for ext in ['.jpg', '.jpeg', '.png', '.webp']) or 'deepessoas' in link_l) and 'backoffice' not in link_l:
                    if link not in urls: urls.append(link)
    return urls

# =============================================================================
# 3. MÓDULO DE DOWNLOAD E ETL
# =============================================================================

def baixar_hash(url):
    if not url: return url, None
    headers = {"User-Agent": "Mozilla/5.0"}
    try:
        res = requests.get(url, timeout=15, headers=headers, verify=False)
        if res.status_code == 200:
            try:
                img = Image.open(BytesIO(res.content)).convert("RGB")
                return url, (imagehash.phash(img), imagehash.average_hash(img), imagehash.dhash(img))
            except Exception: return url, "ERRO_LINK"
        return url, "ERRO_LINK"
    except Exception: return url, "ERRO_LINK"

def pre_carregar_hashes(urls, cache, progresso=None):
    urls_novas = {u for u in urls if u and str(u).strip() != "" and u not in cache}
    if not urls_novas: return
    total = len(urls_novas)
    print(f"   📡 Baixando e convertendo {total} fotos exclusivas detectadas...")
    # Mantém o consumo de memória e conexões sob controle no Streamlit Cloud.
    # A quantidade de imagens e as regras de análise permanecem inalteradas.
    with concurrent.futures.ThreadPoolExecutor(max_workers=6) as ex:
        for concluida, (url, h) in enumerate(ex.map(baixar_hash, urls_novas), start=1):
            cache[url] = h
            if progresso and (concluida == total or concluida % max(1, total // 100) == 0):
                fracao = concluida / total
                progresso(
                    0.35 + (0.35 * fracao),
                    "Baixando e analisando as fotos",
                    f"{concluida:,} de {total:,} imagens processadas".replace(",", "."),
                )

def identificar_duplicatas(df):
    df = df.copy()
    # ✏️ CORREÇÃO CONCEITUAL: marcação de "DUPLICATA" por Code Deep REPETIDO
    # REMOVIDA. Code Deep repetido (visitas -1, -2...) é o bloco NORMAL de
    # visitas, não duplicata. Duplicata real (mesmo imóvel em códigos DIFERENTES)
    # e Desmembramento (mesmo código com imagem/coord diferente) são detectados
    # no motor visual (_processar_par). Aqui só criamos a coluna vazia.
    df['Duplicata'] = ''
    return df.reset_index(drop=True)

# =============================================================================
# 5. MOTOR DE AUDITORIA QUALITATIVA COM MALHA DE MUNICÍPIOS
# =============================================================================

def auditar_base_visitas(df):
    df_audit = df.copy()
    if 'Erros_Logicos' not in df_audit.columns: df_audit['Erros_Logicos'] = ""

    if 'Renda_Familiar_Num' not in df_audit.columns:
        try: df_audit['Renda_Familiar_Num'] = pd.to_numeric(df_audit.get('4.8 Qual a renda familiar?_field'), errors='coerce')
        except Exception: df_audit['Renda_Familiar_Num'] = np.nan

    col_enc = next((c for c in df_audit.columns if '1.2 Alguém foi encontrado' in c), None)
    mask_ausente = df_audit[col_enc].astype(str).str.strip().str.upper().isin(['NÃO', 'NAO', 'ABANDONADO']) if col_enc else pd.Series(False, index=df_audit.index)

    mask_A = pd.Series(False, index=df_audit.index)
    try:
        c1, c2, c3 = df_audit.get('4.2 Número de moradores na residência_field'), df_audit.get('4.8 Qual a renda familiar?_field'), df_audit.get('4.9 Qual o nivel de escolaridade dos adultos (maiores de 18 anos) residentes nesta propriedade?_field')
        cond = pd.Series(False, index=df_audit.index)
        if c1 is not None: cond = cond | c1.notna()
        if c2 is not None: cond = cond | c2.notna()
        if c3 is not None: cond = cond | c3.notna()
        mask_A = mask_ausente & cond
    except Exception: pass

    cols_subj = [c for c in ['5.1.2 Satisfação', '6.3 Enfrenta problema', '6.7 Quantas vezes'] if c in df_audit.columns]
    mask_B = mask_ausente & df_audit[cols_subj].notna().any(axis=1) if cols_subj else pd.Series(False, index=df_audit.index)

    col_ben = df_audit.get('4.4 Quais benefícios sociais ou previdenciários os residentes recebem?_field')
    mask_C = (pd.to_numeric(df_audit['Renda_Familiar_Num'], errors='coerce') > 4000) & col_ben.astype(str).str.contains('Bolsa Família|BPC', na=False, case=False) if col_ben is not None else pd.Series(False, index=df_audit.index)

    col_mora, col_cap, col_pos = '4.2 Número de moradores na residência_field', "6.2.1 Qual a capacidade do reservatório ou caixa d'água?_field", "6.2 Possui reservatório de água potável ou caixa d'água?_field"
    mask_D = (pd.to_numeric(df_audit[col_mora], errors='coerce') > 6) & (pd.to_numeric(df_audit[col_cap], errors='coerce') <= 250) & (df_audit[col_pos].astype(str).str.strip().str.upper() == 'SIM') if all(c in df_audit.columns for c in [col_mora, col_cap, col_pos]) else pd.Series(False, index=df_audit.index)

    col_log = next((c for c in df_audit.columns if '1.6 Logradouro' in c or (c.lower().startswith('1.6') and 'logradouro' in c.lower())), None)
    mask_E = pd.Series(False, index=df_audit.index)
    if col_log:
        log_upper = df_audit[col_log].astype(str).str.strip().str.upper()
        mask_E = (log_upper != '') & (log_upper != 'NAN') & ~log_upper.apply(lambda v: any(v.startswith(t) for t in ('RUA', 'AVENIDA', 'AV ', 'AV.', 'ESTRADA', 'RODOVIA', 'ALAMEDA', 'TRAVESSA', 'PRAÇA', 'PRACA', 'LARGO', 'VIELA', 'SERVIDÃO', 'SERVIDAO', 'CONDOMÍNIO', 'CONDOMINIO', 'SETOR', 'QUADRA', 'QD ', 'QD.', 'LOT', 'SÍTIO', 'SITIO', 'FAZENDA', 'CHÁCARA', 'CHACARA', 'VIA ', 'RAMAL', 'BECO')) if isinstance(v, str) else False)

    col_bairro   = next((c for c in df_audit.columns if '1.5 Bairro' in c or (c.lower().startswith('1.5') and 'bairro' in c.lower())), None)
    col_municipio = next((c for c in df_audit.columns if '1.4 Município' in c or (c.lower().startswith('1.4') and 'munic' in c.lower())), None)
    mask_F = pd.Series(False, index=df_audit.index)
    cache_bairros = {}  

    # 🎯 OTIMIZAÇÃO: Validação externa desligada por padrão. Se ligada, timeout de 1s e cache em memória.
    if VALIDAR_BAIRRO_OSM:
        print("   🌐 Validando bairros via OpenStreetMap (pode demorar)...")
        def validar_bairro_osm(bairro, municipio):
            chave = (bairro, municipio)
            if chave in cache_bairros: return cache_bairros[chave]
            try:
                url, params = "https://nominatim.openstreetmap.org/search", {'q': f"{bairro}, {municipio}, São Paulo, Brasil", 'format': 'json', 'limit': 1, 'addressdetails': 1}
                resp = requests.get(url, params=params, timeout=1, headers={"User-Agent": "DeepAuditor/1.0"})
                if resp.status_code == 200 and len(resp.json()) > 0:
                    cache_bairros[chave] = True; return True
                cache_bairros[chave] = False; return False
            except Exception:
                cache_bairros[chave] = True; return True   # Na falha, aceita para não travar

        if col_bairro:
            bairro_serie = df_audit[col_bairro].astype(str).str.strip()
            mask_validos = ~(bairro_serie == '') & ~(bairro_serie.str.upper() == 'NAN') & ~bairro_serie.str.fullmatch(r'\d+', na=False)
            pares_unicos = df_audit.loc[mask_validos, [col_bairro, col_municipio or col_bairro]].drop_duplicates() if col_municipio else None
            osm_invalidos = set()
            if col_municipio and pares_unicos is not None:
                total_pares = len(pares_unicos)
                for i, (_, row_par) in enumerate(pares_unicos.iterrows()):
                    b, m = str(row_par[col_bairro]).strip(), str(row_par[col_municipio]).strip()
                    if not validar_bairro_osm(b, m):
                        osm_invalidos.add((b, m))
                    if (i+1) % 50 == 0:
                        print(f"      Validados {i+1}/{total_pares} bairros...")
                mask_F_osm = mask_validos & df_audit.apply(lambda r: (str(r[col_bairro]).strip(), str(r[col_municipio]).strip()) in osm_invalidos, axis=1)
            else:
                mask_F_osm = pd.Series(False, index=df_audit.index)
            mask_F = ((bairro_serie == '') | (bairro_serie.str.upper() == 'NAN') | bairro_serie.str.fullmatch(r'\d+', na=False) | mask_F_osm)
    else:
        # Sem validação externa, a máscara F só captura vazios/números
        if col_bairro:
            bairro_serie = df_audit[col_bairro].astype(str).str.strip()
            mask_F = (bairro_serie == '') | (bairro_serie.str.upper() == 'NAN') | bairro_serie.str.fullmatch(r'\d+', na=False)

    col_cg_pos, col_cg_l, col_ci = next((c for c in df_audit.columns if '6.5 Possui caixa de gordura' in c), None), next((c for c in df_audit.columns if '6.5.1' in c), None), next((c for c in df_audit.columns if '6.6 Possui caixa de inspeção' in c), None)
    msg_G = pd.Series("", index=df_audit.index)
    if col_cg_pos and col_cg_l and '4.2 Número de moradores na residência_field' in df_audit.columns:
        possui_cg, litros, moradores = df_audit[col_cg_pos].astype(str).str.strip().str.upper() == 'SIM', pd.to_numeric(df_audit[col_cg_l], errors='coerce'), pd.to_numeric(df_audit['4.2 Número de moradores na residência_field'], errors='coerce').fillna(0)
        col_cat = next((c for c in df_audit.columns if '3.1 Categoria de uso' in c), None)
        eh_empresa = df_audit[col_cat].astype(str).str.upper().str.contains('EMPRESA|PRODUÇÃO|COMERCIAL|INSTITUCIONAL', na=False) if col_cat else pd.Series(False, index=df_audit.index)
        mask_G4 = possui_cg & ~(df_audit[col_ci].astype(str).str.strip().str.upper() == 'SIM') if col_ci else pd.Series(False, index=df_audit.index)
        msg_G = (np.where(possui_cg & litros.isna(), "Mód 6.5: L vazio. ", "") + np.where(possui_cg & (litros > 10) & (moradores <= 5) & ~eh_empresa, "Mód 6.5: Vol excedente. ", "") + np.where(possui_cg & (litros < 5) & (moradores > 5), "Mód 6.5: Vol fraco. ", "") + np.where(mask_G4, "Mód 6.6: Sem cx inspeção. ", ""))
        msg_G = pd.Series(msg_G, index=df_audit.index).str.strip()

    c_san, c_sat_a, c_sat_e = next((c for c in df_audit.columns if '6.1 Recebe serviço' in c), None), next((c for c in df_audit.columns if '6.1.3' in c), None), next((c for c in df_audit.columns if '6.1.4' in c), None)
    msg_H = pd.Series("", index=df_audit.index)
    if c_san:
        nao_rec = df_audit[c_san].astype(str).str.strip().str.upper().isin(['NÃO', 'NAO'])
        msg_H = (np.where(nao_rec & df_audit[c_sat_a].notna() if c_sat_a else False, "Mód 6.1.3: Avaliou água sem ter. ", "") + np.where(nao_rec & df_audit[c_sat_e].notna() if c_sat_e else False, "Mód 6.1.4: Avaliou esgoto sem ter. ", ""))
        msg_H = pd.Series(msg_H).str.strip()

    msg_I  = pd.Series("", index=df_audit.index)
    col_J, col_K = next((c for c in df_audit.columns if 'latitude' in c.lower() and str(c).startswith('1.3')), None), next((c for c in df_audit.columns if 'longitude' in c.lower() and str(c).startswith('1.3')), None)
    col_Y, col_Z = next((c for c in df_audit.columns if 'latitude' in c.lower() and str(c).startswith('2.2')), None), next((c for c in df_audit.columns if 'longitude' in c.lower() and str(c).startswith('2.2')), None)
    col_AN, col_AO = achar_col_gps(df_audit, 'latitude'), achar_col_gps(df_audit, 'longitude')

    if all(c is not None for c in [col_J, col_K, col_Y, col_Z, col_AN, col_AO]):
        laJ, loK, laY, loZ, laAN, loAO = df_audit[col_J].apply(lambda v: limpar_coordenada(v, 'lat')), df_audit[col_K].apply(lambda v: limpar_coordenada(v, 'lon')), df_audit[col_Y].apply(lambda v: limpar_coordenada(v, 'lat')), df_audit[col_Z].apply(lambda v: limpar_coordenada(v, 'lon')), df_audit[col_AN].apply(lambda v: limpar_coordenada(v, 'lat')), df_audit[col_AO].apply(lambda v: limpar_coordenada(v, 'lon'))
        dist_JK_YZ = pd.Series([calcular_distancia(a,b,c,d) for a,b,c,d in zip(laJ, loK, laY, loZ)], index=df_audit.index)
        dist_JK_AN = pd.Series([calcular_distancia(a,b,c,d) for a,b,c,d in zip(laJ, loK, laAN, loAO)], index=df_audit.index)
        dist_YZ_AN = pd.Series([calcular_distancia(a,b,c,d) for a,b,c,d in zip(laY, loZ, laAN, loAO)], index=df_audit.index)
        msg_I = (np.where(dist_JK_YZ > LIMITE_TRIANGULACAO, dist_JK_YZ.apply(lambda d: f"GPS P/I divergem ({int(d)}m). " if pd.notna(d) else ""), "") + np.where(dist_JK_AN > LIMITE_TRIANGULACAO, dist_JK_AN.apply(lambda d: f"GPS P/R divergem ({int(d)}m). " if pd.notna(d) else ""), "") + np.where(dist_YZ_AN > LIMITE_TRIANGULACAO, dist_YZ_AN.apply(lambda d: f"GPS I/R divergem ({int(d)}m). " if pd.notna(d) else ""), ""))
        msg_I = pd.Series(msg_I).str.strip()

    col_cod_geral = detectar_coluna(df_audit, ['codigo_unico', 'code deep', 'código'])
    msg_K = []
    for idx, row in df_audit.iterrows():
        erro_mun = ""
        if col_cod_geral and col_AN and col_AO:
            cod = str(row.get(col_cod_geral, '')).strip()
            lat = limpar_coordenada(row.get(col_AN), 'lat')
            lon = limpar_coordenada(row.get(col_AO), 'lon')
            if cod and lat is not None and lon is not None:
                prefixo = re.split(r'[_0-9]', cod.split('-')[0])[0].upper()
                coords_alvo = CENTROIDES_MUNICIPIOS.get(prefixo)
                if coords_alvo:
                    distancias = [calcular_distancia(lat, lon, c_lat, c_lon) for c_lat, c_lon in coords_alvo]
                    dist_validas = [d for d in distancias if d is not None]
                    if dist_validas:
                        menor_dist = min(dist_validas)
                        if menor_dist > 40000:
                            erro_mun = f"🚨 Mód GPS: Código de {prefixo} preenchido em município distante ({int(menor_dist/1000)}km). "
        msg_K.append(erro_mun)

    mask_J = df_audit[col_enc].astype(str).str.strip().isin(['Não', 'Sim, mas o responsável não estava presente']) if col_enc else pd.Series(False, index=df_audit.index)

    erros_df = pd.DataFrame({
        'a': np.where(mask_A, "Mód 3/4: Sem morador.", ""), 'b': np.where(mask_B, "Mód 5/6: Respondeu sem morador.", ""), 
        'c': np.where(mask_C, "Mód Sócio: Renda alta c/ benefício.", ""), 'd': np.where(mask_D, "Mód Infra: Reservatório pequeno.", ""), 
        'e': np.where(mask_E, "Mód C: Rua não oficial.", ""), 'g': msg_G.values, 'h': msg_H.values, 
        'i': msg_I.values, 'j': np.where(mask_J, "Mód B: Agendada Revisita.", ""), 'k': pd.Series(msg_K, index=df_audit.index).values
    })
    
    erros_finais = erros_df.apply(lambda r: ' | '.join(v for v in r if str(v).strip()), axis=1)
    ta, tn = df_audit['Erros_Logicos'].str.strip() != '', erros_finais != ''
    df_audit['Erros_Logicos'] = np.select([ta & tn, ~ta & tn], [df_audit['Erros_Logicos'].str.strip() + ' | ' + erros_finais, erros_finais], default=df_audit['Erros_Logicos'])
    df_audit['Erros_Logicos'] = df_audit['Erros_Logicos'].str.replace(r'\|\s*\|', '|', regex=True).str.replace(r'^\s*\|\s*', '', regex=True).str.replace(r'\s*\|\s*$', '', regex=True).str.strip()
    return df_audit

# =============================================================================
# 7. EXECUÇÃO E MOTOR DE VALIDAÇÃO VISUAL
# =============================================================================


def processar_censo(arquivo_entrada, progresso=None):
    """Executa a lógica oficial e devolve os relatórios em memória."""
    def atualizar(percentual, etapa, detalhe=""):
        if progresso:
            progresso(percentual, etapa, detalhe)

    atualizar(0.02, "Lendo a base", "Identificando o formato e carregando os registros")
    print("\n📂 Lendo arquivo de dados...")
    try:
        posicao_inicial = arquivo_entrada.tell() if hasattr(arquivo_entrada, "tell") else None
        assinatura = arquivo_entrada.read(4) if hasattr(arquivo_entrada, "read") else b""
        if posicao_inicial is not None:
            arquivo_entrada.seek(posicao_inicial)

        # Arquivos XLSX são contêineres ZIP e começam com a assinatura PK.
        # Para caminhos locais, a extensão também é considerada.
        nome_arquivo = str(getattr(arquivo_entrada, "name", arquivo_entrada)).lower()
        eh_excel = (
            assinatura.startswith(b"PK")
            or assinatura == b"\xd0\xcf\x11\xe0"
            or nome_arquivo.endswith((".xlsx", ".xls"))
        )

        if eh_excel:
            df = pd.read_excel(arquivo_entrada, dtype=str)
        else:
            conteudo = arquivo_entrada.read() if hasattr(arquivo_entrada, "read") else None
            if conteudo is None:
                raise ValueError("O conteúdo do CSV não pôde ser carregado.")
            df = ler_csv_flexivel(conteudo)
    except Exception as e:
        raise ValueError(f"Não foi possível ler o arquivo Excel ou CSV: {e}") from e

    # ✏️ NOVO: corrige o "NÃ£o" (mojibake) já na entrada, antes de qualquer regra.
    df = reparar_mojibake(df)
    print("   🧹 Encoding reparado (mojibake 'NÃ£o' -> 'Não').")

    df = identificar_duplicatas(df)
    atualizar(0.10, "Auditando a qualidade", f"{len(df):,} registros encontrados".replace(",", "."))
    print("🔍 Iniciando auditoria de qualidade...")
    df = auditar_base_visitas(df)

    col_codigo = detectar_coluna(df, ['codigo_unico', 'code deep', 'código'])
    col_agente = detectar_coluna(df, ['agente', 'pesquisador', 'entrevistador'])
    col_nome_agente = next((c for c in df.columns if 'nome do agente' in str(c).lower() or ('nome' in str(c).lower() and 'agente' in str(c).lower())), None)

    if not col_agente and col_nome_agente: col_agente = col_nome_agente
    if not col_codigo: raise ValueError("Não há coluna de Código Único.")

    col_lat, col_lon = achar_col_gps(df, 'latitude'), achar_col_gps(df, 'longitude')

    atualizar(0.20, "Localizando imagens", "Procurando links de fotos em todas as colunas")
    print("\n🔍 Escaneando a planilha inteira atrás de Links de Imagens Reais...")
    df['Links_Detectados'] = df.apply(extrair_fotos_da_linha, axis=1)
    df['URL_Prop'] = df['Links_Detectados'].apply(lambda x: x[0] if len(x) > 0 else "")
    df['URL_Imov'] = df['Links_Detectados'].apply(lambda x: x[1] if len(x) > 1 else "")

    print(f"   📸 O Scanner Absoluto detectou {df['Links_Detectados'].apply(len).sum()} links válidos.")

    df['ID_Completo'] = df[col_codigo].astype(str).str.strip().str.upper()
    df['ID_Base']     = df['ID_Completo'].str.split('-').str[0]
    df['Lat_Num']     = df[col_lat].apply(lambda v: limpar_coordenada(v, 'lat')) if col_lat else np.nan
    df['Lon_Num']     = df[col_lon].apply(lambda v: limpar_coordenada(v, 'lon')) if col_lon else np.nan
    # =============================================================================
    # 6. MOTOR DE APROVAÇÃO DE BACKOFFICE (REGRA DA FUNDEG)
    # =============================================================================
    atualizar(0.28, "Aplicando regras de backoffice", "Classificando aprovações e pendências")
    print("\n📝 Aplicando diretrizes de Análise e Aprovação de Visitas...")

    # Utiliza a função detectar_coluna que já existe no rodar.py para achar os campos
    col_alguem = detectar_coluna(df, ['alguém foi encontrado', 'alguem foi encontrado'])
    col_motivo = detectar_coluna(df, ['motivo de alguém não ter sido', 'motivo'])

    def avaliar_aprovacao(bloco):
        n_visitas = len(bloco)
        alguem_serie = bloco[col_alguem].astype(str).str.strip().str.lower() if col_alguem else pd.Series(dtype=str)
        motivo_serie = bloco[col_motivo].astype(str).str.lower() if col_motivo else pd.Series(dtype=str)

        # Varredura completa no formulário inteiro (incluindo observações)
        texto_global = bloco.fillna('').apply(lambda row: ' '.join(map(str, row)).lower(), axis=1)
        if texto_global.apply(lambda txt: "abandonado" in txt).any():
            return "Aprovar - Abandonado"
        if texto_global.apply(lambda txt: any(t in txt for t in ["em construção", "em construcao"])).any():
            return "Aprovar - Em Construção"
        if texto_global.apply(lambda txt: any(t in txt for t in ["uso temporário", "uso temporario", "veraneio"])).any():
            return "Aprovar - Uso Temporário / Veraneio"

        if not motivo_serie.empty:
            if motivo_serie.str.contains("propriedade inexistente", regex=True, na=False).any():
                return "Aprovar - Propriedade Inexistente"
            if motivo_serie.str.contains("propriedade totalmente produtiva", regex=True, na=False).any():
                return "Aprovar - Propriedade Produtiva"

        if not alguem_serie.empty:
            if alguem_serie.isin(["sim"]).any():
                return "Aprovar - Visita"
            if alguem_serie.isin(["sim, mas se recusou a responder"]).any():
                return "Aprovar - Recusa"

        motivo_pendencia = ""
        termos_ausencia_alguem = [
            "sim, mas o responsável não estava presente",
            "sim, mas o responsavel nao estava presente",
            "não", "nao"
        ]

        if not alguem_serie.empty and alguem_serie.isin(termos_ausencia_alguem).any():
            motivo_pendencia = "Morador Ausente"

        if not motivo_serie.empty:
            if motivo_serie.fillna('').astype(str).str.contains("edificação longe da porteira", na=False).any():
                motivo_pendencia = "Edificação Longe da Porteira"
            elif motivo_serie.fillna('').astype(str).str.contains("não foi localizado", na=False).any():
                motivo_pendencia = "Morador Não Localizado"
            elif motivo_serie.fillna('').astype(str).str.contains("não foi encontrada porteira", na=False).any():
                motivo_pendencia = "Não Encontrou Porteira"

        if motivo_pendencia:
            if n_visitas >= 3:
                return f"Aprovar - {motivo_pendencia} (Histórico >= 3)"
            else:
                return f"Reprovar - {motivo_pendencia}"

        return "Reprovar - Fora do Padrão / Pendente"

    # Aplica a inteligência agrupando pelo 'ID_Base' que foi gerado logo acima
    mapa_aprov = {}
    for base, grupo in df.groupby('ID_Base'):
        mapa_aprov[base] = avaliar_aprovacao(grupo)

    df['Status_Aprovacao_Backoffice'] = df['ID_Base'].map(mapa_aprov)

    print("\n🕵️  Iniciando Motor de Validação Visual...")
    pares = set()

    # 🎯 OTIMIZAÇÃO: Extração de URLs isolada de forma vetorizada instantânea (salva muita CPU e RAM)
    urls_dl = set(df['URL_Prop'].dropna().tolist() + df['URL_Imov'].dropna().tolist())
    urls_dl.discard("")

    for id_base, grupo in df.groupby('ID_Base'):
        if str(id_base).upper().startswith('SEM_ID') or str(id_base) in ('', 'NAN'): continue
        idxs = grupo.index.tolist()
        for ia, ib in combinations(idxs, 2):
            pares.add((ia, ib))

    df_gps = df.dropna(subset=['Lat_Num', 'Lon_Num']).copy()
    if not df_gps.empty:
        print(f"   🌐 Mapeando área com raio reduzido/otimizado de {int(RAIO_BUSCA_VALIDACAO)}m...")
        coords = np.radians(df_gps[['Lat_Num', 'Lon_Num']].values)
        tree = BallTree(coords, metric='haversine')
        indices_vizinhos = tree.query_radius(coords, r=(RAIO_BUSCA_VALIDACAO / 6_371_000))
        for i, vizinhos in enumerate(indices_vizinhos):
            idx_A = df_gps.index[i]
            for j in vizinhos:
                if i < j: 
                    idx_B = df_gps.index[j]
                    if df.at[idx_A, 'ID_Base'] != df.at[idx_B, 'ID_Base']:
                        pares.add((idx_A, idx_B))

    print(f"   📊 Total de pares a serem analisados: {len(pares)}")

    cache_hash = {}
    if urls_dl:
        atualizar(0.35, "Baixando e analisando as fotos", f"0 de {len(urls_dl):,} imagens processadas".replace(",", "."))
    else:
        atualizar(0.70, "Análise de imagens concluída", "Nenhum link de imagem encontrado")
    pre_carregar_hashes(urls_dl, cache_hash, progresso)

    resultado = {idx: {'status': '✅ OK', 'detalhe': '', 'codigo_suspeito': '', 'similaridade': -1.0, 'desmembramento': ''} for idx in df.index}
    df_dict = df.to_dict('index')

    # Loop com progresso
    total_pares = len(pares)
    print("   🔬 Analisando pares visualmente...")
    for i, (ia, ib) in enumerate(pares):
        if progresso and (i == 0 or (i + 1) == total_pares or (i + 1) % max(1, total_pares // 100) == 0):
            fracao = (i + 1) / total_pares if total_pares else 1
            atualizar(
                0.70 + (0.20 * fracao),
                "Comparando registros",
                f"{i + 1:,} de {total_pares:,} pares analisados".replace(",", "."),
            )
        if i % 1000 == 0 and i > 0:
            print(f"      {i}/{total_pares} pares processados...")
        row_A, row_B = df_dict[ia], df_dict[ib]
        hp_A, hi_A = cache_hash.get(row_A['URL_Prop']), cache_hash.get(row_A['URL_Imov'])
        hp_B, hi_B = cache_hash.get(row_B['URL_Prop']), cache_hash.get(row_B['URL_Imov'])

        _sims = [s for s in [calcular_similaridade(hp_A, hp_B), calcular_similaridade(hi_A, hi_B), calcular_similaridade(hp_A, hi_B), calcular_similaridade(hi_A, hp_B)] if s >= 0]
        sim_max = max(_sims + [0])
        # ✏️ NOVO (Regra 1): menor similaridade denuncia a foto intrusa.
        sim_min = min(_sims) if _sims else 0.0
        # ✏️ NOVO (Regra 3): registra o % de similaridade nas duas pontas do par.
        for _ix in (ia, ib):
            if sim_max > resultado[_ix].get('similaridade', -1.0):
                resultado[_ix]['similaridade'] = round(float(sim_max), 2)
        dist = calcular_distancia(row_A['Lat_Num'], row_A['Lon_Num'], row_B['Lat_Num'], row_B['Lon_Num'])
        
        mesmo_codigo = row_A['ID_Base'] == row_B['ID_Base']
        mesmo_agente = (str(row_A.get(col_agente, 'A')).strip() == str(row_B.get(col_agente, 'B')).strip()) if col_agente else False

        inconsistencia_clara, inconsistencia_parcial = False, False
        if dist is None: inconsistencia_clara, inconsistencia_parcial = sim_max >= 90.0, sim_max >= 75.0
        elif dist <= 10.0: inconsistencia_clara, inconsistencia_parcial = sim_max >= 55.0, sim_max >= 45.0
        elif dist <= LIMITE_MESMO_LOCAL: inconsistencia_clara, inconsistencia_parcial = sim_max >= 75.0, sim_max >= 60.0
        else: 
            if mesmo_agente: inconsistencia_clara, inconsistencia_parcial = sim_max >= 80.0, sim_max >= 45.0 
            else: inconsistencia_clara, inconsistencia_parcial = sim_max >= 85.0, sim_max >= 70.0

        for idx_p, row_p, row_o in [(ia, row_A, row_B), (ib, row_B, row_A)]:
            eh_revisita = '-' in str(row_p['ID_Completo']) or '-' in str(row_o['ID_Completo'])
            cod_o = str(row_o.get('ID_Completo', ''))

            if mesmo_codigo:
                # ✏️ DESMEMBRAMENTO (definição do cliente): MESMO Code Deep, mas as
                # visitas representam imóveis diferentes. Aponta se houver QUALQUER
                # UMA das divergências: imagem diferente (sim < SIM_DESMEMBRAMENTO)
                # OU coordenada diferente (dist > DIST_DESMEMBRAMENTO). Fotos
                # iguais/parecidas + mesmo local = revisita normal -> SILÊNCIO.
                imagem_diferente = (len(_sims) > 0) and (sim_max < SIM_DESMEMBRAMENTO)
                coord_diferente = (dist is not None) and (dist > DIST_DESMEMBRAMENTO)
                if imagem_diferente or coord_diferente:
                    _motivos = []
                    if imagem_diferente: _motivos.append(f'imagem diferente ({sim_max:.1f}% sim)')
                    if coord_diferente:  _motivos.append(f'coordenada diferente ({int(dist)}m)')
                    resultado[idx_p]['desmembramento'] = cod_o
                    aplicar_status(resultado, idx_p, '🚨 DESMEMBRAMENTO - FOTO INTRUSA',
                                   'Mesmo Code Deep com ' + ' e '.join(_motivos) + f' vs visita {cod_o}.', cod_o)
                # Fotos iguais/parecidas + mesmo local => SILÊNCIO.
            else:
                # ✏️ DUPLICATA (definição do cliente): Code Deeps DIFERENTES com o
                # MESMO imóvel. Exige AS DUAS: coordenada próxima (<= DIST_DUPLICATA)
                # E imagem parecida (>= SIM_DUPLICATA).
                coord_proxima = (dist is not None) and (dist <= DIST_DUPLICATA)
                imagem_parecida = (len(_sims) > 0) and (sim_max >= SIM_DUPLICATA)
                if coord_proxima and imagem_parecida:
                    aplicar_status(resultado, idx_p, '🚨 DUPLICATA - MESMO IMÓVEL EM OUTRO CÓDIGO', f'Mesmo imóvel do código {cod_o} ({sim_max:.1f}% sim, {int(dist)}m).', cod_o)
                elif inconsistencia_parcial: aplicar_status(resultado, idx_p, '👀 REVISÃO - FOTO PARECIDA', f'Possível ângulo diferente: {cod_o} ({sim_max:.1f}% sim).', cod_o)
                elif dist is not None and dist <= 10: aplicar_status(resultado, idx_p, '⚠️ ALERTA DE PROXIMIDADE', f'GPS muito colado ({int(dist)}m) a outro cadastro.', cod_o)

    atualizar(0.92, "Finalizando as classificações", "Consolidando alertas e inconsistências")
    print("   ✅ Análise visual concluída. Finalizando status...")
    for idx, row in df_dict.items():
        if resultado[idx]['status'] != '✅ OK': continue
        hp, hi = cache_hash.get(row['URL_Prop']), cache_hash.get(row['URL_Imov'])
        sem_foto_valida = (hp is None or isinstance(hp, str)) and (hi is None or isinstance(hi, str))
        sem_gps = pd.isna(row['Lat_Num'])

        if sem_foto_valida:
            if (row['URL_Prop'] != "" or row['URL_Imov'] != "") and (isinstance(hp, str) or isinstance(hi, str)): aplicar_status(resultado, idx, '❌ ERRO DE LINK', 'Bloqueio de servidor (SSL/Permissão).')
            elif sem_gps: aplicar_status(resultado, idx, '📋 INCOMPLETO', 'Sem fotos e sem GPS.')
            else: aplicar_status(resultado, idx, '⚠️ AVISO', 'Scanner não encontrou links de imagem.')
        elif not sem_foto_valida and sem_gps: aplicar_status(resultado, idx, '⚠️ AVISO - FOTO OK, SEM GPS', 'Sem coordenadas GPS.')

    df['Status_Validacao'] = [resultado[i]['status']  for i in df.index]
    df['Codigo_Suspeito']   = [resultado[i]['codigo_suspeito'] for i in df.index]
    df['Detalhe_Inconsistencia']    = [resultado[i]['detalhe'] for i in df.index]

    # ✏️ NOVO (Regra 3): percentual exato de similaridade encontrado.
    df['Percentual_Similaridade'] = [
        (f"{resultado[i]['similaridade']:.2f}%" if resultado[i]['similaridade'] >= 0 else "")
        for i in df.index
    ]

    # ✏️ CORRIGIDO: o desmembramento é apontado sempre que detectado. O sufixo
    # "-1"/"-2" no ID é NORMAL (indica a ordem da visita no Code Deep) e NÃO deve
    # apagar o apontamento — conforme definição do cliente.
    def _desmemb_linha(i):
        desm = resultado[i].get('desmembramento', '')
        if not desm:
            return ""
        return f"Desmembramento: imóvel/coordenada diferente vs visita {desm}"

    df['Apontamento_Desmembramento'] = [_desmemb_linha(i) for i in df.index]


    # =============================================================================
    # 9. EXPORTAÇÃO EXCEL COM DICIONÁRIO
    # =============================================================================
    atualizar(0.97, "Preparando o resultado", "Montando as abas do relatório")
    print("\n📊 Gerando relatório Excel...")
    df_reprovados = df[(df['Erros_Logicos'] != '') | (~df['Status_Validacao'].isin({'✅ OK', '⚠️ AVISO', '⚠️ AVISO - FOTO OK, SEM GPS'}))].copy()

    cols_focal = ['ID_Completo', 'Status_Aprovacao_Backoffice', 'Duplicata', 'Status_Validacao', 'Codigo_Suspeito', 'Apontamento_Desmembramento', 'Percentual_Similaridade', 'Detalhe_Inconsistencia', 'Erros_Logicos']
    if col_nome_agente and col_nome_agente in df.columns: cols_focal.insert(0, col_nome_agente)
    elif col_agente and col_agente in df.columns: cols_focal.insert(0, col_agente)

    df_focal = df_reprovados[cols_focal]

    df_resumo = pd.DataFrame()
    if col_agente and col_agente in df.columns:
        df_resumo = df.groupby(col_agente)['Status_Validacao'].value_counts().unstack(fill_value=0).reset_index()
        df_resumo.columns.name = None
        if col_nome_agente and col_nome_agente != col_agente:
            nomes = df.groupby(col_agente)[col_nome_agente].first().reset_index()
            nomes.columns = [col_agente, 'Nome_Agente']
            df_resumo = nomes.merge(df_resumo, on=col_agente, how='right')
            
        for filtro, nome in [(df['Erros_Logicos'] != '', 'Erros_Qualitativos'), (df['Status_Validacao'] == '🚨 DUPLICATA - MESMO IMÓVEL EM OUTRO CÓDIGO', 'Duplicatas_do_Sistema')]:
            tmp = df[filtro].groupby(df[col_agente]).size().reset_index(name=nome)
            tmp.columns = [col_agente, nome]
            df_resumo = df_resumo.merge(tmp, on=col_agente, how='left').fillna(0)

    df_legenda = pd.DataFrame([
        {"Módulo": "🔴 Validação Visual — Nível 1", "Status / Regra": "🚨 INCONSISTÊNCIA CRÍTICA", "Significado": "Imagem altamente similar detectada entre registros diferentes via Matriz Visual."},
        {"Módulo": "🔴 Validação Visual — Nível 1", "Status / Regra": "🚨 DESMEMBRAMENTO - FOTO INTRUSA", "Significado": "MESMO Code Deep cujas visitas têm imagem OU coordenada diferente entre si (imóveis distintos sob o mesmo código). Ver Apontamento_Desmembramento e Percentual_Similaridade."},
        {"Módulo": "🔴 Validação Visual — Nível 1", "Status / Regra": "🚨 DUPLICATA - MESMO IMÓVEL EM OUTRO CÓDIGO", "Significado": "Code Deeps DIFERENTES trazendo o MESMO imóvel: coordenadas próximas E imagens muito parecidas."},
        {"Módulo": "🟠 Validação Visual — Nível 2", "Status / Regra": "👀 REVISÃO - FOTO PARECIDA", "Significado": "Fotos suspeitas cruzadas. Podem ser do mesmo local em ângulos distintos."},
        {"Módulo": "🟠 Validação Visual — Nível 2", "Status / Regra": "⚠️ ALERTA DE PROXIMIDADE", "Significado": "Agentes registraram códigos independentes dividindo o mesmo local (<= 10m)."},
        {"Módulo": "🟡 Validação Visual — Nível 3", "Status / Regra": "❌ ERRO DE LINK", "Significado": "Corte de servidor (Erro 403/SSL). Permissão negada para leitura da imagem."},
        {"Módulo": "🟡 Validação Visual — Nível 3", "Status / Regra": "📋 INCOMPLETO", "Significado": "Falta de GPS ou a base de dados não contém o link da imagem."},
        {"Módulo": "🔵 Validação Visual — Nível 4", "Status / Regra": "⚠️ AVISO - FOTO OK, SEM GPS", "Significado": "Hash visual validado, mas o formulário não obteve leitura de GPS."},
        {"Módulo": "🔵 Validação Visual — Nível 4", "Status / Regra": "⚠️ AVISO", "Significado": "Sinal GPS validado, porém não há fotos (link nulo/vazio)."},
        {"Módulo": "🟢 Validação Visual — Nível 5", "Status / Regra": "✅ OK", "Significado": "Cadastro conferido e ratificado pelo crivo matricial de qualidade."},
        {"Módulo": "🔍 Qualidade — Regra K", "Status / Regra": "Código de município trocado", "Significado": "O GPS aponta para uma região distante (>40km) da cidade referenciada na sigla do código único."}
    ])

    cols_frente = [c for c in ['ID_Completo', 'Status_Aprovacao_Backoffice', 'Duplicata', 'Status_Validacao', 'Codigo_Suspeito', 'Apontamento_Desmembramento', 'Percentual_Similaridade', 'Detalhe_Inconsistencia', 'Erros_Logicos', 'Bairro_GPS'] if c in df.columns]
    df_saida = df[cols_frente + [c for c in df.columns if c not in cols_frente and c not in {'Lat_Num', 'Lon_Num', 'URL_Prop', 'URL_Imov', 'ID_Base', 'Renda_Familiar_Num', '_cluster', 'Links_Detectados'}]]

    return {
        'Base_Consolidada': df_saida,
        'Resumo_Ponto_Focal': df_focal,
        'Dicionario_e_Regras': df_legenda,
        'Resumo_Por_Agente': df_resumo,
    }
