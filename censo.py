# ── Imports ──────────────────────────────────────────────────────────────────

import re
import unicodedata
import datetime
from io import BytesIO

import pandas as pd
import streamlit as st
from rapidfuzz import process, fuzz
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph


# ── Constantes ────────────────────────────────────────────────────────────────

COLUNAS_PROD = [
    "Visitados", "Pesquisas", "Recusas", "Produtivos",
    "Abandonados", "Ausentes", "Em Construção"
]

SCORE_MATCH   = 95
SCORE_REVISAR = 75


# ── Funções utilitárias ───────────────────────────────────────────────────────

def normalizar(nome: str) -> str:
    """Normaliza um nome: minúsculas, sem acentos, sem caracteres especiais."""
    if pd.isna(nome):
        return ""
    nome = str(nome).lower()
    nome = unicodedata.normalize("NFKD", nome)
    nome = "".join(c for c in nome if not unicodedata.combining(c))
    nome = re.sub(r"[^a-z0-9 ]", "", nome)
    return " ".join(nome.split())


def quebrar_nome(nome: str) -> str:
    """Quebra nomes longos em duas linhas para o PDF."""
    partes = nome.split()
    if len(partes) > 3:
        return " ".join(partes[:3]) + "<br/>" + " ".join(partes[3:])
    return nome


def calcular_score(nome1: str, nome2: str, **kwargs) -> float:
    """Scorer combinado: máximo entre token_set_ratio e partial_ratio."""
    return max(
        fuzz.token_set_ratio(nome1, nome2),
        fuzz.partial_ratio(nome1, nome2),
    )


# ── Funções de carregamento (com cache) ───────────────────────────────────────

@st.cache_data(show_spinner="Carregando censo...")
def carregar_censo(arquivos_censo, data_inicio: datetime.date, data_fim: datetime.date) -> pd.DataFrame:
    """Lê um ou mais arquivos XLSX do censo e concatena, filtrando pelo período."""
    partes = []
    for arq in arquivos_censo:
        df = pd.read_excel(arq)
        df["Data do registro"] = pd.to_datetime(df["Data do registro"]).dt.date
        partes.append(df)

    df_total = pd.concat(partes, ignore_index=True)
    df_total = df_total.drop_duplicates()  # remove linhas duplicadas caso os arquivos se sobreponham
    return df_total[
        (df_total["Data do registro"] >= data_inicio) &
        (df_total["Data do registro"] <= data_fim)
    ].copy()


@st.cache_data(show_spinner="Carregando CSV...")
def carregar_csv(arquivo, data_inicio: datetime.date, data_fim: datetime.date) -> pd.DataFrame:
    df = pd.read_csv(arquivo)
    df["Data"] = pd.to_datetime(df["Data"]).dt.date
    df = df[
        (df["Data"] >= data_inicio) &
        (df["Data"] <= data_fim)
    ].copy()
    df["produtividade_total"] = df[COLUNAS_PROD].sum(axis=1)
    return df.reset_index(drop=True)


# ── Lógica de comparação ──────────────────────────────────────────────────────

def agregar_censo(df_censo: pd.DataFrame) -> pd.DataFrame:
    df = (
        df_censo.groupby("Nome do Agente")
        .size()
        .reset_index(name="produtividade_censo")
        .sort_values("produtividade_censo", ascending=False)
    )
    df["nome_normalizado"] = df["Nome do Agente"].apply(normalizar)
    return df


def agregar_csv(df_csv: pd.DataFrame) -> pd.DataFrame:
    df = (
        df_csv.groupby("Agente")["produtividade_total"]
        .sum()
        .reset_index(name="produtividade_informada")
        .sort_values("produtividade_informada", ascending=False)
    )
    df["nome_normalizado"] = df["Agente"].apply(normalizar)
    return df


def comparar(df_csv_prod: pd.DataFrame, df_censo_prod: pd.DataFrame) -> pd.DataFrame:
    """
    Faz o matching fuzzy entre agentes do formulário e do sistema.
    Também adiciona agentes presentes só no sistema (produtividade informada = 0).
    """
    nomes_censo = df_censo_prod["nome_normalizado"].tolist()
    resultado = []

    # ── Agentes que enviaram formulário ──────────────────────────────────────
    nomes_csv_matched = set()

    for _, linha in df_csv_prod.iterrows():
        melhor = process.extractOne(
            linha["nome_normalizado"], nomes_censo, scorer=calcular_score
        )

        if melhor is None or melhor[1] < SCORE_REVISAR:
            resultado.append({
                "Agente Informado":      linha["Agente"],
                "Agente Sistema":        "Não encontrado",
                "Similaridade":          round(melhor[1], 1) if melhor else 0,
                "Produtividade Informada": linha["produtividade_informada"],
                "Produtividade Sistema": 0,
                "Diferença":             linha["produtividade_informada"],
                "Status":                "❌ Sem correspondência",
            })
            continue

        nome_match, score, _ = melhor
        agente_sis = df_censo_prod[df_censo_prod["nome_normalizado"] == nome_match].iloc[0]
        nomes_csv_matched.add(nome_match)

        diferenca = linha["produtividade_informada"] - agente_sis["produtividade_censo"]
        status    = "✅ Match" if score >= SCORE_MATCH else "⚠ Revisar"

        resultado.append({
            "Agente Informado":      linha["Agente"],
            "Agente Sistema":        agente_sis["Nome do Agente"],
            "Similaridade":          round(score, 1),
            "Produtividade Informada": linha["produtividade_informada"],
            "Produtividade Sistema": agente_sis["produtividade_censo"],
            "Diferença":             diferenca,
            "Status":                status,
        })

    # ── Agentes só no sistema (não enviaram formulário) ──────────────────────
    sem_formulario = df_censo_prod[
        ~df_censo_prod["nome_normalizado"].isin(nomes_csv_matched)
    ]

    for _, ag in sem_formulario.iterrows():
        resultado.append({
            "Agente Informado":      "Não encontrado",
            "Agente Sistema":        ag["Nome do Agente"],
            "Similaridade":          0,
            "Produtividade Informada": 0,
            "Produtividade Sistema": ag["produtividade_censo"],
            "Diferença":             -ag["produtividade_censo"],
            "Status":                "📋 Só no sistema",
        })

    return pd.DataFrame(resultado)


# ── Geração de PDF ────────────────────────────────────────────────────────────

def gerar_pdf(df_divergencia: pd.DataFrame, df_resultado: pd.DataFrame, data_inicio: datetime.date, data_fim: datetime.date) -> BytesIO:
    buffer  = BytesIO()
    doc     = SimpleDocTemplate(buffer, pagesize=letter)
    styles  = getSampleStyleSheet()
    elementos = []

    def formatar_data_br(data):
        return data.strftime("%d/%m/%Y")

    def formatar_data_extenso(data):
        meses = [
            "janeiro", "fevereiro", "março", "abril", "maio", "junho",
            "julho", "agosto", "setembro", "outubro", "novembro", "dezembro"
        ]
        return f"{data.day} de {meses[data.month - 1]} de {data.year}"

    periodo = (
        formatar_data_extenso(data_inicio)
        if data_inicio == data_fim
        else f"{formatar_data_extenso(data_inicio)} a {formatar_data_extenso(data_fim)}"
    )

    
    # ── Seção 1: Divergências ─────────────────────────────────────────────────
    elementos.append(Paragraph(
        "<b><font size=14>Relatório de Divergência de Produtividade</font></b>",
        styles["Title"]
    ))
    elementos.append(Paragraph(" ", styles["Normal"]))  # espaço

    elementos.append(Paragraph(
        f"<b>Período de referência:</b><br/>{periodo}",
        styles["Normal"]
    ))

    elementos.append(Paragraph(" ", styles["Normal"])) 



    def montar_tabela(df):

        colunas = [
            "Agente Informado",
            "Agente Sistema",
            "Produtividade Informada",
            "Produtividade Sistema",
            "Diferença"
        ]

        df_t = df[colunas].copy()

        df_t.columns = [
            "Agente (Formulário)",
            "Agente (Sistema)",
            "Formulário",
            "Sistema",
            "Diferença"
        ]

        df_t["Agente (Formulário)"] = df_t["Agente (Formulário)"].apply(quebrar_nome)

        for col in ["Formulário", "Sistema", "Diferença"]:
            df_t[col] = df_t[col].round(0).astype(int)

        data_table = [df_t.columns.tolist()]

        for _, row in df_t.iterrows():
            data_table.append([
                Paragraph(str(row["Agente (Formulário)"]), styles["Normal"]),
                Paragraph(str(row["Agente (Sistema)"]), styles["Normal"]),
                str(row["Formulário"]),
                str(row["Sistema"]),
                str(row["Diferença"])
            ])

        tabela = Table(data_table, colWidths=[140, 140, 60, 60, 60])

        style = TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.darkblue),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),

            ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
            ('BACKGROUND', (0, 1), (-1, -1), colors.beige),

            ('ALIGN', (2, 1), (-1, -1), 'CENTER')
        ])

        # cores na diferença
        for i, row in enumerate(df_t.itertuples(), start=1):

            if row.Diferença < 0:
                cor = colors.blue
            elif row.Diferença > 0:
                cor = colors.red
            else:
                cor = colors.black

            style.add('TEXTCOLOR', (-1, i), (-1, i), cor)
            style.add('FONTNAME', (-1, i), (-1, i), 'Helvetica-Bold')

        tabela.setStyle(style)
        return tabela
    
   
    df_todos = df_resultado.sort_values(
        by="Diferença",
        ascending=False
    )


    elementos.append(montar_tabela(df_todos))


    doc.build(elementos)
    buffer.seek(0)
    return buffer


# ── Interface Streamlit ───────────────────────────────────────────────────────

st.set_page_config(layout="centered")
st.title("Divergência de Produtividade — Agentes")

# ── Seleção de período ────────────────────────────────────────────────────────

periodo = st.date_input(
    "Selecione o período (clique em duas datas para um range):",
    value=(datetime.date.today(), datetime.date.today()),
    format="DD/MM/YYYY",
)

if len(periodo) < 2:
    st.info("Selecione também a data final do período.")
    st.stop()

data_inicio, data_fim = periodo

# ── Upload dos arquivos ───────────────────────────────────────────────────────

col1, col2 = st.columns(2)


with col1:
    arquivos_censo = st.file_uploader(
        "📊 Base do sistema / censo (XLSX) ",
        type=["xlsx"],
        accept_multiple_files=True,
        key="censo",
    )

with col2:
    arquivo = st.file_uploader(
        "📄 Base de produtividade (CSV)",
        type=["csv"],
        key="csv",
    )


# ── Processamento ─────────────────────────────────────────────────────────────

if arquivo and arquivos_censo:

    st.caption(f"📂 {len(arquivos_censo)} arquivo(s) de censo carregado(s)")

    df_censo = carregar_censo(arquivos_censo, data_inicio, data_fim)
    df_csv   = carregar_csv(arquivo, data_inicio, data_fim)

    if df_censo.empty:
        st.warning(f"Nenhum registro encontrado no censo para o período selecionado.")
        st.stop()

    if df_csv.empty:
        st.warning(f"Nenhum registro encontrado no CSV para o período selecionado.")
        st.stop()

    df_censo_prod = agregar_censo(df_censo)
    df_csv_prod   = agregar_csv(df_csv)

    with st.spinner("Realizando comparação fuzzy..."):
        df_resultado = comparar(df_csv_prod, df_censo_prod)

    df_divergencia = (
        df_resultado[
            (df_resultado["Diferença"] > 0) &
            (df_resultado["Status"] != "❌ Sem correspondência")
        ]
        .sort_values("Diferença", ascending=False)
    )

    # ── Métricas rápidas ─────────────────────────────────────────────────────
    m1, m2, m3 = st.columns(3)
    m1.metric("Agentes no formulário", len(df_csv_prod))
    m2.metric("Agentes no sistema",    len(df_censo_prod))

    st.divider()

    # ── Tabelas ───────────────────────────────────────────────────────────────
    st.subheader("📋 Produtividade Informada")
    st.dataframe(df_csv_prod.drop(columns="nome_normalizado"), use_container_width=True)

    st.subheader("🚨 Quem informou mais do que fez")
    st.dataframe(df_divergencia, use_container_width=True)

    st.subheader("🔍 Comparativo Final")
    st.dataframe(
        df_resultado.sort_values(["Status", "Diferença"], ascending=[True, False]),
        use_container_width=True,
    )

    # ── Download PDF ──────────────────────────────────────────────────────────
    pdf = gerar_pdf(df_divergencia, df_resultado, data_inicio, data_fim)
    nome_pdf = (
        f"relatorio_divergencia_{data_inicio}.pdf"
        if data_inicio == data_fim
        else f"relatorio_divergencia_{data_inicio}_a_{data_fim}.pdf"
    )
    st.download_button(
        label="📄 Baixar Relatório PDF",
        data=pdf,
        file_name=nome_pdf,
        mime="application/pdf",
    )