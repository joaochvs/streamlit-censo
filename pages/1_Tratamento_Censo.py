import concurrent.futures
from io import BytesIO
from pathlib import Path
import queue
import time
import unicodedata

import pandas as pd
import streamlit as st

from tratamento_censo import processar_censo


st.markdown(
    """
    <style>
    .block-container {max-width: 1180px; padding-top: 2rem; padding-bottom: 4rem;}
    .hero-censo {
        padding: 2rem 2.2rem; border-radius: 22px; margin-bottom: 1.4rem;
        background: linear-gradient(135deg, #102a43 0%, #174f5f 58%, #168c82 100%);
        box-shadow: 0 18px 45px rgba(0,0,0,.18); color: white;
    }
    .hero-censo h1 {font-size: 2.25rem; margin: 0 0 .5rem; color: white;}
    .hero-censo p {font-size: 1.02rem; margin: 0; opacity: .88; max-width: 760px;}
    div[data-testid="stFileUploader"] {padding: .25rem 0 .4rem;}
    div[data-testid="stButton"] button, div[data-testid="stDownloadButton"] button {
        min-height: 3rem; border-radius: 12px; font-weight: 700;
    }
    .metric-grid {display:grid; grid-template-columns:repeat(2,minmax(0,1fr)); gap:1rem; margin:1rem 0 1.5rem;}
    .metric-card {
        padding:1.25rem 1.4rem; border:1px solid rgba(128,128,128,.22);
        border-radius:16px; background:rgba(128,128,128,.055);
    }
    .metric-label {font-size:.82rem; opacity:.68; font-weight:700; text-transform:uppercase; letter-spacing:.04em;}
    .metric-value {font-size:2rem; line-height:1.15; font-weight:750; margin:.35rem 0;}
    .metric-note {font-size:.82rem; opacity:.68;}
    @media (max-width: 720px) {.metric-grid {grid-template-columns:1fr}.hero-censo{padding:1.5rem}.hero-censo h1{font-size:1.8rem}}
    </style>
    <div class="hero-censo">
      <h1>Tratamento Censo</h1>
      <p>Envie a base, acompanhe cada etapa da validação e baixe o relatório consolidado pronto para análise.</p>
    </div>
    """,
    unsafe_allow_html=True,
)




def gerar_excel(relatorios: dict[str, pd.DataFrame]) -> bytes:
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        for nome_aba, dados in relatorios.items():
            dados.to_excel(writer, sheet_name=nome_aba[:31], index=False)
            planilha = writer.sheets[nome_aba[:31]]
            planilha.freeze_panes = "A2"
            planilha.auto_filter.ref = planilha.dimensions
    return buffer.getvalue()


def calcular_erro_backoffice(base: pd.DataFrame) -> tuple[float | None, int, int]:
    def normalizar_texto(valor) -> str:
        texto = unicodedata.normalize("NFKD", str(valor))
        texto = "".join(c for c in texto if not unicodedata.combining(c)).lower()
        return "".join(c for c in texto if c.isalnum())

    colunas_normalizadas = {normalizar_texto(c): c for c in base.columns}
    col_status_correto = colunas_normalizadas.get("statusaprovacaobackoffice")
    candidatas_bko = [
        original
        for normalizada, original in colunas_normalizadas.items()
        if "backoffice" in normalizada and normalizada != "statusaprovacaobackoffice"
    ]

    # Seleciona pela presença real de "Reprovado". Isso evita depender de
    # espaços, acentos ou pequenas variações no nome da coluna de origem.
    contagens_reprovados = {}
    series_normalizadas = {}
    for coluna in candidatas_bko:
        serie = base[coluna].fillna("").map(normalizar_texto)
        series_normalizadas[coluna] = serie
        contagens_reprovados[coluna] = int(serie.str.startswith("reprov").sum())

    col_situacao_bko = (
        max(candidatas_bko, key=lambda c: contagens_reprovados[c])
        if candidatas_bko
        else None
    )

    if not col_status_correto or not col_situacao_bko:
        return None, 0, 0

    esperado = base[col_status_correto].fillna("").map(normalizar_texto)
    realizado = series_normalizadas[col_situacao_bko]
    deveria_aprovar = esperado.str.startswith("aprovar")
    reprovados_pelo_bko = realizado.str.startswith("reprov")
    recusas_que_deveriam_ser_aprovadas = reprovados_pelo_bko & deveria_aprovar

    total_avaliado = int(reprovados_pelo_bko.sum())
    total_erros = int(recusas_que_deveriam_ser_aprovadas.sum())
    taxa = (total_erros / total_avaliado * 100) if total_avaliado else 0.0
    return taxa, total_erros, total_avaliado


def executar_tratamento(arquivo_bytes: bytes, progresso):
    relatorios = processar_censo(BytesIO(arquivo_bytes), progresso=progresso)
    progresso(0.99, "Gerando o arquivo para download", "Gravando as abas do Excel")
    return relatorios, gerar_excel(relatorios)


with st.container(border=True):
    st.subheader("Enviar base para tratamento")
    st.caption("Formatos aceitos: Excel .xlsx ou .xls")
    arquivo = st.file_uploader(
        "Selecione o arquivo",
        type=["xlsx", "xls"],
        label_visibility="collapsed",
    )

if arquivo is None:
    st.info("📂 Selecione uma planilha para iniciar o tratamento.")
else:
    tamanho_mb = arquivo.size / (1024 * 1024)
    st.caption(f"✓ {arquivo.name}  •  {tamanho_mb:.1f} MB")

    if st.button("Processar arquivo", type="primary", use_container_width=True):
        executor = None
        try:
            inicio = time.perf_counter()
            eventos = queue.Queue()
            etapa_atual = [0.0, "Preparando o processamento", "Organizando o arquivo enviado"]

            def registrar_progresso(percentual, etapa, detalhe=""):
                eventos.put((percentual, etapa, detalhe))

            painel = st.status("Preparando o processamento", expanded=True)
            barra = st.progress(0)
            detalhes = st.empty()
            cronometro = st.empty()

            def atualizar_painel():
                while True:
                    try:
                        etapa_atual[:] = eventos.get_nowait()
                    except queue.Empty:
                        break
                percentual, etapa, detalhe = etapa_atual
                decorrido = time.perf_counter() - inicio
                minutos, segundos = divmod(int(decorrido), 60)
                painel.update(label=etapa, state="running", expanded=True)
                barra.progress(min(100, max(0, int(percentual * 100))))
                detalhes.markdown(f"**{detalhe}**" if detalhe else "")
                cronometro.markdown(f"⏱️ **Tempo decorrido** &nbsp; `{minutos:02d}:{segundos:02d}`")

            executor = concurrent.futures.ThreadPoolExecutor(max_workers=1)
            futuro = executor.submit(executar_tratamento, arquivo.getvalue(), registrar_progresso)

            # O processamento roda em segundo plano; este laço mantém o
            # cronômetro fluido mesmo durante downloads demorados.
            while not futuro.done():
                atualizar_painel()
                time.sleep(0.25)

            atualizar_painel()
            relatorios, excel_tratado = futuro.result()
            st.session_state["relatorios"] = relatorios
            st.session_state["excel_tratado"] = excel_tratado
            st.session_state["nome_saida"] = f"{Path(arquivo.name).stem}_tratado.xlsx"
            duracao = time.perf_counter() - inicio
            minutos, segundos = divmod(int(duracao), 60)
            barra.progress(100)
            detalhes.markdown("✅ **Arquivo tratado e pronto para download.**")
            cronometro.markdown(f"⏱️ **Tempo total** &nbsp; `{minutos:02d}:{segundos:02d}`")
            painel.update(label="Processamento concluído", state="complete", expanded=False)
        except Exception as erro:
            st.session_state.pop("relatorios", None)
            if "painel" in locals():
                painel.update(label="O processamento foi interrompido", state="error", expanded=True)
            st.error(f"Não foi possível processar o arquivo: {erro}")
        finally:
            if executor:
                executor.shutdown(wait=False)

if "relatorios" in st.session_state:
    relatorios = st.session_state["relatorios"]
    base = relatorios["Base_Consolidada"]
    taxa_erro, total_erros, total_avaliado = calcular_erro_backoffice(base)

    registros_formatados = f"{len(base):,}".replace(",", ".")
    if taxa_erro is None:
        taxa_formatada = "N/D"
        nota_taxa = "A coluna Situação Backoffice não foi encontrada."
    else:
        taxa_formatada = f"{taxa_erro:.1f}%"
        nota_taxa = (
            f"{total_erros:,} recusa(s) deveriam ser aprovadas entre "
            f"{total_avaliado:,} recusa(s) feitas pelo Backoffice.".replace(",", ".")
        )

    st.markdown(
        f"""
        <div class="metric-grid">
          <div class="metric-card">
            <div class="metric-label">Registros processados</div>
            <div class="metric-value">{registros_formatados}</div>
            <div class="metric-note">Linhas analisadas na base consolidada</div>
          </div>
          <div class="metric-card">
            <div class="metric-label">Taxa de erro do Backoffice</div>
            <div class="metric-value">{taxa_formatada}</div>
            <div class="metric-note">{nota_taxa}</div>
          </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    with st.container(border=True):
        st.subheader("Prévia da base consolidada")
        st.caption("Primeiros 100 registros. O arquivo para download contém a base completa.")
        st.dataframe(base.head(100), use_container_width=True, hide_index=True)

    st.download_button(
        "⬇️ Baixar arquivo tratado",
        data=st.session_state["excel_tratado"],
        file_name=st.session_state["nome_saida"],
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
        use_container_width=True,
    )
