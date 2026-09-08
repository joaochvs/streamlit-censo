import re
from pathlib import Path

import streamlit as st

from conferencia_pendentes import processar_conferencia
from tarefas_background import acompanhar_tarefa, obter_gerenciador_tarefas


st.markdown(
    """
    <style>
      .tool-hero{padding:2rem 2.2rem;margin-bottom:1.4rem;border-radius:22px;color:white;
        background:radial-gradient(circle at 88% 18%,rgba(167,139,250,.32),transparent 28%),
        linear-gradient(135deg,#312e81 0%,#6d28d9 58%,#7c3aed 100%);
        box-shadow:0 18px 42px rgba(2,6,23,.22)}
      .tool-hero h1{margin:0 0 .45rem;color:white;font-size:clamp(2rem,4vw,3rem);letter-spacing:-.04em}
      .tool-hero p{margin:0;max-width:800px;color:rgba(255,255,255,.88);line-height:1.6}
      .metric-grid{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:1rem;margin:1.2rem 0}
      .metric-card{padding:1.2rem 1.3rem;border:1px solid rgba(148,163,184,.24);border-radius:18px;background:rgba(30,41,59,.08)}
      .metric-label{color:#94a3b8;font-size:.82rem;font-weight:700;text-transform:uppercase}
      .metric-value{font-size:2rem;font-weight:800;margin:.35rem 0}
      @media(max-width:760px){.metric-grid{grid-template-columns:1fr}.tool-hero{padding:1.5rem}}
    </style>
    <div class="tool-hero">
      <h1>Conferência de Pendentes</h1>
      <p>Filtre os registros pendentes por município e gere o arquivo de conferência com controles, responsáveis e validações.</p>
    </div>
    """,
    unsafe_allow_html=True,
)

with st.container(border=True):
    st.subheader("Município e arquivo de origem")
    municipio = st.text_input("Município", placeholder="Ex.: Campinas", key="municipio_conferencia")
    arquivo = st.file_uploader(
        "Relatório Excel com a aba Base_Consolidada",
        type=["xlsx"],
        key="arquivo_conferencia",
    )

if arquivo is None:
    st.info("📂 Informe o município e selecione o relatório Excel.")
else:
    st.caption(f"✓ {arquivo.name}  •  {arquivo.size / (1024 * 1024):.1f} MB")
    if st.button("Gerar conferência", type="primary", use_container_width=True):
        if not municipio.strip():
            st.warning("Informe o município antes de iniciar o processamento.")
        else:
            for chave in ("conferencia_resultado", "conferencia_excel"):
                st.session_state.pop(chave, None)
            payload = (arquivo.getvalue(), municipio.strip())
            tarefa_id = obter_gerenciador_tarefas().iniciar(processar_conferencia, payload)
            st.session_state["tarefa_conferencia"] = tarefa_id
            st.query_params["tarefa_conferencia"] = tarefa_id

tarefa_id = st.session_state.get("tarefa_conferencia") or st.query_params.get("tarefa_conferencia")
if tarefa_id and "conferencia_resultado" not in st.session_state:
    try:
        resultado, excel = acompanhar_tarefa(tarefa_id)
        st.session_state["conferencia_resultado"] = resultado
        st.session_state["conferencia_excel"] = excel
        municipio_nome = municipio.strip() if municipio.strip() else "municipio"
        nome_seguro = re.sub(r"[^A-Za-z0-9_-]+", "_", municipio_nome).strip("_") or "municipio"
        st.session_state["conferencia_nome"] = f"Conferencia_{nome_seguro}.xlsx"
        st.session_state.pop("tarefa_conferencia", None)
        st.query_params.pop("tarefa_conferencia", None)
    except Exception as erro:
        st.session_state.pop("tarefa_conferencia", None)
        st.query_params.pop("tarefa_conferencia", None)
        if str(erro) == "A tarefa não está mais disponível no servidor.":
            st.info("A tarefa anterior expirou após uma reinicialização. Envie o arquivo novamente.")
        else:
            st.error(f"Não foi possível gerar a conferência: {erro}")

if "conferencia_resultado" in st.session_state:
    resultado = st.session_state["conferencia_resultado"]
    agentes = resultado["Nome do Agente"].dropna().astype(str).str.strip().replace("", None).dropna().nunique()
    st.success("Arquivo de conferência gerado com sucesso.")
    st.markdown(
        f"""
        <div class="metric-grid">
          <div class="metric-card"><div class="metric-label">Pendências selecionadas</div><div class="metric-value">{len(resultado):,}</div></div>
          <div class="metric-card"><div class="metric-label">Agentes encontrados</div><div class="metric-value">{agentes:,}</div></div>
        </div>
        """.replace(",", "."),
        unsafe_allow_html=True,
    )
    with st.container(border=True):
        st.subheader("Prévia da conferência")
        st.caption("Primeiros 100 registros. O Excel contém a base completa e as abas de controle.")
        st.dataframe(resultado.head(100), use_container_width=True, hide_index=True)
    st.download_button(
        "⬇️ Baixar arquivo de conferência",
        data=st.session_state["conferencia_excel"],
        file_name=st.session_state["conferencia_nome"],
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
        use_container_width=True,
    )
