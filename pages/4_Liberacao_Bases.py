from io import BytesIO
from pathlib import Path

import pandas as pd
import streamlit as st

from liberacao_bases import processar_liberacao
from tarefas_background import acompanhar_tarefa, obter_gerenciador_tarefas


st.markdown(
    """
    <style>
      .tool-hero {padding:2rem 2.2rem;margin-bottom:1.4rem;border-radius:22px;color:white;
        background:radial-gradient(circle at 88% 18%,rgba(251,191,36,.3),transparent 28%),
        linear-gradient(135deg,#3f1d0b 0%,#b45309 56%,#ea580c 100%);
        box-shadow:0 18px 42px rgba(2,6,23,.22)}
      .tool-hero h1{margin:0 0 .45rem;color:white;font-size:clamp(2rem,4vw,3rem);letter-spacing:-.04em}
      .tool-hero p{margin:0;max-width:800px;color:rgba(255,255,255,.88);line-height:1.6}
      .metric-grid{display:grid;grid-template-columns:repeat(3,minmax(0,1fr));gap:1rem;margin:1.2rem 0}
      .metric-card{padding:1.2rem 1.3rem;border:1px solid rgba(148,163,184,.24);border-radius:18px;
        background:rgba(30,41,59,.08)}
      .metric-label{color:#94a3b8;font-size:.82rem;font-weight:700;text-transform:uppercase}
      .metric-value{font-size:2rem;font-weight:800;margin:.35rem 0}
      @media(max-width:760px){.metric-grid{grid-template-columns:1fr}.tool-hero{padding:1.5rem}}
    </style>
    <div class="tool-hero">
      <h1>Liberação de Bases</h1>
      <p>Transforme o relatório consolidado em uma base de liberação pronta, preservando o template, as fórmulas e os controles operacionais.</p>
    </div>
    """,
    unsafe_allow_html=True,
)

with st.container(border=True):
    st.subheader("Enviar relatório para liberação")
    st.caption("O Excel deve conter a aba Base_Consolidada. O template oficial já está incorporado à ferramenta.")
    arquivo = st.file_uploader(
        "Selecione o relatório Excel",
        type=["xlsx"],
        key="arquivo_liberacao",
        label_visibility="collapsed",
    )

if arquivo is None:
    st.info("📂 Selecione um relatório Excel para iniciar a liberação.")
else:
    st.caption(f"✓ {arquivo.name}  •  {arquivo.size / (1024 * 1024):.1f} MB")
    if st.button("Processar liberação", type="primary", use_container_width=True):
        for chave in ("liberacao_estatisticas", "liberacao_excel", "liberacao_previa"):
            st.session_state.pop(chave, None)
        tarefa_id = obter_gerenciador_tarefas().iniciar(processar_liberacao, arquivo.getvalue())
        st.session_state["tarefa_liberacao"] = tarefa_id
        st.query_params["tarefa_liberacao"] = tarefa_id

tarefa_id = st.session_state.get("tarefa_liberacao") or st.query_params.get("tarefa_liberacao")
if tarefa_id and "liberacao_excel" not in st.session_state:
    try:
        estatisticas, excel = acompanhar_tarefa(tarefa_id)
        st.session_state["liberacao_estatisticas"] = estatisticas
        st.session_state["liberacao_excel"] = excel
        st.session_state["liberacao_previa"] = pd.read_excel(BytesIO(excel), sheet_name="BASE", nrows=100)
        nome_base = Path(arquivo.name).stem if arquivo is not None else "liberacao"
        st.session_state["liberacao_nome"] = f"{nome_base}_liberacao.xlsx"
        st.session_state.pop("tarefa_liberacao", None)
        st.query_params.pop("tarefa_liberacao", None)
    except Exception as erro:
        st.session_state.pop("tarefa_liberacao", None)
        st.query_params.pop("tarefa_liberacao", None)
        if str(erro) == "A tarefa não está mais disponível no servidor.":
            st.info("A tarefa anterior expirou após uma reinicialização. Envie o arquivo novamente.")
        else:
            st.error(f"Não foi possível gerar a base de liberação: {erro}")

if "liberacao_excel" in st.session_state:
    dados = st.session_state["liberacao_estatisticas"]
    st.success("Base de liberação gerada com sucesso.")
    st.markdown(
        f"""
        <div class="metric-grid">
          <div class="metric-card"><div class="metric-label">Registros lidos</div><div class="metric-value">{dados['registros']:,}</div></div>
          <div class="metric-card"><div class="metric-label">Duplicatas e revisões</div><div class="metric-value">{dados['base']:,}</div></div>
          <div class="metric-card"><div class="metric-label">Desmembramentos</div><div class="metric-value">{dados['desmembramentos']:,}</div></div>
        </div>
        """.replace(",", "."),
        unsafe_allow_html=True,
    )
    with st.container(border=True):
        st.subheader("Prévia da aba BASE")
        st.caption("Primeiros 100 registros. O arquivo para download contém todas as abas e registros.")
        st.dataframe(st.session_state["liberacao_previa"], use_container_width=True, hide_index=True)
    st.download_button(
        "⬇️ Baixar base de liberação",
        data=st.session_state["liberacao_excel"],
        file_name=st.session_state["liberacao_nome"],
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
        use_container_width=True,
    )
