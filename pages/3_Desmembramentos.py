from io import BytesIO
from pathlib import Path

import streamlit as st

from desmembramentos import processar_desmembramentos
from tarefas_background import acompanhar_tarefa, obter_gerenciador_tarefas


st.markdown(
    """
    <style>
    .block-container {max-width: 1180px; padding-top: 2rem; padding-bottom: 4rem;}
    .hero-desm {
        padding: 2rem 2.2rem; border-radius: 22px; margin-bottom: 1.4rem;
        background: linear-gradient(135deg, #25233f 0%, #47366f 55%, #7657b5 100%);
        box-shadow: 0 18px 45px rgba(0,0,0,.18); color: white;
    }
    .hero-desm h1 {font-size: 2.25rem; margin: 0 0 .5rem; color: white;}
    .hero-desm p {font-size: 1.02rem; margin: 0; opacity: .88; max-width: 760px;}
    div[data-testid="stButton"] button, div[data-testid="stDownloadButton"] button {
        min-height: 3rem; border-radius: 12px; font-weight: 700;
    }
    @media (max-width: 720px) {.hero-desm{padding:1.5rem}.hero-desm h1{font-size:1.8rem}}
    </style>
    <div class="hero-desm">
      <h1>Desmembramento</h1>
      <p>Valide imagens, coordenadas e regras da base e gere o arquivo consolidado em CSV.</p>
    </div>
    """,
    unsafe_allow_html=True,
)


def executar(arquivo_bytes: bytes, progresso):
    resultado = processar_desmembramentos(BytesIO(arquivo_bytes), progresso=progresso)
    csv = resultado.to_csv(index=False, sep=";").encode("utf-8-sig")
    return resultado, csv


with st.container(border=True):
    st.subheader("Enviar base para processamento")
    st.caption("Formatos aceitos: Excel (.xlsx ou .xls) e CSV (.csv)")
    arquivo = st.file_uploader(
        "Selecione a base",
        type=["xlsx", "xls", "csv"],
        label_visibility="collapsed",
        key="arquivo_desmembramentos",
    )

if arquivo is None:
    st.info("📂 Selecione uma base para iniciar o processamento.")
else:
    st.caption(f"✓ {arquivo.name}  •  {arquivo.size / (1024 * 1024):.1f} MB")
    if st.button("Processar base", type="primary", use_container_width=True):
        st.session_state.pop("desmembramentos_resultado", None)
        gerenciador = obter_gerenciador_tarefas()
        tarefa_id = gerenciador.iniciar(executar, arquivo.getvalue())
        st.session_state["tarefa_desmembramento"] = tarefa_id
        st.query_params["tarefa_desmembramento"] = tarefa_id

tarefa_desmembramento = (
    st.session_state.get("tarefa_desmembramento")
    or st.query_params.get("tarefa_desmembramento")
)
if tarefa_desmembramento and "desmembramentos_resultado" not in st.session_state:
    try:
        resultado, csv_saida = acompanhar_tarefa(tarefa_desmembramento)
        st.session_state["desmembramentos_resultado"] = resultado
        st.session_state["desmembramentos_csv"] = csv_saida
        st.session_state.pop("tarefa_desmembramento", None)
        st.query_params.pop("tarefa_desmembramento", None)
        nome_base = Path(arquivo.name).stem if arquivo is not None else "base"
        st.session_state["desmembramentos_nome"] = f"{nome_base}_desmembramento.csv"
    except Exception as erro:
        st.session_state.pop("tarefa_desmembramento", None)
        st.query_params.pop("tarefa_desmembramento", None)
        if str(erro) == "A tarefa não está mais disponível no servidor.":
            st.info("A tarefa anterior expirou após uma reinicialização do servidor. Envie o arquivo novamente.")
        else:
            st.error(f"Não foi possível processar o arquivo: {erro}")

if "desmembramentos_resultado" in st.session_state:
    resultado = st.session_state["desmembramentos_resultado"]
    st.success(f"{len(resultado):,} registros processados.".replace(",", "."))
    with st.container(border=True):
        st.subheader("Prévia do resultado")
        st.caption("Primeiros 100 registros. O CSV contém a base completa.")
        st.dataframe(resultado.head(100), use_container_width=True, hide_index=True)

    st.download_button(
        "⬇️ Baixar CSV processado",
        data=st.session_state["desmembramentos_csv"],
        file_name=st.session_state["desmembramentos_nome"],
        mime="text/csv",
        type="primary",
        use_container_width=True,
    )
