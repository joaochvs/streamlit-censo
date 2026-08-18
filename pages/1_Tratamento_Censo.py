from io import BytesIO
from pathlib import Path

import pandas as pd
import streamlit as st

from tratamento_censo import processar_censo




def gerar_excel(relatorios: dict[str, pd.DataFrame]) -> bytes:
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        for nome_aba, dados in relatorios.items():
            dados.to_excel(writer, sheet_name=nome_aba[:31], index=False)
            planilha = writer.sheets[nome_aba[:31]]
            planilha.freeze_panes = "A2"
            planilha.auto_filter.ref = planilha.dimensions
    return buffer.getvalue()


st.title("Tratamento de Censo")
st.write(
    "Envie a base em Excel. O sistema executa as validações, apresenta uma "
    "prévia e gera um arquivo tratado com todas as abas de relatório."
)

arquivo = st.file_uploader("Selecione o arquivo", type=["xlsx", "xls"])

if arquivo is None:
    st.info("Aguardando o envio de uma planilha Excel.")
else:
    tamanho_mb = arquivo.size / (1024 * 1024)
    st.caption(f"{arquivo.name} • {tamanho_mb:.1f} MB")

    if st.button("Processar arquivo", type="primary", use_container_width=True):
        try:
            with st.spinner("Processando a base. O download das imagens pode levar alguns minutos..."):
                arquivo.seek(0)
                relatorios = processar_censo(arquivo)
                excel_tratado = gerar_excel(relatorios)
                st.session_state["relatorios"] = relatorios
                st.session_state["excel_tratado"] = excel_tratado
                st.session_state["nome_saida"] = f"{Path(arquivo.name).stem}_tratado.xlsx"
        except Exception as erro:
            st.session_state.pop("relatorios", None)
            st.error(f"Não foi possível processar o arquivo: {erro}")

if "relatorios" in st.session_state:
    relatorios = st.session_state["relatorios"]
    base = relatorios["Base_Consolidada"]
    resumo = relatorios["Resumo_Ponto_Focal"]

    st.success("Tratamento concluído.")
    coluna_1, coluna_2 = st.columns(2)
    coluna_1.metric("Registros processados", f"{len(base):,}".replace(",", "."))
    coluna_2.metric("Registros para revisão", f"{len(resumo):,}".replace(",", "."))

    st.subheader("Prévia da base consolidada")
    st.dataframe(base.head(100), use_container_width=True, hide_index=True)
    st.caption("A prévia mostra os primeiros 100 registros. O download contém a base completa.")

    st.download_button(
        "Baixar arquivo tratado",
        data=st.session_state["excel_tratado"],
        file_name=st.session_state["nome_saida"],
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
        use_container_width=True,
    )
