import streamlit as st


st.set_page_config(
    page_title="Ferramentas de Censo",
    page_icon="📊",
    layout="wide",
)

pagina = st.navigation(
    {
        "Ferramentas": [
            st.Page(
                "pages/1_Tratamento_Censo.py",
                title="Tratamento de Censo",
                icon="📊",
                default=True,
            ),
            st.Page(
                "pages/2_Produtividade_Agentes.py",
                title="Produtividade dos Agentes",
                icon="📈",
            ),
        ]
    }
)
pagina.run()
