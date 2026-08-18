# Ferramentas de Censo

Aplicativo Streamlit com duas ferramentas:

- tratamento e auditoria de bases de censo;
- comparação da produtividade informada pelos agentes com a produtividade do sistema.

## Executar localmente

```powershell
python -m pip install -r requirements.txt
python -m streamlit run app.py
```

O tratamento gera as abas `Base_Consolidada`, `Resumo_Ponto_Focal`,
`Dicionario_e_Regras` e `Resumo_Por_Agente`. A ferramenta de produtividade
oferece relatórios em Excel e PDF.

Para publicar para toda a equipe, hospede esta pasta em um servidor interno ou
em um serviço compatível com Streamlit. Como o processamento acessa links de
imagens presentes na planilha, o servidor precisa conseguir acessar esses links.
