# Painel AppBarber

Dashboard em Streamlit para analisar relatórios exportados do AppBarber.

## Como usar

1. Exporte do AppBarber:
   - `total recebimento periodo.xlsx`
   - `ranking serviços.xlsx`
   - `ranking produtos.xlsx`
2. Abra o app.
3. Envie os arquivos na barra lateral.
4. Informe despesas manuais e valor guardado na conta.
5. Veja o cenário atual e baixe o Excel consolidado.

## Rodar localmente

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Deploy no Streamlit Cloud

No Streamlit Community Cloud, escolha este repositório e use:

```text
app.py
```

como arquivo principal.
