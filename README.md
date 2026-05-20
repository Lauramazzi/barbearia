# Barbearia Heloisa Mazzi

Dashboard em Streamlit para gestão financeira mensal da Barbearia Heloisa Mazzi.

O app consolida relatórios exportados do AppBarber com despesas manuais, parcelas,
valor guardado em conta, saldo de caixa, metas financeiras e metas de produção.

## Como usar

1. Se já tiver usado antes, envie a base histórica baixada no mês anterior.
2. Exporte do AppBarber:
   - `total recebimento periodo.xlsx`
   - `ranking serviços.xlsx`
   - `ranking produtos.xlsx`
3. Escolha o mês analisado.
4. Lance despesas na aba `Despesas`, classificando centro, categoria, tipo e parcelas.
5. Ajuste contas recorrentes na aba `Contas e parcelas`.
6. Cadastre metas financeiras e de produção na aba `Metas`.
7. Confira caixa, valor guardado, saldo e evolução.
8. Baixe a base histórica atualizada e use esse arquivo no próximo mês.

## Abas principais

- `Gestão do mês`: resumo financeiro do mês escolhido.
- `Despesas`: lançamentos mensais, despesas pessoais, barbearia e parcelas.
- `Contas e parcelas`: cadastro de contas recorrentes.
- `Metas`: metas de faturamento, despesas, caixa, valor guardado e produção.
- `Histórico e evolução`: acompanhamento mês a mês.
- `Rankings`: serviços e produtos exportados do AppBarber.
- `Exportar base`: arquivo para manter o histórico.

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
