
# Site de Controle de Manutenção

Este app **Streamlit** lê o arquivo `data/controle.xlsx` (cópia do seu `Controle_de_Pedidos_Manut.xlsm`) e oferece:

- **Dashboard** com filtros por CD, ano, mês, grupo, subgrupo e status.
- **KPIs** e gráficos de despesas por mês e por grupo.
- **Tabela detalhada** das requisições filtradas.
- **BGT x REQ**: compara o orçamento mensal (Planilha *Base_Budget*) com o executado das requisições (Planilha *Base_Requisicoes*), incluindo o **Disponível** (BGT - REQ).
- **Estornos Abertos**: exibe a planilha *Base_Estornos_Abertos* quando presente.
- **Exportação**: baixa os dados filtrados em Excel.

## Como rodar

1. Crie um ambiente e instale as dependências:

```bash
pip install -r requirements.txt
```

2. Inicie o app:

```bash
streamlit run app.py
```

3. O app lerá `data/controle.xlsx`. Para atualizar os dados, substitua o arquivo nessa pasta por uma nova versão do seu controle.

## Observações

- As colunas são detectadas automaticamente e os meses do orçamento são mapeados para um formato YYYY-MM.
- Campos de data como **MÊS COMPETÊNCIA** são convertidos para datas/Período quando possível.
- Se quiser permitir **edição** e **gravação** de novos lançamentos, posso evoluir para gravar em um arquivo `data/lancamentos.csv` e consolidar com o Excel original.

