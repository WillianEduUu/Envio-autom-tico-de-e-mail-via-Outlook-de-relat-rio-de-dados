# Relatório de Vendas por Loja (Python)

## Descrição

Este script realiza a leitura de uma base de dados de vendas em formato Excel, calcula métricas agregadas por loja e envia automaticamente um relatório por e-mail utilizando o Microsoft Outlook.

O processamento é feito com a biblioteca **pandas**, que permite manipulação eficiente de dados tabulares. Após a análise, os resultados são convertidos para tabelas HTML e inseridos no corpo de um e-mail enviado via automação do Outlook utilizando **win32com**.

O objetivo do script é automatizar a geração e distribuição de relatórios operacionais de vendas.

---

## Dependências

O código utiliza as seguintes bibliotecas:

* `pandas`
* `pywin32` (win32com.client)

Instalação das dependências:

```bash
pip install pandas pywin32
```

O script depende também de uma instalação local do **Microsoft Outlook**, pois o envio do e-mail é feito através da interface COM do Windows.

---

## Fonte de Dados

O script lê o arquivo:

```
Vendas_novas_lojas.xlsx
```

Este arquivo contém os registros de vendas individuais utilizados para teste do script. A planilha possui dados simulados de vendas distribuídos entre diferentes lojas fictícias.

Entre as colunas utilizadas na análise estão:

* `ID Loja` – identificador da loja responsável pela venda
* `Valor Final` – valor monetário total da venda
* `Quantidade` – número de produtos vendidos na transação

Os dados são carregados em um **DataFrame do pandas** através da função:

```python
pd.read_excel()
```

---

## Etapas do Processamento

### Leitura da base de dados

A planilha Excel é carregada em memória como um DataFrame:

```python
tabela_vendas = pd.read_excel('Vendas_novas_lojas.xlsx')
```

Em seguida, a configuração

```python
pd.set_option('display.max_columns', None)
```

permite visualizar todas as colunas da tabela ao imprimir o DataFrame no terminal.

---

### Cálculo do faturamento por loja

O faturamento é calculado agrupando os registros pela coluna `ID Loja` e somando os valores da coluna `Valor Final`:

```python
faturamento = tabela_vendas[['ID Loja', 'Valor Final']].groupby('ID Loja').sum()
```

O resultado é uma tabela contendo o faturamento total de cada loja.

---

### Quantidade total de produtos vendidos

A quantidade de itens vendidos é calculada de forma semelhante, agrupando pela loja e somando a coluna `Quantidade`:

```python
quantidade = tabela_vendas[['ID Loja', 'Quantidade']].groupby('ID Loja').sum()
```

Essa tabela representa o volume total de produtos vendidos por cada loja.

---

### Cálculo do ticket médio

O ticket médio representa o valor médio por produto vendido em cada loja. Ele é calculado dividindo o faturamento total pela quantidade total de produtos vendidos:

```python
ticket_medio = (faturamento['Valor Final'] / quantidade['Quantidade']).to_frame()
```

Em seguida, a coluna é renomeada para facilitar a leitura no relatório:

```python
ticket_medio = ticket_medio.rename(columns={0: 'Ticket Médio'})
```

---

## Geração e envio do relatório por e-mail

Após o cálculo das métricas, o script cria um e-mail automaticamente utilizando o Outlook:

```python
outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
```

O destinatário, assunto e corpo do e-mail são definidos no código. O corpo da mensagem é construído em **HTML**, permitindo incluir as tabelas de resultados diretamente no e-mail.

As tabelas geradas pelo pandas são convertidas para HTML utilizando:

```python
DataFrame.to_html()
```

No caso do faturamento e do ticket médio, é aplicado um formatador para exibir os valores no formato monetário:

```python
formatters={'Valor Final': 'R${:,.2f}'.format}
```

Por fim, o e-mail é enviado automaticamente:

```python
mail.Send()
```

---

## Saída do Script

Ao executar o script, ocorrem duas saídas principais:

1. Impressão no terminal das tabelas calculadas (faturamento, quantidade e ticket médio).
2. Envio automático de um e-mail contendo o relatório de vendas formatado.

O e-mail enviado contém três tabelas:

* Faturamento total por loja
* Quantidade total de produtos vendidos por loja
* Ticket médio por loja

Essas informações permitem uma visualização rápida do desempenho de cada unidade.
