# Relatório Automatizado de Vendas por Loja

## Descrição
Este script realiza a leitura de uma base de dados de vendas em formato Excel, calcula métricas agregadas por loja e envia automaticamente um relatório por e-mail utilizando o Microsoft Outlook.

O processamento dos dados é feito com a biblioteca **pandas**, que permite manipulação eficiente de dados tabulares. Após o cálculo das métricas, os resultados são convertidos para tabelas HTML e inseridos no corpo de um e-mail enviado automaticamente via automação do Outlook utilizando **win32com**.

A base de dados utilizada neste projeto é uma planilha de vendas com múltiplas lojas fictícias (gerada apenas para fins de teste e prática de análise de dados).

---

## Dependências
O script utiliza as seguintes bibliotecas:

* `pandas`
* `pywin32` (`win32com.client`)

Instalação:
```bash
pip install pandas pywin32
```

Para que o envio de e-mails funcione corretamente, é necessário que o **Microsoft Outlook esteja instalado e configurado no sistema**, pois o script utiliza a interface COM do Windows para criar e enviar o e-mail.

---

## Base de Dados
O script utiliza o arquivo:

```
Vendas.xlsx
```

A planilha contém registros individuais de vendas realizados por diferentes lojas fictícias (como *Loja Centro*, *Mega Store Paulista*, *Shopping Aurora*, entre outras).

As principais colunas utilizadas no processamento são:

* `ID Loja` — identificador da loja responsável pela venda
* `Valor Final` — valor monetário total da venda
* `Quantidade` — número de produtos vendidos na transação

Cada linha da planilha representa uma venda individual.

---

## Etapas do Processamento
### Leitura da planilha
A base de dados é carregada em memória como um **DataFrame do pandas**:

```python
tabela_vendas = pd.read_excel('Vendas.xlsx')
```

O comando

```python
pd.set_option('display.max_columns', None)
```

é utilizado apenas para garantir que todas as colunas da tabela sejam exibidas ao imprimir o DataFrame no terminal.

---

### Cálculo do faturamento por loja

O faturamento total de cada loja é calculado agrupando os registros pela coluna `ID Loja` e somando os valores da coluna `Valor Final`:

```python
faturamento = tabela_vendas[['ID Loja', 'Valor Final']].groupby('ID Loja').sum()
```

O resultado é uma tabela contendo o faturamento total gerado por cada loja.

---

### Cálculo da quantidade de produtos vendidos

A quantidade total de itens vendidos é obtida agrupando os registros por loja e somando a coluna `Quantidade`:

```python
quantidade = tabela_vendas[['ID Loja', 'Quantidade']].groupby('ID Loja').sum()
```

Essa etapa permite identificar o volume total de produtos vendidos por cada unidade.

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
