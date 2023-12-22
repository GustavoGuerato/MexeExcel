# Automação do Excel com Openpyxl em Python

## Código 1: Modificando uma Planilha Existente

```python
import openpyxl
from random import uniform
```
# Carregar uma planilha existente
pedidos = openpyxl.load_workbook('pedidos.xlsx')
nomes_planilhas = pedidos.sheetnames
planilha1 = pedidos['Página1']

# Modificar a Planilha1
for linha in range(5, 16):
    numero_pedido = linha - 1
    planilha1.cell(linha, 1).value = numero_pedido
    planilha1.cell(linha, 2).value = 1200 + linha
    preco = round(uniform(10, 120), 2)
    planilha1.cell(linha, 3).value = preco

# Imprimir valores não nulos na Planilha1
for linha in planilha1:
    if linha[0].value is not None:
        print(linha[0].value, end=' ')
    if linha[1].value is not None:
        print(linha[1].value, end=' ')
    if linha[2].value is not None:
        print(linha[2].value)

# Salvar a planilha modificada
pedidos.save('nova_planilha.xlsx')
