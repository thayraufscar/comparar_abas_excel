# Comparar Registros entre Duas Abas de uma Planilha Excel

Este projeto contém um script Python para comparar os registros de duas abas de uma mesma planilha Excel.
O script identifica quais registros existem em uma aba mas não na outra.

## Como funciona
1. Lê a planilha `entrada.xlsx` (substitua pelo nome do seu arquivo).
2. Normaliza os textos: remove espaços extras e coloca em minúsculas.
3. Compara os registros entre duas abas.
4. Gera um novo arquivo `comparacao_abas.xlsx` com duas abas:
   - `Somente_Sheet1` → registros exclusivos da primeira aba
   - `Somente_Sheet2` → registros exclusivos da segunda aba

## Requisitos
- Python 3.x
- Pandas
- openpyxl

Para instalar dependências:
```bash
pip install pandas openpyxl
```

## Como usar
1. Coloque sua planilha no mesmo diretório do script e renomeie para `entrada.xlsx`.
2. Abra o arquivo `comparar_abas.py` e ajuste:
   - `aba1` → Nome da primeira aba a comparar.
   - `aba2` → Nome da segunda aba a comparar.
   - `coluna` → Nome da coluna com os textos a verificar.
3. Execute:
```bash
python comparar_abas.py
```
4. O resultado estará em `comparacao_abas.xlsx`.
