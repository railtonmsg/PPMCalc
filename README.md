PPMCalc

Script simples para calcular massa teórica e erro em ppm a partir de dados de LC-MS exportados para Excel.
Ajusta a fórmula química conforme o tipo de aduto informado e gera um novo arquivo .xlsx com os resultados.

Requisitos

- Python 3 (3.8 ou superior)
- Bibliotecas Python:
    pip install pandas pyfiglet openpyxl

Como usar

1. Coloque o arquivo ppm.py na mesma pasta das planilhas .xlsx ou .xltx.
2. No terminal, entre na pasta e execute:
       python ppm.py
3. Escolha a opção 1 para ver as planilhas disponíveis.
4. Digite o número da planilha que quer processar.
5. O script cria um arquivo resultado_<nome_da_planilha>.xlsx.

Adutos suportados

O script ajusta a fórmula para os principais adutos, por exemplo:

- Negativos: [M-H]-, [M-2H]-, [M+Cl]-, [M+FA-H]- etc.
- Positivos: [M+H]+, [M+Na]+, [M+K]+, [M+NH4]+, [2M+Na]+ etc.

(A lista completa está no código, função remove_quantidade_hidrogenios.)

Observações

- Se precisar adicionar novos elementos (ex.: Br, I, F) ou adutos, edite o dicionário no código.
- O cálculo usa massas monoisotópicas.

Créditos

Autores: Railton Marques, Fernando Zanchi — LABIOQUIM/CEBIO‑FIOCRUZ/RO (2024).

