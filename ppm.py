import pandas as pd
import numpy as np 
import os
import pyfiglet


print(pyfiglet.figlet_format("PPMCalc"))
print("Autores: Railton Marques, Fernando Zanchi - LABIOQUIM/CEBIO-FIOCRUZ/RO, 2024.")

arquivos = os.listdir('.')

planilhas = [arquivo for arquivo in arquivos if arquivo.endswith(('.xlsx', '.xltx'))]

while True:
    print("\nOpções:")
    print("1 - Ver planilhas disponíveis")
    print("Q - Sair")
    opcao = input("Escolha uma opção: ").strip().upper()

    if opcao == "1":
        if not planilhas:
            print("\nNenhuma planilha .xlsx encontrada no diretório.")
            continue 

        print("\nPlanilhas disponíveis:")
        for i, planilha in enumerate(planilhas):
            print(f"{i + 1}: {planilha}")

        escolha = input("Escolha o número da planilha que deseja carregar (ou Q para voltar): ").strip().upper()

        if escolha == "Q":
            continue 
        elif escolha.isdigit() and 1 <= int(escolha) <= len(planilhas):
            nome_planilha = planilhas[int(escolha) - 1]
            try:
                # Carregando a planilha escolhida
                data_n = pd.read_excel(nome_planilha, header=3)
                print(f"\nPlanilha '{nome_planilha}' carregada com sucesso.")
 
                def remove_quantidade_hidrogenios(formula, adduct_type):
                    import re

                    # Mapeamento de adutos
                    adduct_mapping = {

                        "[M-2H]-": {"H": -2},
                        "[M-2H]2-": {"H": -4},
                        "[M-H]-": {"H": -1},
                        "[M-H]1-": {"H": -1},
                        "[M-H2O-H]-": {"H": -3,"O": -1},
                        "[M+Cl]-": {},
                        "[M+FA-H]-": {},
                        "[M+HCOO]-": {},
                        "[M-2H2O+H]+": {"H": -3,"O": -2},
                        "[M-3H2O+H]+": {"H": -5,"O": -3},
                        "[2M+Na]+": {"Na": 1},
                        "[M-2H2O+H]+": {"H": -3,"O": -2},
                        "[M-3H2O+H]+": {"H": -5,"O": -3},
                        "[M-H+2Na]+": {"H": -1,"Na": +2},
                        "[M-H2O+H]+": {"H": -1,"O": -1},
                        "[M]+": {},
                        "[M+2H]2+": {"H": +4},
                        "[M+H-H2O]-": {"H": -1,"O": -1},
                        "[M+H-H2O]+": {"H": -1,"O": -1},
                        "[M+H]+": {"H": +1},
                        "[M+K]+": {"K": +1},
                        "[M+Na]+": {"Na": +1},
                        "[M+NH4]+": {"N": 1, "H": 4},
                        "[M+Cl]-": {"Cl": +1},
                        "[M+FA-H]-": {"H": +1,"O": +2, "C": +1},
                        "[M+HCOO]-": {"H": +1, "C": +1, "O": +2},

                    }
                    
                    # Verifica se o aduto está no mapeamento
                    if adduct_type in adduct_mapping:
                        alteracoes = adduct_mapping[adduct_type]
                        for elemento, quantidade in alteracoes.items():
                            match = re.search(rf'{elemento}(\d*)', formula)
                            if match:
                                quantidade_atual = match.group(1)
                                if quantidade_atual == '':
                                    quantidade_atual = '1'
                                quantidade_atual = int(quantidade_atual)
                                quantidade_nova = quantidade_atual + quantidade
                                if quantidade_nova > 0:
                                    formula = re.sub(rf'{elemento}(\d*)', f'{elemento}{quantidade_nova}', formula)
                                else:
                                    formula = re.sub(rf'{elemento}(\d*)', '', formula)
                            else:
                                if quantidade > 0:
                                    formula += f'{elemento}{quantidade}'

                    return formula

                # Função para calcular a massa molecular
                def calcular_massa_molecular(formula, adduct_type): #espera o argumento formula e o aduto


                    massas_atomicas = {
                    "C": 12.000000,
                    "H": 1.007825,
                    "O": 15.994915,
                    "S": 31.972072,
                    "P": 30.973762,
                    "Na": 22.989769,
                    "N": 14.003074,
                    "Cl": 35.452138,
                    "K": 39.0983,
                    }

                    formula = remove_quantidade_hidrogenios(formula, adduct_type) # atualiza a formula
                            
                    massa_molecular = 0.0
                    elemento_atual = ""
                    quantidade_atual = ""
                    for i, elemento in enumerate(formula):
                        if elemento.isalpha():
                            if elemento_atual:
                                if elemento_atual in massas_atomicas:
                                    massa_molecular += massas_atomicas[elemento_atual] * (int(quantidade_atual) if quantidade_atual else 1)

                            
                            if i < len(formula) - 1 and formula[i+1].isalpha():
                                
                                elemento_atual = elemento + formula[i+1]
                                quantidade_atual = ""
                                continue
                            else:
                        
                                elemento_atual = elemento
                                quantidade_atual = ""
                        
                        
                        elif elemento.isdigit():
                            quantidade_atual += elemento


                    if elemento_atual in massas_atomicas:
                        massa_molecular += massas_atomicas[elemento_atual] * (int(quantidade_atual) if quantidade_atual else 1)
                    return massa_molecular

                #Calculo do erro ppm
                def calcular_erro_ppm(valor_teorico, valor_expresso):

                    erro_ppm = ((valor_expresso - valor_teorico) / valor_teorico) * 10**6
                    return erro_ppm

                resultados = []

                # Colunas average e Stdev
                average_n = data_n.filter(like="Average")
                stdev_n = data_n.filter(like="Stdev")

                for i in average_n.columns:
                    data_n.loc[0, i] = f"Average_{data_n.loc[0, i]}"

                for i in stdev_n.columns:
                    data_n.loc[0, i] = f"Stdev_{data_n.loc[0, i]}"

                data_n.columns = data_n.loc[0]
                data_n = data_n.drop(index=0).reset_index(drop=True)

                average = data_n.filter(like="Average_")
                stdev = data_n.filter(like="Stdevs_")

                for index, row in data_n.iterrows():

                    averages = {col: row[col] for col in average}
                    stdevs = {col: row[col] for col in stdev}

                    alignment_id = row.loc["Alignment ID"]
                    Average_Rt = row.loc["Average Rt(min)"]
                    average_mz = row.loc["Average Mz"]
                    metabolite_name = row.loc["Metabolite name"]
                    adduct_type = row.loc["Adduct type"]
                    formula = row.loc["Formula"]
                    ontology = row.loc["Ontology"]
                    smiles = row.loc["SMILES"]
                    ms_ms_spectrum = row.loc["MS/MS spectrum"]
                    
                    if pd.isna(formula):
                        continue

                    if metabolite_name == "Unknown":  # Verifica se o nome do metabólico é NaN, passa para a proxima linha
                        continue

                    teorico = calcular_massa_molecular(formula, adduct_type)
                    erro_ppm = calcular_erro_ppm(teorico, average_mz)

                    # Adiciona os resultados à lista
                    resultados.append({
                    'ID de alinhamento': alignment_id,
                    'Rt Médio (min)': Average_Rt,
                    'Nome do metabólico': metabolite_name,
                    'Tipo de aduto': adduct_type,
                    'Fórmula': formula,
                    'Ontologia': ontology,
                    'SMILES': smiles,
                    'Massa Teórica': teorico,
                    'Massa Expressa': average_mz,
                    'Erro PPM': erro_ppm,
                    'Espectro MS/MS': ms_ms_spectrum,
                    **averages,
                    **stdevs,
                    })


                df_resultados = pd.DataFrame(resultados)
                    
                nome_resultado = f"resultado_{nome_planilha}"
                df_resultados.to_excel(f'{nome_resultado}.xlsx', index=False)

                    
                print(f"Resultados salvos com sucesso em '{nome_resultado}.xlsx'")

            except Exception as e:
                print(f"\nErro ao carregar a planilha: {e}")
        else:
            print("Entrada inválida. Por favor, tente novamente.")

    elif opcao == "Q":
        exit() 

    else:
        print("Opção inválida. Por favor, escolha uma opção válida.")





