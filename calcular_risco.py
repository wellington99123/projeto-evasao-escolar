import pandas as pd
import os

ARQUIVO = 'alunos.xlsx'
def calcular_risco(linha) :
    risco = 0
    risco += min(linha['faltas_bimestre'] / 40, 1) * 35
    risco += max(0, (5 - linha['media_notas']) / 5) * 20
    risco += max(0, (3 - linha['renda_salarios']) / 3) * 12
    risco += (linha['reprovacoes'] / 4) * 8
    if linha['trabalha'] == 'Sim':
        risco += 25
    return round(risco)