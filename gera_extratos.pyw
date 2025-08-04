import pandas as pd
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.lib.units import cm
from reportlab.lib import colors
import tkinter as tk
from tkinter import filedialog


arquivo_pdf = "relatorio.pdf"
c = canvas.Canvas(arquivo_pdf, pagesize=A4)


def decimal_para_hhmm(decimal_horas):
    negativo = decimal_horas < 0
    decimal_horas = abs(decimal_horas)
    horas = int(decimal_horas)
    minutos = int(round((decimal_horas - horas) * 60))
    # Corrige minutos 60 para +1 hora
    if minutos == 60:
        horas += 1
        minutos = 0
    resultado = f"{horas:02d}:{minutos:02d}"
    return f"-{resultado}" if negativo else resultado

def calcular_saldo_acumulado(dados_func):
    dados_func = dados_func.reset_index(drop=True)
    saldo_acumulado = []
    saldo = 0.0
    for idx, row in dados_func.iterrows():
        acrescimo = row['hora_banco'] if not pd.isna(row['hora_banco']) else 0.0
        desconto = row['horas_descontadas'] if not pd.isna(row['horas_descontadas']) else 0.0
        if idx == 0:
            saldo = acrescimo - desconto  # saldo começa do zero no primeiro mês
        else:
            saldo = saldo + acrescimo - desconto
        saldo_acumulado.append(saldo)
    return saldo_acumulado

# Meses em ordem para ordenação
ordem_meses = ['janeiro', 'fevereiro', 'março', 'abril', 'maio', 'junho',
               'julho', 'agosto', 'setembro', 'outubro', 'novembro', 'dezembro']

# Caminho do arquivo Excel
# Abre uma janela para o usuário escolher o arquivo Excel
root = tk.Tk()
root.withdraw()  # esconde a janela principal do Tkinter
arquivo = filedialog.askopenfilename(
    title="Selecione o arquivo Excel",
    filetypes=(("Arquivos Excel", "*.xlsx;*.xls"), ("Todos os arquivos", "*.*"))
)
if not arquivo:
    print("Nenhum arquivo selecionado. Encerrando o programa.")
    exit()

print(arquivo)
#arquivo = r"C:\Users\Rodrigo\Desktop\RH\banco_horas\new\banco_horas.xlsx"

# Carregar as planilhas
empregados = pd.read_excel(arquivo, sheet_name='empregados')
acrescimos = pd.read_excel(arquivo, sheet_name='acrescimos')
descontos = pd.read_excel(arquivo, sheet_name='descontos')

# Corrigir nome da coluna descontos
descontos.rename(columns={'hora_descontada': 'horas_descontadas'}, inplace=True)

# Padronizar nomes de meses minúsculos
acrescimos['mes'] = acrescimos['mes'].astype(str).str.lower()
descontos['mes'] = descontos['mes'].astype(str).str.lower()

# Agrupar totais de acréscimos e descontos mês a mês por funcionário
acrescimos_agrup = acrescimos.groupby(['matricula', 'nome', 'mes'])['hora_banco'].sum().reset_index()
descontos_agrup = descontos.groupby(['matricula', 'nome', 'mes'])['horas_descontadas'].sum().reset_index()

# Mesclar em relatorio_mes para combinar acréscimos e descontos
relatorio_mes = pd.merge(acrescimos_agrup, descontos_agrup,
                        on=['matricula', 'nome', 'mes'], how='outer').fillna(0)

# Ordenar meses pela ordem definida
relatorio_mes['mes'] = pd.Categorical(relatorio_mes['mes'], categories=ordem_meses, ordered=True)
relatorio_mes = relatorio_mes.sort_values(['matricula', 'mes'])

def gerar_pdf_funcionario(c, matricula, nome):
    #c = canvas.Canvas(f"relatorio_{matricula}_{nome}.pdf", pagesize=A4)
    largura_pagina, altura_pagina = A4
    y = altura_pagina - 2 * cm

    # Títulos centralizados
    c.setFont("Courier-Bold", 14)
    titulo1 = "BANCO DE HORAS"
    c.drawCentredString(largura_pagina / 2, y, titulo1)
    y -= 1.0 * cm
    c.setFont("Courier", 12)
    titulo2 = "JULHO a AGOSTO de 2025"
    c.drawCentredString(largura_pagina / 2, y, titulo2)
    y -= 1.2 * cm   # Pequeno espaço após o título

    # Nome do funcionário
    c.setFont("Courier-Bold", 12)
    c.drawString(2 * cm, y, f"{matricula} - {nome}")
    y -= 1.0 * cm

    # Subtítulo
    c.setFont("Courier-Bold", 12)
    c.drawString(2 * cm, y, "Saldo de horas mês a mês")
    y -= 1.0 * cm

    # Tabela cabeçalho
    col_larguras = [4*cm, 4*cm, 4*cm, 4*cm]
    x_inicial = 2 * cm
    linha_altura = 0.7 * cm

    c.setFont("Courier-Bold", 11)
    headers = ["Mês", "Acréscimo", "Descontos", "Saldo"]
    x = x_inicial
    for i, header in enumerate(headers):
        c.drawString(x + 0.3 * cm, y, header)
        x += col_larguras[i]
    # Linha simples após cabeçalho
    c.setLineWidth(1)
    c.line(x_inicial, y - 5, x_inicial + sum(col_larguras) - 60, y - 5)
    c.setLineWidth(1)
    y -= linha_altura

    # Dados
    dados_func = relatorio_mes[relatorio_mes['matricula'] == matricula].copy().reset_index(drop=True)
    dados_func['saldo_acm'] = calcular_saldo_acumulado(dados_func)

    total_acrescimos = dados_func['hora_banco'].sum()
    total_descontos = dados_func['horas_descontadas'].sum()
    saldo_final = dados_func['saldo_acm'].iloc[-1] if not dados_func.empty else 0.0

    c.setFont("Courier", 10)
    for _, row in dados_func.iterrows():
        x = x_inicial
        dados_linha = [
            row['mes'].capitalize(),
            decimal_para_hhmm(row['hora_banco']),
            decimal_para_hhmm(row['horas_descontadas']),
            decimal_para_hhmm(row['saldo_acm']),
        ]
        for i, dado in enumerate(dados_linha):
            c.drawString(x + 0.3 * cm, y, dado)
            x += col_larguras[i]
        y -= linha_altura
        if y < 5 * cm:
            c.showPage()
            y = altura_pagina - 2 * cm

    # Linha dupla acima da linha de totais
    y_totais = y - 2
    c.setLineWidth(0.5)
    c.line(x_inicial, y_totais + 17, x_inicial + sum(col_larguras) - 60, y_totais + 17)
    c.setLineWidth(0.5)
    c.line(x_inicial, y_totais + 15, x_inicial + sum(col_larguras) - 60, y_totais + 15)
    c.setLineWidth(1)
    # Linha de totais
    c.setFont("Courier-Bold", 10)
    totais_linha = [
        "Totais",
        decimal_para_hhmm(total_acrescimos),
        decimal_para_hhmm(total_descontos),
        decimal_para_hhmm(saldo_final),
    ]
    x = x_inicial
    for i, dado in enumerate(totais_linha):
        c.drawString(x + 0.3 * cm, y, dado)
        x += col_larguras[i]
    y -= linha_altura + 0.5 * cm

    # Histórico detalhado de descontos
    c.setFont("Courier-Bold", 12)
    c.drawString(2 * cm, y, "Histórico de descontos")
    y -= 0.8 * cm

    c.setFont("Courier-Bold", 10)
    c.drawString(2.2 * cm, y, "data")
    c.drawString(6.5 * cm, y, "horas")
    # Linha após cabeçalho do histórico
    #c.setLineWidth(0.8)
    #c.line(2.2 * cm, y - 2, 10 * cm, y - 2)
    #c.setLineWidth(1)
    y -= 0.6 * cm
    c.setFont("Courier", 10)

    hist = descontos[descontos['matricula'] == matricula].copy()
    total_horas = 0.0

    if 'data' in hist.columns:
        hist['data'] = pd.to_datetime(hist['data'], errors='coerce')
        hist = hist.sort_values('data')
        hist['data_str'] = hist['data'].dt.strftime('%d/%m/%Y')
        hist['data_str'] = hist['data_str'].fillna('sem data')
        for _, row_hist in hist.iterrows():
            c.drawString(2.2 * cm, y, row_hist['data_str'])
            c.drawString(6.5 * cm, y, decimal_para_hhmm(row_hist['horas_descontadas']))
            total_horas += row_hist['horas_descontadas']
            y -= 0.5 * cm
            if y < 3 * cm:
                c.showPage()
                y = altura_pagina - 2 * cm
    else:
        hist['mes'] = hist['mes'].astype(str).str.lower()
        hist['mes'] = pd.Categorical(hist['mes'], categories=ordem_meses, ordered=True)
        hist_agrup = hist.groupby('mes')['horas_descontadas'].sum().reset_index()
        hist_agrup = hist_agrup.sort_values('mes')
        for _, row_hist in hist_agrup.iterrows():
            c.drawString(2.2 * cm, y, row_hist['mes'].capitalize())
            c.drawString(6.5 * cm, y, decimal_para_hhmm(row_hist['horas_descontadas']))
            total_horas += row_hist['horas_descontadas']
            y -= 0.5 * cm
            if y < 3 * cm:
                c.showPage()
                y = altura_pagina - 2 * cm

    c.setFont("Courier-Bold", 11)
    c.drawString(2.2 * cm, y-1, "Total")
    c.drawString(6.5 * cm, y-1, decimal_para_hhmm(total_horas))
    c.showPage()
    

# Rodar para todos os empregados
for _, row in empregados.iterrows():
    gerar_pdf_funcionario(c, row['matricula'], row['nome'])

c.save()
