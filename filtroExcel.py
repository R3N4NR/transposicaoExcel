import xlwings as xw
import tkinter as tk
from tkinter import filedialog

# Função que será chamada ao selecionar o arquivo
def processar_arquivo(arquivo):
    try:
        wb = xw.Book(arquivo)

        for sheet in wb.sheets:
            sheet.range('S2:S100').formula = '=IF(LEFT(A2, 4) = "Nota", B2, "")'
            sheet.range('T2:T100').formula = '=IF(ISERROR(SEARCH(" - ", A1)), "", IF(AND(ISNUMBER(VALUE(LEFT(A1, SEARCH(" - ", A1) - 1))), SEARCH(" - ", A1) > 0), A1, ""))'

            last_row = sheet.range('T' + str(sheet.cells.last_cell.row)).end('up').row
            for row in range(2, last_row + 1):
                cell_value = sheet.range(f'T{row}').value
                if cell_value:
                    move_to_row = max(2, row - 4)
                    sheet.range(f'T{move_to_row}').value = cell_value
                    sheet.range(f'T{row}').value = ''

            preencher_celulas_st(sheet)

        criar_tabela_dinamica(wb)
        wb.save()
        wb.close()
        resultado_label.config(text="Processamento concluído com sucesso!")
    except Exception as e:
        resultado_label.config(text=f"Erro: {str(e)}")

def preencher_celulas_st(sheet):
    ultima_linha = sheet.range('S' + str(sheet.cells.last_cell.row)).end('up').row
    
    for i in range(1, ultima_linha + 1):
        valor_atual = sheet.range(f'S{i}').value
        valor_t = sheet.range(f'T{i}').value
        
        if valor_atual and valor_t:
            j = i + 1
            while j <= ultima_linha and sheet.range(f'T{j}').value:
                if not sheet.range(f'S{j}').value:
                    sheet.range(f'S{j}').value = valor_atual
                j += 1

def criar_tabela_dinamica(wb):
    sheet = wb.sheets.add(name='Tabela Dinâmica')
    dados = []

    for sh in wb.sheets:
        ultima_linha = sh.range('S' + str(sh.cells.last_cell.row)).end('up').row
        for i in range(2, ultima_linha + 1):
            s_value = sh.range(f'S{i}').value
            t_value = sh.range(f'T{i}').value
            if s_value is not None and t_value is not None:
                dados.append([s_value, t_value])

    if dados:
        sheet.range('A1').value = ['Nota', 'Insumo']
        sheet.range('A2').value = dados
        
        last_row = sheet.range('A' + str(sheet.cells.last_cell.row)).end('up').row
        
        tabela_dinamica = sheet.api.PivotTableWizard(
            SourceData=sheet.range(f'A1:B{last_row}').api,
            TableDestination=sheet.range('D1').api,
            Function=-4100
        )

        tabela_dinamica.PivotFields('Nota').Orientation = 1  # xlRowField
        tabela_dinamica.PivotFields('Insumo').Orientation = 2  # xlDataField
        tabela_dinamica.PivotFields('Nota').Orientation = 3  # xlPageField (Filtro)
        tabela_dinamica.PivotFields('Insumo').Orientation = 4  # xlColumnField (Filtro)

        tabela_dinamica.RefreshTable()

# Criação da interface gráfica
root = tk.Tk()
root.title("Processador de Arquivo Excel")
root.geometry("400x200")

# Label para mostrar o resultado
resultado_label = tk.Label(root, text="")
resultado_label.pack(pady=20)

# Função para abrir o diálogo de seleção de arquivo
def selecionar_arquivo():
    arquivo = filedialog.askopenfilename(title="Selecione um arquivo Excel", filetypes=[("Excel Files", "*.xlsx")])
    if arquivo:
        processar_arquivo(arquivo)

# Botão para selecionar o arquivo
botao_selecionar = tk.Button(root, text="Selecionar arquivo .xlsx", command=selecionar_arquivo)
botao_selecionar.pack(pady=20)

root.mainloop()
