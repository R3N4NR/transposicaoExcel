import xlwings as xw
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import logging
from threading import Thread
import os

class TextHandler(logging.Handler):
    def __init__(self, text_widget):
        super().__init__()
        self.text_widget = text_widget

    def emit(self, record):
        log_entry = self.format(record)
        self.text_widget.configure(state="normal")
        self.text_widget.insert(tk.END, log_entry + '\n')
        self.text_widget.configure(state="disabled")
        self.text_widget.see(tk.END)



# Define o caminho do arquivo de log na raiz do usuário
log_file_path = os.path.join(os.path.expanduser("~"), "filtroExcel.log")
def setup_logger(text_widget):
    logger = logging.getLogger()
    logger.setLevel(logging.INFO)
    formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')

    # Arquivo de log
    file_handler = logging.FileHandler(log_file_path)
    file_handler.setFormatter(formatter)
    logger.addHandler(file_handler)

    # Text handler para o Text widget
    text_handler = TextHandler(text_widget)
    text_handler.setFormatter(formatter)
    logger.addHandler(text_handler)

def processar_arquivo(arquivo):
    logging.info(f"Iniciando o processamento do arquivo: {arquivo}")
    app = xw.App(visible=False)
    try:
        wb = app.books.open(arquivo)
        app.screen_updating = False
        app.calculation = 'manual'
        logging.info("Configurações de atualização e cálculo desativadas.")

        sheet_consolidado = None
        for sheet in wb.sheets:
            if sheet.name == 'Filtro':
                sheet_consolidado = sheet
                sheet_consolidado.clear()
                logging.info("Planilha 'Filtro' encontrada e limpa.")
                break
        if not sheet_consolidado:
            sheet_consolidado = wb.sheets.add(name='Filtro')
            logging.info("Planilha 'Filtro' criada.")

        for sheet in wb.sheets:
            if sheet.name != 'Filtro':
                logging.info(f"Processando planilha: {sheet.name}")
                desmesclar_e_mover(sheet)
                sheet.range('S2:S100').formula = '=IF(LEFT(A2, 4) = "Nota", B2, "")'
                sheet.range('T2:T100').formula = '=IF(ISERROR(SEARCH(" - ", A1)), "", IF(AND(ISNUMBER(VALUE(LEFT(A1, SEARCH(" - ", A1) - 1))), SEARCH(" - ", A1) > 0), A1, ""))'
                sheet.range('U2:U100').formula = '=IF(LEFT(E2, 10) = "Fornecedor", H2, "")'
                preencher_celulas_st(sheet)
        
        consolidar_dados(wb, sheet_consolidado)
        apagar_stu(wb)
        apagar_planilhas(wb)
        wb.save()
        wb.close()
        
        logging.info("Processamento concluído e arquivo salvo.")
        
        messagebox.showinfo("Concluído", "O processamento foi concluído com sucesso!")
        
    except Exception as e:
        logging.error(f"Erro durante o processamento: {str(e)}")
    finally:
        app.screen_updating = True
        app.calculation = 'automatic'
        app.quit()
        logging.info("Configurações do Excel restauradas e aplicativo fechado.")

def desmesclar_e_mover(sheet):
    logging.info(f"Desmesclando células na planilha {sheet.name}")
    ultima_linha = sheet.cells(sheet.cells.last_cell.row, 1).end('up').row
    for i in range(2, ultima_linha + 1):
        for col in 'ABCDEFGHIJKLMNO':
            cell = sheet.range(f'{col}{i}')
            if cell.merge_cells:
                cell.unmerge()
                logging.info(f"Célula {col}{i} desmesclada.")
        if sheet.range(f'D{i}').value == "Fornecedor":
            for col in reversed(range(4, 11)):
                origem = sheet.cells(i, col)
                destino = sheet.cells(i, col + 1)
                destino.value = origem.value
                origem.clear_contents()

def preencher_celulas_st(sheet):
    logging.info("Preenchendo células ST na planilha.")
    ultima_linha = sheet.range('S' + str(sheet.cells.last_cell.row)).end('up').row
    nota_atual = fornecedor_atual = None
    for i in range(2, ultima_linha + 1):
        if sheet.range(f'S{i}').value:
            nota_atual = sheet.range(f'S{i}').value
        if sheet.range(f'U{i}').value:
            fornecedor_atual = sheet.range(f'U{i}').value
        if sheet.range(f'T{i}').value:
            sheet.range(f'S{i}').value = nota_atual
            sheet.range(f'U{i}').value = fornecedor_atual
            
def aplicar_filtro(sheet):
    logging.info("Aplicando filtro nas colunas A a I.")
    sheet.range("A:C").api.AutoFilter(1)  


def consolidar_dados(wb, sheet_consolidado):
    logging.info("Consolidando dados na planilha 'Filtro'.")
    sheet_consolidado.range('A1').value = ['Nota', 'Fornecedor', 'Insumo']
    dados_consolidados = []
    for sheet in wb.sheets:
        if sheet.name != 'Filtro':
            ultima_linha = sheet.range('S' + str(sheet.cells.last_cell.row)).end('up').row
            notas = sheet.range(f'S2:S{ultima_linha}').value
            insumos = sheet.range(f'T2:T{ultima_linha}').value
            fornecedores = sheet.range(f'U2:U{ultima_linha}').value
            for nota, insumo, fornecedor in zip(notas, insumos, fornecedores):
                if nota and insumo:
                    dados_consolidados.append([nota, fornecedor, insumo])
    sheet_consolidado.range('A2').value = dados_consolidados
    
    sheet_consolidado.range('A:A').column_width = 20  # Largura da coluna A
    sheet_consolidado.range('B:B').column_width = 85  # Largura da coluna B
    sheet_consolidado.range('C:C').column_width = 70  # Largura da coluna C
    
    aplicar_filtro(sheet_consolidado)
    
    logging.info("Dados consolidados na planilha 'Filtro'.")

def apagar_stu(wb):
    for sheet in wb.sheets:
        if sheet.name != 'Filtro':
            sheet.range('S:U').clear_contents()
            logging.info(f"Colunas S, T e U apagadas na planilha {sheet.name}")
            
def apagar_planilhas(wb):
    for sheet in wb.sheets:
        if 'Planilha' in sheet.name:
            logging.info(f"Deletando a planilha: {sheet.name}")
            sheet.delete()
    logging.info("Planilhas com prefixo removido.")
    
def selecionar_arquivo():
    arquivo = filedialog.askopenfilename(title="Selecione um arquivo Excel", filetypes=[("Excel Files", "*.xlsx")])
    if arquivo:
        progress_bar.start()
        Thread(target=lambda: processar_arquivo(arquivo)).start()
        root.after(100, check_processing)

def check_processing():
    if logging.getLogger().hasHandlers():
        progress_bar.stop()

# Configuração da Interface Gráfica
root = tk.Tk()
root.title("Processador de Arquivo Excel")
root.geometry("700x500")

# Configuração do Text widget para exibir o log
log_text = tk.Text(root, wrap="word", state="disabled")
log_text.pack(padx=10, pady=10, expand=True, fill="both")

# Barra de progresso para indicar carregamento
progress_bar = ttk.Progressbar(root, orient="horizontal", mode="indeterminate")
progress_bar.pack(pady=10, fill="x")

# Botão para selecionar o arquivo
botao_selecionar = tk.Button(root, text="Selecionar arquivo .xlsx", command=selecionar_arquivo)
botao_selecionar.pack(pady=10)

# Configuração do logger para enviar os logs ao Text widget
setup_logger(log_text)

root.mainloop()
