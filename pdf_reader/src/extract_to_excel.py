import pdfplumber
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
import logging

# Configuração de logging
logging.basicConfig(
    level=logging.INFO,
    filename='log.txt',
    format='%(asctime)s - %(levelname)s - %(message)s'
)

def process_pdf(pdf_path):
    registros_estruturados = []

    with pdfplumber.open(pdf_path) as pdf:
        logging.info(f"PDF carregado com sucesso: {pdf_path}")

        for page_num, page in enumerate(pdf.pages, start=1):
            logging.info(f"Processando página {page_num}")

            for table in page.extract_tables():
                header = table[0]
                data_rows = table[1:]

                if header[0] and "PLANTE" in header[0]:
                    lbl0 = header[0].split('\n')[0].strip()
                    lbl1 = header[1].split('\n')[0].strip() if len(header) > 1 else ""
                    registros_estruturados.append({"Nome": lbl0, "Resultado": lbl1})
                    fallback = header[0].split('\n')[-1].strip()

                    for row in data_rows:
                        left = (row[0] or "").strip()
                        right = (row[1] or "").strip() if len(row) > 1 else ""
                        if left:
                            registros_estruturados.append({"Nome": left, "Resultado": right})
                        elif right:
                            registros_estruturados.append({"Nome": fallback, "Resultado": right})
                    continue

    logging.info(f"Registros estruturados: {len(registros_estruturados)}")
    return pd.DataFrame(registros_estruturados)

def main():
    root = tk.Tk()
    root.withdraw()

    pdf_path = filedialog.askopenfilename(
        title="Selecione o PDF de entrada",
        filetypes=[("PDF files", "*.pdf")]
    )
    if not pdf_path:
        messagebox.showinfo("Cancelado", "Nenhum PDF selecionado.")
        return

    df_tabela = process_pdf(pdf_path)

    if df_tabela.empty:
        messagebox.showerror("Erro", "Nenhum dado de tabela extraído do PDF.")
        return

    excel_template_path = filedialog.askopenfilename(
        title="Selecione o Excel Template",
        filetypes=[("Excel files", "*.xlsx")]
    )
    if not excel_template_path:
        messagebox.showinfo("Cancelado", "Nenhum template selecionado.")
        return

    save_path = filedialog.asksaveasfilename(
        title="Salvar Excel como",
        defaultextension=".xlsx",
        filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
        initialfile="resultado_final.xlsx"
    )
    if not save_path:
        messagebox.showinfo("Cancelado", "Operação cancelada.")
        return

    try:
        excel_template = pd.ExcelFile(excel_template_path)
        novo_df = excel_template.parse('Novo')

        plante_nome = df_tabela.loc[0, 'Resultado']
        codigo_plante = df_tabela.loc[1, 'Resultado']
        supplier_nome = df_tabela.loc[2, 'Resultado']
        dados_lidos = df_tabela.iloc[1:].reset_index(drop=True)

        for idx, row in dados_lidos.iterrows():
            linha_excel = idx + 2
            novo_df.loc[linha_excel, novo_df.columns[1]] = plante_nome
            novo_df.loc[linha_excel, novo_df.columns[4]] = codigo_plante
            novo_df.loc[linha_excel, novo_df.columns[5]] = supplier_nome

        with pd.ExcelWriter(save_path, engine="openpyxl") as writer:
            df_tabela.to_excel(writer, sheet_name="Planilha1", index=False)
            novo_df.to_excel(writer, sheet_name="Novo", index=False)

        messagebox.showinfo("Sucesso", f"Arquivo salvo com sucesso em:\n{save_path}")
        logging.info(f"Dados gravados com sucesso em: {save_path}")

    except PermissionError:
        messagebox.showerror("Erro", "Sem permissão para salvar o arquivo.")
        logging.error("Erro: Sem permissão para salvar o arquivo.")
    except FileNotFoundError:
        messagebox.showerror("Erro", "Diretório não encontrado.")
        logging.error("Erro: Diretório não encontrado.")
    except Exception as e:
        messagebox.showerror("Erro ao salvar", str(e))
        logging.error(f"Erro desconhecido ao salvar: {e}")

if __name__ == "__main__":
    main()
