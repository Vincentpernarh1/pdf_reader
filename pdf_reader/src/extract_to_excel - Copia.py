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
    df_last_cap = pd.DataFrame()

    with pdfplumber.open(pdf_path) as pdf:
        logging.info(f'PDF carregado com sucesso: {pdf_path}')
        print(f'PDF carregado com sucesso: {pdf_path}')
        
        for page_num, page in enumerate(pdf.pages, start=1):
            logging.info(f'Processando página {page_num}')

            for table in page.extract_tables():
                if not table or not table[0]:
                    continue

                header = [h.strip().upper() if h else "" for h in table[0]]
                data_rows = table[1:]

                # Tabela com colunas: "#", "PN", "DESCRIPTION"
                if header[:3] == ["#", "PN", "DESCRIPTION"]:
                    for row in data_rows:
                        if len(row) < 3:
                            continue
                        linha = (row[0] or "").strip()
                        pns = (row[1] or "").strip().split("\n")
                        desc = (row[2] or "").strip()
                        for pn in pns:
                            registros_estruturados.append({
                                "Nome": linha,
                                "Resultado": pn.strip(),
                                "Descricao": desc
                            })
                    continue

                # Tabela Last Capacity
                if "LAST CAPACITY" in header[0]:
                    df_last_cap = pd.DataFrame(data_rows, columns=[
                        "Shift/Day", "Hours/Shift", "Days/Week", "Parts/Day", "Parts/Week"
                    ])
                    df_last_cap = df_last_cap.dropna(how="all")
                    continue

                # Tabelas padrão (2 colunas)
                header = table[0]
                data_rows = table[1:]

                nome_cabecalho = (header[0] or "").split('\n')[0].strip()
                valor_cabecalho = header[1].strip() if len(header) > 1 and header[1] else ""
                campo = nome_cabecalho.upper()

                if campo == "PROTOCOLO" and not valor_cabecalho:
                    if data_rows:
                        candidato = (data_rows[0][0] or "").strip()
                        if candidato.isdigit():
                            valor_cabecalho = candidato

                elif campo == "SUMMARY CAPACITY INCREASE / SHORT DESCRIPTION" and not valor_cabecalho:
                    if data_rows:
                        candidato = (data_rows[0][1] or "").strip() if len(data_rows[0]) > 1 else ""
                        if candidato:
                            valor_cabecalho = candidato

                if nome_cabecalho:
                    registros_estruturados.append({"Nome": nome_cabecalho, "Resultado": valor_cabecalho})

                fallback = header[0].split('\n')[-1].strip()

                for row in data_rows:
                    left = (row[0] or "").strip()
                    right = (row[1] or "").strip() if len(row) > 1 else ""

                    PNs = left.split(" ")
                    print(PNs)

                    if left:
                        registros_estruturados.append({"Nome": left, "Resultado": right})
                    elif right:
                        registros_estruturados.append({"Nome": fallback, "Resultado": right})

    df_novo = pd.DataFrame(registros_estruturados)
    print("\n📋 DataFrame completo (df_novo):")
    print(df_novo)
    return df_novo, df_last_cap

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

    df_novo, df_last_cap = process_pdf(pdf_path)



    

    if df_novo.empty:
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
        filetypes=[("Excel files", "*.xlsx")]
    )
    if not save_path:
        messagebox.showinfo("Cancelado", "Operação cancelada.")
        return

    try:
        excel_template = pd.ExcelFile(excel_template_path)
        novo_df = excel_template.parse('Novo')

        plante_nome = df_novo.loc[0, 'Resultado']
        codigo_plante = df_novo.loc[1, 'Resultado']
        supplier_nome = df_novo.loc[2, 'Resultado']
        analist_dhl = df_novo.loc[4, 'Resultado']

        dados_lidos = df_novo[df_novo['Descricao'].notnull()].reset_index(drop=True)

        for idx, row in dados_lidos.iterrows():
            linha_excel = idx + 2
            novo_df.loc[linha_excel, novo_df.columns[1]] = plante_nome
            novo_df.loc[linha_excel, novo_df.columns[3]] = analist_dhl
            novo_df.loc[linha_excel, novo_df.columns[4]] = codigo_plante
            novo_df.loc[linha_excel, novo_df.columns[5]] = supplier_nome
            novo_df.loc[linha_excel, novo_df.columns[6]] = row['Resultado']
            novo_df.loc[linha_excel, novo_df.columns[7]] = row['Descricao']

        with pd.ExcelWriter(save_path, engine="openpyxl") as writer:
            df_novo.to_excel(writer, sheet_name="Planilha1", index=False)
            novo_df.to_excel(writer, sheet_name="Novo", index=False)
            if not df_last_cap.empty:
                df_last_cap.to_excel(writer, sheet_name="Last Capacity", index=False)

        messagebox.showinfo("Sucesso", f"Arquivo salvo com sucesso em:\n{save_path}")
        logging.info(f"Dados gravados com sucesso em: {save_path}")

    except PermissionError:
        messagebox.showerror("Erro", "Sem permissão para salvar o arquivo.")
    except FileNotFoundError:
        messagebox.showerror("Erro", "Diretório não encontrado.")
    except Exception as e:
        messagebox.showerror("Erro ao salvar", str(e))

if __name__ == "__main__":
    main()