import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from openpyxl import Workbook
from openpyxl.styles import Border, Side, PatternFill, Alignment, Font
from openpyxl.utils import get_column_letter

# Estrutura de dados em memória
foranias = {}  # { 'Forania 1': { 'Grupo 1': [ {servo}, ... ] } }

# ---- Funções ----

def adicionar_forania():
    nome = entrada_forania.get().strip()
    if not nome:
        messagebox.showwarning("Aviso", "Informe o nome da Forania.")
        return
    if nome in foranias:
        messagebox.showinfo("Info", "Forania já existe.")
        return
    foranias[nome] = {}
    atualizar_combo_forania()
    entrada_forania.delete(0, tk.END)
    renderizar()

def adicionar_grupo():
    forania = combo_forania.get()
    nome = entrada_grupo.get().strip()
    if not (forania and nome):
        messagebox.showwarning("Aviso", "Informe a forania e o grupo.")
        return
    if nome in foranias[forania]:
        messagebox.showinfo("Info", "Grupo já existe.")
        return
    foranias[forania][nome] = []
    atualizar_combo_grupo()
    entrada_grupo.delete(0, tk.END)
    renderizar()

def adicionar_servo():
    forania = combo_forania.get()
    grupo = combo_grupo.get()
    nome = entrada_servo_nome.get().strip()
    if not (forania and grupo and nome):
        messagebox.showwarning("Aviso", "Preencha todos os campos.")
        return

    servo = {
        "nome": nome,
        "modulo_basico": var_modulo.get(),
        "exp_oracao": var_oracao.get(),
        "apostila1": var_ap1.get(),
        "apostila2": var_ap2.get()
    }

    foranias[forania][grupo].append(servo)
    entrada_servo_nome.delete(0, tk.END)
    for var in [var_modulo, var_oracao, var_ap1, var_ap2]:
        var.set(False)
    renderizar()

def atualizar_combo_forania():
    combo_forania['values'] = list(foranias.keys())

def atualizar_combo_grupo(*args):
    forania = combo_forania.get()
    if forania:
        combo_grupo['values'] = list(foranias[forania].keys())




# def exportar_excel():
#     if not foranias:
#         messagebox.showwarning("Aviso", "Nenhum dado para exportar.")
#         return

#     caminho = filedialog.asksaveasfilename(
#         defaultextension=".xlsx",
#         filetypes=[("Excel files", "*.xlsx")]
#     )

#     if not caminho:
#         return

#     wb = Workbook()
#     ws = wb.active
#     ws.title = "Servos"

#     # Estilos
#     thin_border = Border(
#         left=Side(style="thin"),
#         right=Side(style="thin"),
#         top=Side(style="thin"),
#         bottom=Side(style="thin")
#     )

#     row_index = 1

#     for forania, grupos in foranias.items():
#         start_row = row_index

#         # Cabeçalho da Forania
#         ws.merge_cells(start_row=row_index, start_column=1, end_row=row_index, end_column=6)
#         cell_forania = ws.cell(row=row_index, column=1, value=forania)
#         cell_forania.font = Font(bold=True)
#         cell_forania.alignment = Alignment(horizontal="center", vertical="center")
#         cell_forania.fill = PatternFill(start_color="FFD966", end_color="FFD966", fill_type="solid")
#         row_index += 1  # sem linha em branco aqui

#         for grupo, servos in grupos.items():
#             # Cabeçalho do Grupo
#             ws.merge_cells(start_row=row_index, start_column=1, end_row=row_index, end_column=6)
#             cell_grupo = ws.cell(row=row_index, column=1, value=grupo)
#             cell_grupo.font = Font(bold=True)
#             cell_grupo.alignment = Alignment(horizontal="center", vertical="center")
#             cell_grupo.fill = PatternFill(start_color="FF9999", end_color="FF9999", fill_type="solid")
#             row_index += 1

#             # Cabeçalhos da tabela (em negrito)
#             headers = ["Nome", "Módulo Básico", "Exp. Oração", "Apostila 1", "Apostila 2"]
#             for col, header in enumerate(headers, start=1):
#                 cell = ws.cell(row=row_index, column=col, value=header)
#                 cell.font = Font(bold=True)
#                 cell.alignment = Alignment(horizontal="center", vertical="center")
#                 cell.fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
#             row_index += 1

#             # Dados dos servos
#             for servo in servos:
#                 ws.cell(row=row_index, column=1, value=servo["nome"]).alignment = Alignment(horizontal="left", vertical="center")
#                 ws.cell(row=row_index, column=2, value="Sim" if servo["modulo_basico"] else "Não").alignment = Alignment(horizontal="center")
#                 ws.cell(row=row_index, column=3, value="Sim" if servo["exp_oracao"] else "Não").alignment = Alignment(horizontal="center")
#                 ws.cell(row=row_index, column=4, value="Sim" if servo["apostila1"] else "Não").alignment = Alignment(horizontal="center")
#                 ws.cell(row=row_index, column=5, value="Sim" if servo["apostila2"] else "Não").alignment = Alignment(horizontal="center")
#                 row_index += 1

#         # 🔹 Linha em branco entre foranias
#         row_index += 1  

#         # Bordas para o bloco
#         end_row = row_index - 1
#         for r in range(start_row, end_row + 1):
#             for c in range(1, 7):
#                 ws.cell(row=r, column=c).border = thin_border

#     # Ajuste de largura das colunas
#     from openpyxl.utils import get_column_letter
#     for i, col in enumerate(ws.columns, start=1):
#         max_length = 0
#         for cell in col:
#             try:
#                 if cell.value:
#                     max_length = max(max_length, len(str(cell.value)))
#             except AttributeError:
#                 pass
#         ws.column_dimensions[get_column_letter(i)].width = max_length + 2

#     wb.save(caminho)
#     messagebox.showinfo("Sucesso", f"Dados exportados com sucesso para:\n{caminho}")



def exportar_excel():
    if not foranias:
        messagebox.showwarning("Aviso", "Nenhum dado para exportar.")
        return

    caminho = filedialog.asksaveasfilename(
        defaultextension=".xlsx",
        filetypes=[("Excel files", "*.xlsx")]
    )

    if not caminho:
        return

    wb = Workbook()
    ws = wb.active
    ws.title = "Servos"

    # Estilos
    thin_border = Border(
        left=Side(style="thin"),
        right=Side(style="thin"),
        top=Side(style="thin"),
        bottom=Side(style="thin")
    )

    row_index = 1
    num_cols = 5  # Agora limitamos a 5 colunas (A até E)

    for forania, grupos in foranias.items():
        start_row = row_index

        # Cabeçalho da Forania
        ws.merge_cells(start_row=row_index, start_column=1, end_row=row_index, end_column=num_cols)
        cell_forania = ws.cell(row=row_index, column=1, value=forania)
        cell_forania.font = Font(bold=True)
        cell_forania.alignment = Alignment(horizontal="center", vertical="center")
        cell_forania.fill = PatternFill(start_color="FFD966", end_color="FFD966", fill_type="solid")
        row_index += 1

        for grupo, servos in grupos.items():
            # Cabeçalho do Grupo
            ws.merge_cells(start_row=row_index, start_column=1, end_row=row_index, end_column=num_cols)
            cell_grupo = ws.cell(row=row_index, column=1, value=grupo)
            cell_grupo.font = Font(bold=True)
            cell_grupo.alignment = Alignment(horizontal="center", vertical="center")
            cell_grupo.fill = PatternFill(start_color="FF9999", end_color="FF9999", fill_type="solid")
            row_index += 1

            # Cabeçalho da Tabela
            headers = ["Nome", "Módulo Básico", "Exp. Oração", "Apostila 1", "Apostila 2"]
            for col, header in enumerate(headers, start=1):
                cell = ws.cell(row=row_index, column=col, value=header)
                cell.font = Font(bold=True)
                cell.alignment = Alignment(horizontal="center", vertical="center")
                cell.fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
            row_index += 1

            # Linhas de dados
            for servo in servos:
                ws.cell(row=row_index, column=1, value=servo["nome"]).alignment = Alignment(horizontal="left", vertical="center")
                ws.cell(row=row_index, column=2, value="Sim" if servo["modulo_basico"] else "Não").alignment = Alignment(horizontal="center")
                ws.cell(row=row_index, column=3, value="Sim" if servo["exp_oracao"] else "Não").alignment = Alignment(horizontal="center")
                ws.cell(row=row_index, column=4, value="Sim" if servo["apostila1"] else "Não").alignment = Alignment(horizontal="center")
                ws.cell(row=row_index, column=5, value="Sim" if servo["apostila2"] else "Não").alignment = Alignment(horizontal="center")
                row_index += 1

        # 🔹 Linha em branco total (sem bordas, sem preenchimento)
        row_index += 1
        for c in range(1, num_cols + 1):
            cell = ws.cell(row=row_index, column=c, value=None)
            cell.border = Border()  # Sem bordas
            cell.fill = PatternFill()  # Sem cor

        row_index += 1  # avança para a próxima forania

        # Bordas
        end_row = row_index - 2
        for r in range(start_row, end_row + 1):
            for c in range(1, num_cols + 1):
                cell = ws.cell(row=r, column=c)
                if cell.value:
                    cell.border = thin_border

    # Ajuste de largura das colunas
    from openpyxl.utils import get_column_letter
    for i in range(1, num_cols + 1):
        max_length = 0
        for cell in ws[get_column_letter(i)]:
            if cell.value:
                max_length = max(max_length, len(str(cell.value)))
        ws.column_dimensions[get_column_letter(i)].width = max_length + 2

    wb.save(caminho)
    messagebox.showinfo("Sucesso", f"Dados exportados com sucesso para:\n{caminho}")

def renderizar():
    for widget in frame_resultados.winfo_children():
        widget.destroy()

    for forania, grupos in foranias.items():
        lf_forania = ttk.LabelFrame(frame_resultados, text=f"Forania: {forania}")
        lf_forania.pack(fill='x', padx=5, pady=5, anchor='w')

        for grupo, servos in grupos.items():
            lf_grupo = ttk.LabelFrame(lf_forania, text=f"Grupo: {grupo}")
            lf_grupo.pack(fill='x', padx=10, pady=5)

            tree = ttk.Treeview(lf_grupo, columns=("Nome", "Módulo Básico", "Exp. Oração", "Apostila 1", "Apostila 2"), show="headings")
            for col in tree["columns"]:
                tree.heading(col, text=col)
                tree.column(col, width=120)
            for servo in servos:
                tree.insert('', 'end', values=(
                    servo["nome"],
                    "Sim" if servo["modulo_basico"] else "Não",
                    "Sim" if servo["exp_oracao"] else "Não",
                    "Sim" if servo["apostila1"] else "Não",
                    "Sim" if servo["apostila2"] else "Não"
                ))
            tree.pack(fill='x')

# ---- UI principal ----
janela = tk.Tk()
janela.title("Cadastro de Foranias, Grupos e Servos")
janela.geometry("850x700")

# --- Frame de Cadastro ---
frame_cadastro = ttk.LabelFrame(janela, text="Cadastro")
frame_cadastro.pack(fill='x', padx=10, pady=10)

# Forania
ttk.Label(frame_cadastro, text="Nova Forania:").grid(row=0, column=0, sticky='e')
entrada_forania = ttk.Entry(frame_cadastro)
entrada_forania.grid(row=0, column=1)
ttk.Button(frame_cadastro, text="Adicionar Forania", command=adicionar_forania).grid(row=0, column=2, padx=5)

# Grupo
ttk.Label(frame_cadastro, text="Nova Grupo:").grid(row=1, column=0, sticky='e')
combo_forania = ttk.Combobox(frame_cadastro, state="readonly")
combo_forania.grid(row=1, column=1)
combo_forania.bind("<<ComboboxSelected>>", atualizar_combo_grupo)

entrada_grupo = ttk.Entry(frame_cadastro)
entrada_grupo.grid(row=1, column=2)
ttk.Button(frame_cadastro, text="Adicionar Grupo", command=adicionar_grupo).grid(row=1, column=3)

# Servo
ttk.Label(frame_cadastro, text="Grupo (Seleção):").grid(row=2, column=0, sticky='e')
combo_grupo = ttk.Combobox(frame_cadastro, state="readonly")
combo_grupo.grid(row=2, column=1)

ttk.Label(frame_cadastro, text="Nome do Servo:").grid(row=3, column=0, sticky='e')
entrada_servo_nome = ttk.Entry(frame_cadastro)
entrada_servo_nome.grid(row=3, column=1)

# Checkbuttons
var_modulo = tk.BooleanVar()
var_oracao = tk.BooleanVar()
var_ap1 = tk.BooleanVar()
var_ap2 = tk.BooleanVar()

ttk.Checkbutton(frame_cadastro, text="Módulo Básico", variable=var_modulo).grid(row=4, column=0, sticky='w')
ttk.Checkbutton(frame_cadastro, text="Exp. Oração", variable=var_oracao).grid(row=4, column=1, sticky='w')
ttk.Checkbutton(frame_cadastro, text="Apostila 1", variable=var_ap1).grid(row=4, column=2, sticky='w')
ttk.Checkbutton(frame_cadastro, text="Apostila 2", variable=var_ap2).grid(row=4, column=3, sticky='w')

ttk.Button(frame_cadastro, text="Adicionar Servo", command=adicionar_servo).grid(row=5, column=1, pady=5)

# --- Exportar ---
ttk.Button(janela, text="Exportar para Excel", command=exportar_excel).pack(pady=5)

# --- Frame com Scrollbar para resultados ---
frame_scroll = ttk.Frame(janela)
frame_scroll.pack(fill='both', expand=True, padx=10, pady=10)

canvas = tk.Canvas(frame_scroll)
scrollbar = ttk.Scrollbar(frame_scroll, orient='vertical', command=canvas.yview)
scrollable_frame = ttk.Frame(canvas)

scrollable_frame.bind(
    "<Configure>",
    lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
)

canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
canvas.configure(yscrollcommand=scrollbar.set)

canvas.pack(side="left", fill='both', expand=True)
scrollbar.pack(side="right", fill="y")

frame_resultados = scrollable_frame

janela.mainloop()
