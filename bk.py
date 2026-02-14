import pandas as pd
import os
import ttkbootstrap as tb
from ttkbootstrap.constants import *
from tkinter import filedialog, messagebox
import threading

AUTOR = "António Vanga II"
VERSAO = "1.0.0"

# ==============================
# REGRAS DE COMISSÃO
# ==============================
def calcular_percentagem(total):
    if total <= 300_000:
        return 0.05
    elif total <= 700_000:
        return 0.10
    elif total <= 1_000_000:
        return 0.12
    else:
        return 0.15

# ==============================
# LIMPAR VALORES
# ==============================
def limpar_valor(valor):
    if pd.isna(valor):
        return None
    texto = "".join(c for c in str(valor) if c in "0123456789.,")
    if "," in texto and "." in texto:
        if texto.rfind(",") > texto.rfind("."):
            texto = texto.replace(".", "").replace(",", ".")
        else:
            texto = texto.replace(",", "")
    elif "," in texto:
        texto = texto.replace(",", ".")
    try:
        return float(texto)
    except ValueError:
        return None

# ==============================
# LEITURA
# ==============================
def ler_arquivo(caminho):
    ext = os.path.splitext(caminho)[1].lower()
    if ext in (".xlsx", ".xlsm"):
        return pd.read_excel(caminho, engine="openpyxl")
    elif ext == ".xls":
        return pd.read_excel(caminho, engine="xlrd")
    elif ext == ".csv":
        return pd.read_csv(caminho, engine="python")
    elif ext == ".ods":
        return pd.read_excel(caminho, engine="odf")
    else:
        raise ValueError("Formato não suportado")

# ==============================
# PROCESSAMENTO
# ==============================
def processar(caminho):
    df = ler_arquivo(caminho)
    df.columns = df.columns.str.strip()

    if not {"Nome Utilizador", "Total"}.issubset(df.columns):
        raise ValueError("Colunas obrigatórias: Nome Utilizador | Total")

    df["Nome Utilizador"] = df["Nome Utilizador"].astype(str).str.strip()
    df["Total_num"] = df["Total"].apply(limpar_valor)
    df = df[df["Total_num"].notna() & (df["Nome Utilizador"] != "")]
    return df.reset_index(drop=True)

# ==============================
# RESUMO
# ==============================
def gerar_resumo(df):
    resumo = df.groupby("Nome Utilizador", as_index=False)["Total_num"].sum()
    resumo["Percentagem"] = resumo["Total_num"].apply(calcular_percentagem)
    resumo["Comissão"] = resumo["Total_num"] * resumo["Percentagem"]
    resumo.rename(columns={"Total_num": "Total Feito"}, inplace=True)
    return resumo

# ==============================
# JANELA
# ==============================
root = tb.Window(themename="litera")
root.title("Sistema de Comissões XD - Vendedor")
root.geometry("950x600")
root.resizable(False, False)

status = tb.StringVar(value="Pronto")
vendedor_var = tb.StringVar(value="Todos")
gerar_comissao_var = tb.StringVar(value="Sim")
df_preview = pd.DataFrame()

# ==============================
# MENU
# ==============================
menu_bar = tb.Menu(root)
root.config(menu=menu_bar)

# Menu Arquivo
arquivo_menu = tb.Menu(menu_bar, tearoff=0)
menu_bar.add_cascade(label="Arquivo", menu=arquivo_menu)
arquivo_menu.add_command(label="Importar", command=lambda: importar())
arquivo_menu.add_command(label="Gerar Resumo", command=lambda: salvar_resumo())
arquivo_menu.add_command(label="Excluir Linha Selecionada", command=lambda: excluir_linhas())
arquivo_menu.add_separator()
arquivo_menu.add_command(label="Sair", command=root.destroy)

# Menu Ajuda
ajuda_menu = tb.Menu(menu_bar, tearoff=0)
menu_bar.add_cascade(label="Ajuda", menu=ajuda_menu)
ajuda_menu.add_command(label="Sobre", command=lambda: messagebox.showinfo(
    "Sobre",
    f"Sistema de Comissões XD - Vendedor v{VERSAO}\nCriado por {AUTOR}"
))

# ==============================
# HEADER
# ==============================
header = tb.Frame(root, padding=(20, 15))
header.pack(fill=X)

tb.Label(header, text="Sistema de Comissões XD - Vendedor ",
         font=("Segoe UI", 18, "bold")).pack(anchor="w")

tb.Label(header, text=f"Criado por {AUTOR}",
         font=("Segoe UI", 9),
         bootstyle="secondary").pack(anchor="w")

# ==============================
# AÇÕES
# ==============================
actions = tb.Frame(root, padding=(20, 10))
actions.pack(fill=X)

tb.Button(actions, text="Importar",
          bootstyle="primary-outline",
          width=14,
          command=lambda: importar()).pack(side=LEFT, padx=(0, 8))

tb.Button(actions, text="Gerar Resumo",
          bootstyle="success-outline",
          width=14,
          command=lambda: salvar_resumo()).pack(side=LEFT)

tb.Button(actions, text="Excluir Linha",
          bootstyle="danger-outline",
          width=14,
          command=lambda: excluir_linhas()).pack(side=LEFT, padx=(8,0))

tb.Label(actions, text="Comissões:").pack(side=LEFT, padx=(15, 5))

combo_comissao = tb.Combobox(
    actions,
    textvariable=gerar_comissao_var,
    values=["Sim", "Não"],
    state="readonly",
    width=8
)
combo_comissao.pack(side=LEFT)

tb.Frame(actions).pack(side=LEFT, expand=True)

tb.Label(actions, text="Visualização:").pack(side=LEFT, padx=(0, 5))

combo_vendedor = tb.Combobox(
    actions,
    textvariable=vendedor_var,
    state="readonly",
    width=30
)
combo_vendedor.pack(side=LEFT)

# ==============================
# TABELA
# ==============================
table_frame = tb.Frame(root, padding=(20, 10))
table_frame.pack(fill=BOTH, expand=True)

tree = tb.Treeview(
    table_frame,
    columns=("Nome Utilizador", "Total"),
    show="headings",
    height=16
)
tree.heading("Nome Utilizador", text="Nome Utilizador")
tree.heading("Total", text="Total")
tree.column("Nome Utilizador", width=600)
tree.column("Total", width=200, anchor=E)
tree.pack(fill=BOTH, expand=True)

# ==============================
# STATUS
# ==============================
footer = tb.Frame(root, padding=(20, 8))
footer.pack(fill=X)
tb.Label(footer, textvariable=status).pack(anchor="w")

# ==============================
# LOADING
# ==============================
loading_running = False
loading_overlay = tb.Frame(root)
loading_overlay.place(relx=0, rely=0, relwidth=1, relheight=1)
loading_overlay.lower()

loading_card = tb.Frame(loading_overlay, padding=20, bootstyle="light")
loading_card.place(relx=0.5, rely=0.5, anchor="center")

tb.Label(loading_card, text="A processar...",
         font=("Segoe UI", 11)).pack(pady=(0, 10))

loading_meter = tb.Meter(
    loading_card,
    interactive=False,
    metersize=70,
    amountused=0,
    amounttotal=100
)
loading_meter.pack()

def animar_loading():
    if not loading_running:
        return
    loading_meter.amountusedvar.set(
        (loading_meter.amountusedvar.get() + 4) % 100
    )
    root.after(50, animar_loading)

def mostrar_loading():
    global loading_running
    loading_running = True
    loading_overlay.lift()
    animar_loading()

def esconder_loading():
    global loading_running
    loading_running = False
    loading_overlay.lower()

# ==============================
# FUNÇÕES UI
# ==============================
def atualizar_combo():
    valores = ["Todos", "Resumo por vendedor"]
    valores.extend(sorted(df_preview["Nome Utilizador"].unique()))
    combo_vendedor["values"] = valores
    vendedor_var.set("Todos")
    filtrar()

def filtrar(event=None):
    tree.delete(*tree.get_children())
    opcao = vendedor_var.get()

    if opcao == "Resumo por vendedor":
        df = df_preview.groupby("Nome Utilizador", as_index=False)["Total_num"].sum()
        for _, r in df.iterrows():
            total = f"{r['Total_num']:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
            tree.insert("", END, values=(r["Nome Utilizador"], total))

    elif opcao == "Todos":
        for _, r in df_preview.iterrows():
            total = f"{r['Total_num']:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
            tree.insert("", END, values=(r["Nome Utilizador"], total))

    else:
        df_vendedor = df_preview[df_preview["Nome Utilizador"] == opcao]
        total_vendedor = df_vendedor["Total_num"].sum()

        for _, r in df_vendedor.iterrows():
            total = f"{r['Total_num']:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
            tree.insert("", END, values=(r["Nome Utilizador"], total))

        total_fmt = f"{total_vendedor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        tree.insert("", END, values=("Total =", total_fmt))

combo_vendedor.bind("<<ComboboxSelected>>", filtrar)

# ==============================
# IMPORTAR
# ==============================
def importar():
    caminho = filedialog.askopenfilename(
        filetypes=[("Excel / CSV / ODS", "*.xlsx *.xls *.csv *.ods")]
    )
    if not caminho:
        return

    mostrar_loading()

    def task():
        global df_preview
        try:
            df_preview = processar(caminho)
            root.after(0, atualizar_combo)
            root.after(0, lambda: status.set("Ficheiro importado com sucesso"))
        except Exception as e:
            root.after(0, lambda: messagebox.showerror("Erro", str(e)))
        finally:
            root.after(0, esconder_loading)

    threading.Thread(target=task, daemon=True).start()

# ==============================
# SALVAR RESUMO
# ==============================
def salvar_resumo():
    if df_preview.empty:
        messagebox.showwarning("Aviso", "Sem dados.")
        return

    opcao = vendedor_var.get()
    df = df_preview.copy()

    if opcao not in ["Todos", "Resumo por vendedor"]:
        df = df[df["Nome Utilizador"] == opcao]

    resumo = gerar_resumo(df)

    if gerar_comissao_var.get() == "Não":
        resumo = resumo[["Nome Utilizador", "Total Feito"]]

    caminho = filedialog.asksaveasfilename(
        defaultextension=".xlsx",
        initialfile="Resumo_Comissoes.xlsx"
    )
    if not caminho:
        return

    with pd.ExcelWriter(caminho, engine="xlsxwriter") as writer:
        resumo.to_excel(writer, index=False, sheet_name="Resumo")

        workbook = writer.book
        worksheet = writer.sheets["Resumo"]

        money_fmt = workbook.add_format({'num_format': '#,##0.00'})
        worksheet.set_column("B:B", 18, money_fmt)

        if "Percentagem" in resumo.columns:
            worksheet.set_column("C:C", 12, workbook.add_format({'num_format': '0%'}))
        if "Comissão" in resumo.columns:
            worksheet.set_column("D:D", 18, money_fmt)

    messagebox.showinfo("Sucesso", "Resumo gerado com sucesso.")

# ==============================
# EXCLUIR LINHAS
# ==============================
def excluir_linhas():
    selecionados = tree.selection()
    if not selecionados:
        messagebox.showwarning("Aviso", "Nenhuma linha selecionada para excluir.")
        return

    confirmar = messagebox.askyesno("Confirmar", "Deseja realmente excluir as linhas selecionadas?")
    if not confirmar:
        return

    global df_preview
    opcao = vendedor_var.get()

    for item in selecionados:
        valores = tree.item(item, "values")
        nome = valores[0]
        total = valores[1].replace(".", "").replace(",", ".")
        try:
            total = float(total)
        except:
            total = None

        if total is None:
            continue
        if opcao == "Todos":
            df_preview = df_preview.drop(df_preview[(df_preview["Nome Utilizador"] == nome) & (df_preview["Total_num"] == total)].index)
        elif opcao == "Resumo por vendedor":
            df_preview = df_preview[df_preview["Nome Utilizador"] != nome]
        else:
            df_preview = df_preview.drop(df_preview[(df_preview["Nome Utilizador"] == nome) & (df_preview["Total_num"] == total)].index)

    df_preview = df_preview.reset_index(drop=True)
    filtrar()

root.mainloop()
