import os
import sys
import shutil
import threading
from pathlib import Path
import customtkinter as ctk
from tkinter import messagebox
import pythoncom
from gerar_termo import GeradorDeTermos

# ============================================================
# 💡 FUNÇÃO CRÍTICA PARA CORRIGIR CAMINHOS DO PYINSTALLER
# ============================================================

def recurso_executavel(relative_path):
    """Obtém o caminho absoluto para o recurso, seja no modo de desenvolvimento
    ou no modo de executável PyInstaller.
    """
    try:
        base_path = Path(sys._MEIPASS)
    except Exception:
        base_path = Path(os.path.abspath("."))
    return base_path / relative_path


# ============================================================
# ⚙️ CONFIGURAÇÕES E MAPEAMENTO DE USUÁRIOS
# ============================================================

USUARIOS_MAPEADOS = {
    "eversonsilva": "Everson",
    "raphaelsychocki": "Raphael",
    "arnaldosantos": "Arnaldo",
    "joaoguilherme": "João",
    "felipecorrea": "Felipe",
    "vitorstein": "Vitor",
    "alessandrovicente": "Alessandro",
    "alexandrecamargo": "Alexandre",
    "viniciusbonifacio": "Vinicius",
    "endersongolindano": "Telecom",
    "yasminbarbosa": "Telecom"
}

PASTA_LOCAL = Path.cwd()

# ============================================================
# 🎨 TEMA “CAQUI” — VISUAL LARANJA E ELEGANTE (modo escuro fixo)
# ============================================================

COLOR_PRIMARY = "#D87C00"
COLOR_PRIMARY_DARK = "#A85E00"
COLOR_BG = "#2C2A28"
COLOR_TEXT = "#F2E8DA"
PADDING_X = 30
SPACING_FIELD_Y = 18
HEIGHT_ENTRY = 36


# ============================================================
# 🧠 FUNÇÕES AUXILIARES
# ============================================================

def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)


def obter_usuario_windows():
    try:
        return os.getlogin()
    except Exception:
        return os.getenv("USERNAME") or "Desconhecido"


def localizar_pasta_drive_usuario(nome_drive):
    home = Path.home()
    caminhos_possiveis = [
        Path("G:/Meu Drive") / nome_drive,
        Path("D:/Meu Drive") / nome_drive,
        home / "Google Drive" / nome_drive,
        home / "Meu Drive" / nome_drive,
    ]

    local_appdata = os.getenv("LOCALAPPDATA")
    if local_appdata:
        drivefs = Path(local_appdata) / "Google" / "DriveFS"
        if drivefs.exists():
            for subdir in drivefs.iterdir():
                if subdir.is_dir():
                    caminho = subdir / "My Drive" / nome_drive
                    caminhos_possiveis.append(caminho)

    for caminho in caminhos_possiveis:
        if caminho.exists():
            return caminho

    raise FileNotFoundError(f"Não foi possível localizar a pasta 'Meu Drive/{nome_drive}'.")


def mover_para_pasta_drive(arquivo_pdf, pasta_drive):
    destino = pasta_drive / arquivo_pdf.name
    shutil.move(str(arquivo_pdf), str(destino))
    return destino


# ============================================================
# 🖥️ INTERFACE GRÁFICA
# ============================================================

class App(ctk.CTk):
    def __init__(self):
        super().__init__()

        # ==============================
        # JANELA PRINCIPAL
        # ==============================
        self.title("Caquímetro - Gerador de Termos")
        self.geometry("750x780")
        self.minsize(600, 700)
        
        # Ícone
        try:
            self.iconbitmap(resource_path("caqui.ico"))
        except ctk.TclError:
            print("Aviso: O arquivo 'caqui.ico' não foi encontrado ou está inacessível.")

        # 🔒 Tema escuro fixo
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("blue")
        self.configure(fg_color=COLOR_BG)

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(5, weight=1)

        # ==============================
        # TÍTULO
        # ==============================
        ctk.CTkLabel(
            self,
            text="🍊 Caquímetro - Gerador de Termos",
            font=("Segoe UI", 22, "bold"),
            text_color=COLOR_PRIMARY
        ).grid(row=0, column=0, pady=(25, 10))

        # ==============================
        # INFORMAÇÕES DO USUÁRIO
        # ==============================
        header_frame = ctk.CTkFrame(self, fg_color="transparent")
        header_frame.grid(row=1, column=0, padx=PADDING_X, pady=(5, 5), sticky="ew")
        header_frame.columnconfigure(0, weight=1)

        self.usuario = obter_usuario_windows()
        self.nome_drive = USUARIOS_MAPEADOS.get(self.usuario.lower(), self.usuario)

        user_info_frame = ctk.CTkFrame(header_frame, fg_color="transparent")
        user_info_frame.grid(row=0, column=0, sticky="w")

        ctk.CTkLabel(user_info_frame,
                     text=f"👤 Usuário Detectado: {self.usuario.upper()}",
                     font=("Segoe UI", 14, "bold"),
                     text_color=COLOR_TEXT).pack(anchor="w", pady=(5, 0))
        ctk.CTkLabel(user_info_frame,
                     text=f"📂 Drive Destino: Pasta {self.nome_drive}",
                     font=("Segoe UI", 12),
                     text_color="#C8BBA6").pack(anchor="w", pady=(0, 5))


        ctk.CTkFrame(self, height=2, fg_color=COLOR_PRIMARY).grid(
            row=2, column=0, padx=PADDING_X, pady=(15, 15), sticky="ew"
        )

        # ==============================
        # COMBOBOX DE TIPOS DE TERMO
        # ==============================
        ctk.CTkLabel(self, text="Selecione o Tipo de Termo:",
                     font=("Segoe UI", 14, "bold"),
                     text_color=COLOR_TEXT).grid(row=3, column=0, pady=(5, 0))

        self.tipo_termo = ctk.CTkComboBox(
            self,
            values=["Equipamento", "VPN", "Administrador Local","Telecom"],
            command=self.atualizar_campos,
            width=300, height=HEIGHT_ENTRY,
            fg_color="#3C3836", button_color=COLOR_PRIMARY,
            dropdown_hover_color=COLOR_PRIMARY_DARK

        )
        self.tipo_termo.set("Equipamento")
        self.tipo_termo.grid(row=4, column=0, pady=(5, 20))

        # ==============================
        # FRAME ROLÁVEL DE CAMPOS
        # ==============================
        self.frame_scroll = ctk.CTkScrollableFrame(self, fg_color="transparent")
        self.frame_scroll.grid(row=5, column=0, padx=PADDING_X, pady=10, sticky="nsew")
        self.frame_scroll.grid_columnconfigure(0, weight=1)

        self.frame_campos = ctk.CTkFrame(self.frame_scroll, fg_color="transparent")
        self.frame_campos.grid(row=0, column=0, sticky="new")
        self.frame_campos.columnconfigure(0, weight=1, uniform="col")
        self.frame_campos.columnconfigure(1, weight=1, uniform="col")

        self.campos = {}
        self.atualizar_campos("Equipamento")

        # ==============================
        # BOTÃO PRINCIPAL
        # ==============================
        ctk.CTkButton(
            self.frame_scroll,
            text="GERAR E ENVIAR TERMO",
            command=self.executar_processo_thread,
            width=300, height=55,
            font=("Segoe UI", 16, "bold"),
            corner_radius=12,
            fg_color=COLOR_PRIMARY,
            hover_color=COLOR_PRIMARY_DARK,
            text_color="white"
        ).grid(row=7, column=0, pady=(30, 15), columnspan=2)

        # Barra de progresso
        self.progress_bar = None


    # ============================================================
    # CAMPOS
    # ============================================================
    def atualizar_campos(self, tipo):
        for w in self.frame_campos.winfo_children():
            w.destroy()
        self.campos.clear()

        if tipo == "Equipamento":
            sections = [
                ("Colaborador", [
                    ("Nome", "CPF"),
                    ("Setor/Cargo", "Empresa")
                ]),
                ("Equipamento", [
                    ("Descrição do Equipamento", "Marca/Modelo"),
                    ("Série", "Patrimônio"),
                    ("Estado do equipamento", None),
                    ("Técnico responsável", None)
                ])
            ]
        elif tipo == "VPN":
            sections = [("VPN", [("Nome", "Cargo"), ("Departamento", None)])]

        elif tipo == "Telecom":
            sections = [
                ("Colaborador", [
                    ("Nome", "CPF"),
                    ("Setor/Cargo", "Empresa")
                ]),
                ("Equipamento", [
                    ("Descrição do Equipamento", "Marca/Modelo"),
                    ("Série", "Número da Linha"),
                    ("Técnico responsável", None)
                ])] 
        else:
            sections = [("Administrador Local", [("Nome", "CPF")])]

        row = 0
        for title, fields in sections:
            ctk.CTkLabel(self.frame_campos, text=title.upper(),
                         font=("Segoe UI", 13, "bold"), text_color=COLOR_PRIMARY).grid(
                row=row, column=0, columnspan=2, pady=(SPACING_FIELD_Y, 0), sticky="w")
            row += 1
            for f1, f2 in fields:
                ctk.CTkLabel(self.frame_campos, text=f1, anchor="w",
                             font=("Segoe UI", 12, "bold"), text_color=COLOR_TEXT).grid(row=row, column=0, sticky="w", padx=10)
                e1 = ctk.CTkEntry(self.frame_campos, height=HEIGHT_ENTRY, fg_color="#3A3835", border_color=COLOR_PRIMARY)
                e1.grid(row=row + 1, column=0, sticky="ew", padx=10, pady=(0, 10))
                self.campos[f1] = e1

                if f2:
                    ctk.CTkLabel(self.frame_campos, text=f2, anchor="w",
                                 font=("Segoe UI", 12, "bold"), text_color=COLOR_TEXT).grid(row=row, column=1, sticky="w", padx=10)
                    e2 = ctk.CTkEntry(self.frame_campos, height=HEIGHT_ENTRY, fg_color="#3A3835", border_color=COLOR_PRIMARY)
                    e2.grid(row=row + 1, column=1, sticky="ew", padx=10, pady=(0, 10))
                    self.campos[f2] = e2
                else:
                    e1.grid_configure(columnspan=2)
                row += 2


    # ============================================================
    # THREAD DO PROCESSAMENTO
    # ============================================================
    def executar_processo_thread(self):
        threading.Thread(target=self.executar_processo, daemon=True).start()


    def executar_processo(self):
        try:
            pythoncom.CoInitialize()
        except Exception:
            pass

        tipo = self.tipo_termo.get()
        dados = {k: e.get().strip() for k, e in self.campos.items()}

        if not all(dados.values()):
            messagebox.showwarning("Atenção", "Preencha todos os campos antes de continuar.")
            if self.progress_bar:
                self.progress_bar.stop()
                self.progress_bar.grid_forget()
            return

        if not self.progress_bar:
            self.progress_bar = ctk.CTkProgressBar(self, width=400, height=10, progress_color=COLOR_PRIMARY)
        self.progress_bar.grid(row=8, column=0, pady=(0, 25))
        self.progress_bar.configure(mode="indeterminate")
        self.progress_bar.start()

        try:
            gerador = GeradorDeTermos()
            if tipo == "Equipamento":
                pdf = gerador.preencher_termo_equipamento(
                    nome=dados["Nome"], cpf=dados["CPF"], setor=dados["Setor/Cargo"],
                    empresa=dados["Empresa"], equipamento=dados["Descrição do Equipamento"],
                    marca=dados["Marca/Modelo"], serie=dados["Série"],
                    patrimonio=dados["Patrimônio"], estadoequip=dados["Estado do equipamento"],
                    tecnico=dados['Técnico responsável']
                )

            elif tipo == "Telecom":
                pdf = gerador.preencher_termo_telecom(
                    nome=dados["Nome"], cpf=dados["CPF"], setor=dados["Setor/Cargo"],
                    empresa=dados["Empresa"], equipamento=dados["Descrição do Equipamento"],
                    marca=dados["Marca/Modelo"], serie=dados["Série"],
                    numero=dados["Número da Linha"],tecnico=dados['Técnico responsável']
                )

            elif tipo == "VPN":
                pdf = gerador.preencher_termo_vpn(dados["Nome"], dados["Cargo"], dados["Departamento"])
            else:
                pdf = gerador.preencher_termo_adm(dados["Nome"], dados["CPF"])

            pasta_drive = localizar_pasta_drive_usuario(self.nome_drive)
            destino = mover_para_pasta_drive(pdf, pasta_drive)
            messagebox.showinfo("Sucesso", f"✅ Termo gerado e enviado para:\n\n{destino}")

        except Exception as e:
            messagebox.showerror("Erro", f"Ocorreu um erro:\n{e}")
        finally:
            if self.progress_bar:
                self.progress_bar.stop()
                self.progress_bar.grid_forget()
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass


# ============================================================
# 🏁 EXECUÇÃO
# ============================================================

if __name__ == "__main__":
    app = App()
    app.mainloop()
