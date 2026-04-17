from tkinter import filedialog, messagebox

class UIHelpers:
    def __init__(self, root):
        self.root = root

    # --- DIÁLOGOS DE ARCHIVOS ---
    def ask_save_file(self, **kwargs):
        return filedialog.asksaveasfilename(parent=self.root, **kwargs)

    def ask_open_file(self, **kwargs):
        return filedialog.askopenfilename(parent=self.root, **kwargs)

    def ask_directory(self, **kwargs):
        return filedialog.askdirectory(parent=self.root, **kwargs)

    # --- MÉTODOS PRIVADOS PARA MANEJAR PARÁMETROS ---
    def _parse_args(self, a, b, default_title):
        if b is None:
            return default_title, a  # Título por defecto, 'a' es el mensaje
        return a, b  # 'a' es título, 'b' es mensaje

    # --- MESSAGEBOX (Todos con el mismo estándar: Título, Mensaje) ---
    def show_info(self, title_or_msg, message=None):
        t, m = self._parse_args(title_or_msg, message, "Información")
        messagebox.showinfo(t, m, parent=self.root)

    def show_error(self, title_or_msg, message=None):
        t, m = self._parse_args(title_or_msg, message, "Error")
        messagebox.showerror(t, m, parent=self.root)

    def show_warning(self, title_or_msg, message=None):
        t, m = self._parse_args(title_or_msg, message, "Advertencia")
        messagebox.showwarning(t, m, parent=self.root)

    def ask_yes_no(self, title_or_msg, message=None):
        t, m = self._parse_args(title_or_msg, message, "Confirmación")
        return messagebox.askyesno(t, m, parent=self.root)

    def ask_ok_cancel(self, title_or_msg, message=None):
        t, m = self._parse_args(title_or_msg, message, "Confirmación")
        return messagebox.askokcancel(t, m, parent=self.root)

    def ask_retry_cancel(self, title_or_msg, message=None):
        t, m = self._parse_args(title_or_msg, message, "Reintentar")
        return messagebox.askretrycancel(t, m, parent=self.root)

# ------------------------
# OPCIONAL: FUNCIONES ESTÁTICAS (Limpias)
# ------------------------
class UIHelpersStatic:
    @staticmethod
    def show_error(root, title, message):
        messagebox.showerror(title, message, parent=root)

    @staticmethod
    def show_info(root, title, message):
        messagebox.showinfo(title, message, parent=root)
