import os
from abc import ABC, abstractmethod
from tkinter import Tk, Label, Button, filedialog, StringVar, Radiobutton, Entry, messagebox
from docx2pdf import convert as docx_to_pdf
from pdf2docx import Converter

# Clase base abstracta
class Documento(ABC):
    def __init__(self, ruta_archivo):
        self.ruta = ruta_archivo

    @abstractmethod
    def convertir(self, ruta_salida):
        pass

# Word → PDF
class DocumentoWord(Documento):
    def convertir(self, ruta_salida):
        if not self.ruta.lower().endswith(".docx"):
            raise ValueError("El archivo de entrada no es .docx")
        if not ruta_salida.lower().endswith(".pdf"):
            raise ValueError("La salida debe terminar en .pdf")
        if not os.path.isfile(self.ruta):
            raise FileNotFoundError(f"El archivo no existe: {self.ruta}")
        docx_to_pdf(self.ruta, ruta_salida)

# PDF → Word
class DocumentoPDF(Documento):
    def convertir(self, ruta_salida):
        if not self.ruta.lower().endswith(".pdf"):
            raise ValueError("El archivo de entrada no es .pdf")
        if not ruta_salida.lower().endswith(".docx"):
            raise ValueError("La salida debe terminar en .docx")
        if not os.path.isfile(self.ruta):
            raise FileNotFoundError(f"El archivo no existe: {self.ruta}")
        cv = Converter(self.ruta)
        cv.convert(ruta_salida, start=0, end=None)
        cv.close()

# Clase convertidora
class Convertidor:
    def __init__(self, documento: Documento):
        self.documento = documento

    def ejecutar_conversion(self, ruta_salida):
        self.documento.convertir(ruta_salida)

# Interfaz gráfica con tkinter
class Aplicacion:
    def __init__(self, master):
        self.master = master
        master.title("Conversor Word ↔ PDF")

        self.tipo_conversion = StringVar(value="word_to_pdf")
        self.ruta_archivo = ""
        
        Label(master, text="Seleccione el tipo de conversión:").pack()
        Radiobutton(master, text="Word → PDF", variable=self.tipo_conversion, value="word_to_pdf").pack()
        Radiobutton(master, text="PDF → Word", variable=self.tipo_conversion, value="pdf_to_word").pack()

        Button(master, text="Seleccionar archivo", command=self.seleccionar_archivo).pack(pady=5)
        self.label_archivo = Label(master, text="Ningún archivo seleccionado", fg="gray")
        self.label_archivo.pack()

        Label(master, text="Ruta de salida:").pack()
        self.entry_salida = Entry(master, width=50)
        self.entry_salida.pack(pady=5)

        Button(master, text="Convertir", command=self.convertir).pack(pady=10)

    def seleccionar_archivo(self):
        tipo = self.tipo_conversion.get()
        filetypes = [("Word", "*.docx")] if tipo == "word_to_pdf" else [("PDF", "*.pdf")]
        archivo = filedialog.askopenfilename(filetypes=filetypes)
        if archivo:
            self.ruta_archivo = archivo
            self.label_archivo.config(text=os.path.basename(archivo), fg="black")
            sugerida = archivo.rsplit(".", 1)[0] + (".pdf" if tipo == "word_to_pdf" else ".docx")
            self.entry_salida.delete(0, "end")
            self.entry_salida.insert(0, sugerida)

    def convertir(self):
        salida = self.entry_salida.get().strip()
        if not self.ruta_archivo:
            messagebox.showerror("Error", "Debe seleccionar un archivo de entrada.")
            return
        if not salida:
            messagebox.showerror("Error", "Debe ingresar una ruta de salida.")
            return

        try:
            if self.tipo_conversion.get() == "word_to_pdf":
                documento = DocumentoWord(self.ruta_archivo)
            else:
                documento = DocumentoPDF(self.ruta_archivo)
            convertidor = Convertidor(documento)
            convertidor.ejecutar_conversion(salida)
            messagebox.showinfo("Éxito", f"Conversión completada:\n{salida}")
        except Exception as e:
            messagebox.showerror("Error", str(e))

# Lanzamiento
if __name__ == "__main__":
    root = Tk()
    root.geometry("400x300")
    app = Aplicacion(root)
    root.mainloop()
