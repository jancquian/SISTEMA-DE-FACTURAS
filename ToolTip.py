"""
Código generado con ChatGPT Open IA.

Consultado por el pasante de Ingeniería en Sistemas Computacionales (ISC)
Juan Carlos Garcia Jimenez (juancarlosgarciajimenez123@gmail.com) para la empresa
Distribución e Ingeniería VAR S.A. de C.V en los años 2023-2024.

El archivo ToolTip.py contiene la clase ToolTip; la cual se encarga de proveer el
funcionamiento de texto desplegable al momento de posicionar el cursor en el
área de un botón.
"""

import tkinter as tk


class ToolTip:
    def __init__(self, widget, text):
        self.widget = widget
        self.text = text
        self.tip_window = None

    # El método “show_tooltip” se encarga de desplegar el texto una vez que el
    # cursor ha entrado en la región del botón.
    def show_tooltip(self, event):
        x, y, _, _ = self.widget.bbox("insert")
        x += self.widget.winfo_rootx() + 25
        y += self.widget.winfo_rooty() + 20

        self.tip_window = tw = tk.Toplevel(self.widget)
        tw.wm_overrideredirect(True)
        tw.wm_geometry("+%d+%d" % (x, y))

        label = tk.Label(tw, text=self.text, justify='left',
                         background="#ffffe0", relief='solid', borderwidth=1)
        label.pack(ipadx=1)

    # El método “hide_tooltip” se encarga de destruir el widget del texto una
    # vez que el cursor ha salido del área del botón.
    def hide_tooltip(self, event):
        if self.tip_window:
            self.tip_window.destroy()
        self.tip_window = None


# La función “create_tooltip” genera un objeto de la clase ToolTip y le asocia los
# eventos de “entrada” que ejecuta el método “show_tooltip” y evento de “salida”
# que ejecuta el método “hide_tooltip”.
def create_tooltip(widget, text):
    tool_tip = ToolTip(widget, text)
    widget.bind("<Enter>", tool_tip.show_tooltip)
    widget.bind("<Leave>", tool_tip.hide_tooltip)
