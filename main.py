"""
El presente código “main.py” fue elaborado por el pasante de Ingeniería en Sistemas Computacionales (ISC)
Juan Carlos Garcia Jimenez (juancarlosgarciajimenez123@gmail.com) para la empresa
Ingeniería y Distribución VAR S.A. de C.V en los años 2023-2024.
"""
import BookWorm as Worm
import VPrincipal as Vprin
import tkinter as tk
# import psycopg2
# import sys

if __name__ == '__main__':

    worm = Worm.BookWorm("Facturacion.db")
    worm.crecon_db()
    # worm.clasificar_facturas("Facturas Descargadas", "Facturas Clasificadas")

    root = tk.Tk()

    Vprin.VPrincipal(root, worm)

    root.mainloop()

