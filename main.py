import BookWorm as Worm
import VPrincipal as Vprin
import tkinter as tk
import psycopg2
import sys
# Press the green button in the gutter to run the script.
if __name__ == '__main__':

    '''
    contador = 0
    while True:
        try:
            conexion_pg = psycopg2.connect(user="postgres",
                                           password="",
                                           host="192.168.1.70",
                                           port=5432,
                                           dbname="divar")
            break
        except:
            print("Esperando conexion, intento {0} de 50...".format(contador + 1))
            contador = contador + 1
            if contador == 49:
                print("No fue posible conectar a la base de datos... llame al ddministrador del sistema.")
                sys.exit()
    '''

    worm = Worm.BookWorm("Facturacion.db")
    worm.crecon_db()
    # worm.clasificar_facturas("Facturas Descargadas", "Facturas Clasificadas")

    root = tk.Tk()

    Vprin.VPrincipal(root, worm)

    root.mainloop()

    # worm.clasificar_facturas("Facturas Descargadas", "Facturas Clasificadas")

    # worm.consulta_sql('FACTURAS', 'NOMBRE', 'GAS')
    # worm.imprimir_consulta()

    # worm.descomprimir_zip("prueba.zip", "Facturas Descargadas")
