"""
En este archivo se presenta el código de una interfaz simple para posicionar motores en el punto inicial del TESA.
"""


################## Importación de librerías ##################
import tkinter as tk
from tkinter import ttk
# Definiciones de movimiento de motores y funciones
import FuncionesComparadorInterfazSimple
from FuncionesComparadorInterfazSimple import moverManualInterfaz, ActivaPedal




root = tk.Tk()
root.title("Comparador de bloques TESA")
root.configure(bg="white")

main_label = ttk.Label(root,text="Utilize las flechas del teclado para colocar los motores en la posición inicial.",anchor=tk.CENTER, background="white")
main_label.grid(row=0, column=0, pady=(30, 0), padx=30)

flechas_label = ttk.Label(root, text = "←↕→", anchor = tk.CENTER, background = "white")
flechas_label.grid(row = 1, column = 0, pady = (10, 10))

exit_label = ttk.Label(root, text = "Presione Enter ↲ para salir.", anchor = tk.CENTER,
                       background = "white")
exit_label.grid(row = 2, column = 0, pady = (0, 50), padx = 30)

print("Preparacion Pedal")
ActivaPedal()
sleep(3)
#GPIO.output(pin_enableCalibrationMotor, motorEnabledState)  # Encender motores


def funcionMotores(event):
    print(event.keysym)
    print(type(event.keysym))
    moverManualInterfaz(event)


def muere(event):
    print("Terminacion Pedal")
    ActivaPedal()
    #GPIO.output(pin_enableCalibrationMotor, motorDisabledState)  # Apagar motores (Tal vez no es necesario)
    root.destroy()


root.bind("<Up>", funcionMotores)
root.bind("<Down>", funcionMotores)
root.bind("<Left>", funcionMotores)
root.bind("<Right>", funcionMotores)
root.bind("<Return>", muere)
root.bind("<q>", muere)










root.mainloop()