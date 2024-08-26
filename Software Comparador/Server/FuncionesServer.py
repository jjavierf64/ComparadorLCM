#! /usr/bin/env python

"""
Funciones para el funcionamiento correcto del servidor
"""



################## Importación de librerías ##################
import RPi.GPIO as GPIO                                             # Biblioteca para el control de los motores a pasos y el servomotor
from RpiMotorLib import RpiMotorLib                                 # Biblioteca para motores a pasos
from RpiMotorLib import rpiservolib                                 # Biblioteca para servomotor
from time import sleep                                              # Biblioteca para sleep
import time
import serial
from decimal import Decimal                                         # Biblioteca para trabajar correctamente operaciones aritméticas con flotantes decimales
import curses														# Biblioteca para interacción con el teclado
import os                                                           # Biblioteca para interactuar con el sistema operativo
import warnings
import csv                                                          


################################################################

global motorEnabledState
global motorDisabledState 
motorEnabledState = GPIO.HIGH
motorDisabledState = GPIO.LOW

################## Configuración de entradas/salidas ##################

GPIO_pins1 = (22, 27, 17)           #pines de modo para el motor1
direction1 = 9                      #pin de dirección para el motor1
step1 = 11                          #pin de step para el motor1

GPIO_pins2 = (5, 6, 13)             #pines de modo para el motor2
direction2 = 20                     #pin de dirección para el motor2
step2 = 21                          #pin de step para el motor2

pin_enableCalibrationMotor = 24                         #pin de enable

GPIO_pins3 = (14, 15, 18)           #Pines de modo de paso
direction3 = 19                     #Pin de sentido de giro
step3 = 16                          #Pin de dar paso
pin_enablePlateMotor = 23


pin_startRotationLimitSensor = 4               #Pin para el sensor infrarrojo de rotacion de angulo nicial
pin_endRotationLimitSensor = 3                 #Pin para el sensor infrarrojo de rotacion de angulo final

steperMotorPlate = RpiMotorLib.A4988Nema(direction3, step3, GPIO_pins3, "A4988") #Parámetros del motor

steperMotor1 = RpiMotorLib.A4988Nema(direction1, step1, GPIO_pins1, "A4988") #Parámetros del motor1
steperMotor2 = RpiMotorLib.A4988Nema(direction2, step2, GPIO_pins2, "A4988") #Parámetros del motor2

GPIO.setup(pin_enableCalibrationMotor, GPIO.OUT)     
                                                                                                                                           
GPIO.output(pin_enableCalibrationMotor, motorDisabledState)       #Modo seguro, motores inhabilitados

GPIO.setup(pin_enablePlateMotor, GPIO.OUT)     
                                                                                                                                           
GPIO.output(pin_enablePlateMotor, motorDisabledState)       #Modo seguro, motores de plato inhabilitados


GPIO.setmode(GPIO.BCM)              #Numeración Broadcom
GPIO.setup(pin_startRotationLimitSensor, GPIO.IN)    #Se define como entrada el sensor
GPIO.setup(pin_endRotationLimitSensor, GPIO.IN)    #Se define como entrada el sensor

######## Longitudes de Plantillas en Steps
#### Plantilla1
# Motor 1
## Todos están configurados para 1/8 de step
global stepsP1_12  # De punto 1 a 2
global stepsP1_24v # De punto 2 a 4 en dirección vertical
global stepsP1_43v # De punto 4 a 3 en dirección vertical
stepsP1_12 = 205
stepsP1_24v = 40
stepsP1_43v = 80

# Motor 2
## Configurado para Full de step
global stepsP1_23h # De punto 2 a 3 en dirección horizontal
stepsP1_23h = 340


###
#### Plantilla2
# Motor 2
## Configurado para Full de step
global stepsP2_23h # De punto 2 a 3 en dirección horizontal
stepsP2_23h = 392






################## Captura de datos TESA ##################

serTESA=serial.Serial("/dev/ttyUSBI", baudrate=1200, bytesize=serial.SEVENBITS, parity=serial.PARITY_EVEN,
                          stopbits=serial.STOPBITS_TWO, xonxoff=True, timeout=0.5) #Configuración de puerto

def DatosTESA():                                   
    detenerse=0                     #Constante para while que captura dato
    def recv(serial):               #Definición de una función para recibir datos
        while True:
            
            data=serial.read(30)    #Lectura de 30 bytes
            if data == "":
                continue
            else:
                break
            sleep(0.02)
        return data
    while detenerse == 0:
        data=recv(serTESA)          #Llamada de la función
        if data != b"":             #Comparación de datos recibidos, vacío hasta que se de la medición
            try:
                medicion=float(data) #Pasando de string a float
                MedicionBloque=medicion #Guardando dato en lista
            
            except:
                divisionDatos=data.split()
                medicion=float(divisionDatos[1])    #Pasando de string a decimal
                MedicionBloque=medicion #Guardando dato en lista
            detenerse = 1           #Condición para salir del while
    
    return MedicionBloque






################## Captura de datos Fluke ##################

serFluke=serial.Serial("/dev/ttyUSBK", baudrate=9600, bytesize=serial.EIGHTBITS, parity=serial.PARITY_NONE,
                       stopbits=serial.STOPBITS_ONE, xonxoff=True, timeout=0.5) #Configuración de puerto

def DatosFluke():
    
    serFluke.flushInput()

    detenerse=0 #Constante para while que captura dato
    
    MedicionTemp1=0 #Creación de variable para almacenar mediciones de temperatura 1
    MedicionTemp2=0 #Creación de variable para almacenar mediciones de temperatura 2
    MedicionTemp3=0 #Creación de variable para almacenar mediciones de temperatura 3
    MedicionTemp4=0 #Creación de variable para almacenar mediciones de temperatura 4


    def recv(serial): #Definición de una función para recibir datos
        while True:
            data=serial.read(34) #Lectura de 32 bytes
            if data == "":
                continue
            else:
                break
            sleep(0.02)
        return data


    while detenerse <= 3:
        data=recv(serFluke) #Llamada de la función

        if data != b"": #Comparación de datos recibidos, vacío hasta que se de la medición
            medicion=data.split()#Separar los 4 datos en una lista
            print(medicion)
            if int(medicion[0]) == 1 and MedicionTemp1==0:
                MedicionTemp1=float(medicion[1]) #Guardando temperatura 2
                detenerse += 1  #Condición para salir del while
            elif int(medicion[0]) == 2 and MedicionTemp2==0:
                MedicionTemp2=float(medicion[1]) #Guardando temperatura 2
                detenerse += 1  #Condición para salir del while
            elif int(medicion[0]) == 3 and MedicionTemp3==0:
                MedicionTemp3=float(medicion[1]) #Guardando temperatura 3
                detenerse += 1  #Condición para salir del while
            elif int(medicion[0]) == 4 and MedicionTemp4==0:
                MedicionTemp4=float(medicion[1]) #Guardando temperatura 4
                detenerse += 1  #Condición para salir del while

    return MedicionTemp1, MedicionTemp2, MedicionTemp3, MedicionTemp4




################## Captura de datos Vaisala ##################

serVaisala=serial.Serial("/dev/ttyUSBD", baudrate=4800, bytesize=serial.SEVENBITS,
                             parity=serial.PARITY_EVEN, stopbits=serial.STOPBITS_ONE, timeout= 0.5) #Configuración de puerto

def DatosVaisala():
    
    serVaisala.write(b'FORM 4.2 " P=" P " " U6 3.2 "T=" T " " U3 3.2 "RH=" RH " " U4\r\n') # Formato para la toma de datos
    serVaisala.write(b"SEND\r\n") #Envío de instrucción para capturar datos del Vaisala

    
    detenerse=0 #Constante para while que captura dato
    
    DatoPresVaisala=0 #Creación de variable para almacenar mediciones de presión atmosférica
    DatoTempVaisala=0 #Creación de variable para almacenar mediciones de temperatura ambiente
    DatoHumeVaisala=0 #Creación de variable para almacenar mediciones de humedad relativa
    
    def recv(serial): #Definición de una función para recibir datos
        while True:
            data=serial.read(85)
            if data == "":
                continue
            else:
                break
            sleep(0.02)
        return data
    
    
    while detenerse == 0:
        data=recv(serVaisala)                   #Llamada de la función
        if data.split()[0] == b'OK': 						#Comparación de datos recibidos, vacío hasta que se de la medición
            todos=data.split()					#Separar los 4 datos en una lista
            
            DatoPresVaisala=float(todos[3]) #Guardando presión atmosférica en lista
            DatoTempVaisala=float(todos[6]) #Guardando temperatura en lista
            DatoHumeVaisala=float(todos[9]) #Guardando humedad relativa 3 en lista
            
            detenerse = 1                    #Condición para salir del while
            
            
    return DatoHumeVaisala





################## Servo Motor ##################

servo_pin = 26 #Pin que envía la señal al servomotor

def ActivaPedal(servo_pin=26): 

    myservotest = rpiservolib.SG90servo("servoone", 50, 2, 12) #Parámetros del servomotor

    myservotest.servo_move(servo_pin, 2.3, .5, False, .01)     #Movimiento a posición 2.3
    myservotest.servo_move(servo_pin, 7.5, .5, False, .01)     #Movimiento a posición 7.5


################## Mover Posición Centro a Posición 1 ########################

def moverDe0a1_():
    ActivaPedal(servo_pin)								#Sube el palpador
    sleep(int(1))        
    GPIO.output(pin_enableCalibrationMotor, motorEnabledState)       #habilita los motores
    steperMotor1.motor_go(False, "1/8", round(stepsP1_12/2), .0025, False, 2) #Mov de 0 a 1
    GPIO.output(pin_enableCalibrationMotor, motorDisabledState)       #Modo seguro, motores inhabilitados               
    ActivaPedal(servo_pin)								#Baja el palpador
    return


################## Mover Plato Giratorio ########################


def moverPlato_(pos):
    def gohome(): #Función para mover el plato a la posición 1 o home
        start = GPIO.input(pin_startRotationLimitSensor)
        clockwise=True  #dirección de giro normal
        while start !=1:
            #Movimiento de posición 1 a 2 con el palpador arriba
            GPIO.output(pin_enablePlateMotor, motorEnabledState)       #habilita los motores
            steperMotorPlate.motor_go(clockwise, "1/16", 16, .00045, False, 0)
            GPIO.output(pin_enablePlateMotor, motorDisabledState)       #Modo seguro, motores inhabilitados
            
            start = GPIO.input(pin_startRotationLimitSensor) #Vuelve a obtener valores de inicio y fin
            end = GPIO.input(pin_endRotationLimitSensor)
            
            if end ==1: #Si se salta el primer sensor y el segundo lo detecta, invierte el giro
                clockwise = False 
        return

    def moveTo(numSteps):
        GPIO.output(pin_enablePlateMotor, motorEnabledState)       #habilita los motores
        steperMotorPlate.motor_go(False, "1/16", numSteps, .00045, False, 1)
        GPIO.output(pin_enablePlateMotor, motorDisabledState)       #Modo seguro, motores inhabilitados
        return

    
    #mover home 
    gohome()
    
    if pos == "2":
        #mover a 2
        moveTo(6_500)
    elif pos == "3":
        #mover a 3
        moveTo(13_000)
    elif pos == "4":
        #mover a 4
        moveTo(19_500)
    
    return




################## Secuencia desviación de longitud central ##################

def Centros(tiempoinicial, tiempoestabilizacion, Repeticiones):
    def tomar_medicion(tiempoestabilizacion): #Subfuncion para tomar medicion
        try:
            sleep(int(tiempoestabilizacion))				    #Se le da un tiempo al palpador abajo en el bloque patrón
            ActivaPedal(servo_pin)								#Sube el palpador
            MedicionBloque=DatosTESA()                   		#Justo después de levantar el palpador TESA toma la medición
        except:
            sleep(5)
            ActivaPedal(servo_pin)
            MedicionBloque = tomar_medicion(tiempoestabilizacion)
        
        return MedicionBloque



	# Tiempo de estabilización entra en segundos
	
    global valorNominalBloque
    global dato
    global t1
    t1=time.time()                                   #finaliza el conteo de espera de bloques
    tic=time.perf_counter()                                 #Toma el tiempo inicial
    
    listaMediciones=[]
    
    #Antes de empezar a medir es necesario que el palpador vuelva a subir un momento sobre el patrón
    ActivaPedal(servo_pin)								#Sube el palpador
    sleep(5)					#Se le da un tiempo al palpador arriba sobre el bloque patrón
    ActivaPedal(servo_pin)								#Baja el palpador  
    
    for i in range(int(Repeticiones)):
		
		#Medición del bloque patrón (inicia con el palpador abajo)
        MedicionBloque = tomar_medicion(tiempoestabilizacion)
        print(MedicionBloque)
        listaMediciones.append(MedicionBloque)              	#Valor del patrón en posición 1 (centro patrón)
        
        #Movimiento de posición 1 a 2 con el palpador arriba
        sleep(1) #Esperar a que suba el palpador
        GPIO.output(pin_enableCalibrationMotor, motorEnabledState)       #habilita los motores
        steperMotor1.motor_go(True, "1/8", stepsP1_12, .0025, False, 2)
        GPIO.output(pin_enableCalibrationMotor, motorDisabledState)       #Modo seguro, motores inhabilitados	
        
        #Medición del calibrando
        ActivaPedal(servo_pin)                              #Baja el palpador
        MedicionBloque=tomar_medicion(tiempoestabilizacion)                 		#Justo después de levantar el palpador TESA toma la medición
        listaMediciones.append(MedicionBloque)               	#Valor del calibrando en posición 2 (centro calibrando)
        print(MedicionBloque)
        
        #Movimiento de 2 a 1 con el palpador arriba
        sleep(1) #Esperar a que suba el palpador
        GPIO.output(pin_enableCalibrationMotor, motorEnabledState)       #habilita los motores
        steperMotor1.motor_go(False, "1/8", stepsP1_12, .0025, False, 2)
        GPIO.output(pin_enableCalibrationMotor, motorDisabledState)       #Modo seguro, motores inhabilitados
        ActivaPedal(servo_pin)								#Baja el palpador (termina cada repetición con el palpador abajo)

    sleep(5)
    ActivaPedal(servo_pin)								#Sube el palpador
    sleep(5)
    GPIO.output(pin_enableCalibrationMotor, motorEnabledState)       #habilita los motores
    steperMotor1.motor_go(True, "1/8", round(stepsP1_12/2), .0025, False, 2) #Mov de 1 a HOME
    GPIO.output(pin_enableCalibrationMotor, motorDisabledState)       #Modo seguro, motores inhabilitados               
    ActivaPedal(servo_pin)								#Baja el palpador
    
    toc=time.perf_counter()                                 #Toma el tiempo final
    global tiempoCorrida
    tiempoCorrida=toc-tic                                   #retorna el tiempo de corrida en segundos
    global t0
    t0=time.time()                                   #inicia el conteo de espera de bloques
    return listaMediciones, tiempoCorrida, t0







################## Secuencia desviación de longitud central + planitud (Plantilla 1) ##################

def Completa1(tiempoinicial, tiempoestabilizacion, Repeticiones):
    def tomar_medicion(tiempoestabilizacion): #Subfuncion para tomar medicion
        try:
            sleep(int(tiempoestabilizacion))				    #Se le da un tiempo al palpador abajo en el bloque patrón
            ActivaPedal(servo_pin)								#Sube el palpador
            MedicionBloque=DatosTESA()                   		#Justo después de levantar el palpador TESA toma la medición
        except:
            sleep(5)
            ActivaPedal(servo_pin)
            MedicionBloque = tomar_medicion(tiempoestabilizacion)
        
        return MedicionBloque


    global valorNominalBloque
    global dato
    global t1
    t1=time.time()                                   #finaliza el conteo de espera de bloques
    tic=time.perf_counter()                                 #Toma el tiempo inicial
    
    listaMediciones=[]
    
    ActivaPedal(servo_pin) #Sube palpador
    sleep(5)
    ActivaPedal(servo_pin) #Baja palpador
    
    for i in range(int(Repeticiones)):
        
        #Medición del bloque patrón (inicia con el palpador abajo)
        MedicionBloque=tomar_medicion(tiempoestabilizacion)
        listaMediciones.append(MedicionBloque) #Valor del patrón en posición 1 (centro)
        print(MedicionBloque)

        #Movimiento de posición 1 a 2 con el palpador arriba
        sleep(1) #Esperar a que suba el palpador
        GPIO.output(pin_enableCalibrationMotor, motorEnabledState)       #habilita los motores
        steperMotor1.motor_go(True, "1/8", stepsP1_12, .005, False, 2) #Mov de 1 a 2
        GPIO.output(pin_enableCalibrationMotor, motorDisabledState)       #Modo seguro, motores inhabilitados

        #Medición del calibrando
        ActivaPedal(servo_pin) #Baja palpador
        MedicionBloque=tomar_medicion(tiempoestabilizacion)
        listaMediciones.append(MedicionBloque)  #Valor del calibrando en posición 2 (esquina)
        print(MedicionBloque)

        sleep(1) #Esperar a que suba el palpador
        GPIO.output(pin_enableCalibrationMotor, motorEnabledState)       #habilita los motores
        steperMotor1.motor_go(False, "1/8", stepsP1_24v, .005, False, 2) #Mov1 de 2 a 3
        steperMotor2.motor_go(False, "Full", stepsP1_23h, .005, False, 1) #Mov2 de 2 a 3
        steperMotor1.motor_go(True, "1/8", stepsP1_43v, .005, False, 1) #Mov3 de 2 a 3
        GPIO.output(pin_enableCalibrationMotor, motorDisabledState)       #Modo seguro, motores inhabilitados

        ActivaPedal(servo_pin) #Baja palpador
        MedicionBloque=tomar_medicion(tiempoestabilizacion)
        listaMediciones.append(MedicionBloque) #Valor del calibrando en posición 3 (esquina)
        print(MedicionBloque)
        
        sleep(1) #Esperar a que suba el palpador
        GPIO.output(pin_enableCalibrationMotor, motorEnabledState)       #habilita los motores            
        steperMotor1.motor_go(False, "1/8", stepsP1_43v, .005, False, 2) #Mov de 3 a 4
        GPIO.output(pin_enableCalibrationMotor, motorDisabledState)       #Modo seguro, motores inhabilitados

        ActivaPedal(servo_pin) #Baja palpador
        MedicionBloque=tomar_medicion(tiempoestabilizacion)
        listaMediciones.append(MedicionBloque)  #Valor del calibrando en posición 4 (esquina)
        print(MedicionBloque)

        sleep(1) #Esperar a que suba el palpador
        GPIO.output(pin_enableCalibrationMotor, motorEnabledState)       #habilita los motores
        steperMotor2.motor_go(True, "Full", stepsP1_23h*2 , .005, False, 2) #Mov de 4 a 5
        GPIO.output(pin_enableCalibrationMotor, motorDisabledState)       #Modo seguro, motores inhabilitados

        ActivaPedal(servo_pin) #Baja palpador
        MedicionBloque=tomar_medicion(tiempoestabilizacion)
        listaMediciones.append(MedicionBloque)  #Valor del calibrando en posición 5 (esquina)
        print(MedicionBloque)

        sleep(1) #Esperar a que suba el palpador
        GPIO.output(pin_enableCalibrationMotor, motorEnabledState)       #habilita los motores
        steperMotor1.motor_go(True, "1/8", stepsP1_43v, .005, False, 2) #Mov de 5 a 6
        GPIO.output(pin_enableCalibrationMotor, motorDisabledState)       #Modo seguro, motores inhabilitados

        ActivaPedal(servo_pin) #Baja palpador
        MedicionBloque=tomar_medicion(tiempoestabilizacion)
        listaMediciones.append(MedicionBloque)  #Valor del calibrando en posición 6 (esquina)
        print(MedicionBloque)

        sleep(1) #Esperar a que suba el palpador
        GPIO.output(pin_enableCalibrationMotor, motorEnabledState)       #habilita los motores
        steperMotor1.motor_go(False, "1/8", stepsP1_43v, .005, False, 2) #Mov de 6 a 5
        steperMotor2.motor_go(False, "Full", stepsP1_23h, .005, False, 1) #Mov de 5 a Esp2
        steperMotor1.motor_go(False, "1/8", stepsP1_12-stepsP1_24v, .005, False, 1) #Mov de Esp2 a 1
        GPIO.output(pin_enableCalibrationMotor, motorDisabledState)       #Modo seguro, motores inhabilitados

        ActivaPedal(servo_pin) #Baja palpador
        MedicionBloque=tomar_medicion(tiempoestabilizacion)
        listaMediciones.append(MedicionBloque)  #Valor del patrón en posición 1 (centro)
        print(MedicionBloque)
        sleep(5)
        ActivaPedal(servo_pin) #Baja palpador
        

    
    sleep(5)
    ActivaPedal(servo_pin) #Sube palpador
    sleep(5)
    GPIO.output(pin_enableCalibrationMotor, motorEnabledState)       #habilita los motores
    steperMotor1.motor_go(True, "1/8", round(stepsP1_12/2), .005, False, 2) #Mov de 1 a HOME
    GPIO.output(pin_enableCalibrationMotor, motorDisabledState)       #Modo seguro, motores inhabilitados

    ActivaPedal(servo_pin) #Baja palpador

    toc=time.perf_counter()                                 #Toma el tiempo final
    global tiempoCorrida
    tiempoCorrida=toc-tic                            #retorna el tiempo de corrida en segundos
    global t0
    t0=time.time()                                   #inicia el conteo de espera de bloques
    return listaMediciones, tiempoCorrida, t0









################## Secuencia desviación de longitud central + planitud (Plantilla 2) ##################

def Completa2(tiempoinicial, tiempoestabilizacion, Repeticiones):
    def tomar_medicion(tiempoestabilizacion): #Subfuncion para tomar medicion
        try:
            sleep(int(tiempoestabilizacion))				    #Se le da un tiempo al palpador abajo en el bloque patrón
            ActivaPedal(servo_pin)								#Sube el palpador
            MedicionBloque=DatosTESA()                   		#Justo después de levantar el palpador TESA toma la medición
        except:
            sleep(5)
            ActivaPedal(servo_pin)
            MedicionBloque = tomar_medicion(tiempoestabilizacion)
        
        return MedicionBloque


    global valorNominalBloque
    global dato    
    global t1
    t1=time.time()                                   #finaliza el conteo de espera de bloques
    tic=time.perf_counter()                                 #Toma el tiempo inicial
    
    listaMediciones=[]
    ActivaPedal(servo_pin) #Sube palpador
    sleep(5)
    ActivaPedal(servo_pin) #Baja palpador
    
    for i in range(int(Repeticiones)):

        MedicionBloque = tomar_medicion(tiempoestabilizacion) #Tomar medición
        listaMediciones.append(MedicionBloque) #Valor del patrón en posición 1 (centro)
        print(MedicionBloque)

        sleep(1)
        GPIO.output(pin_enableCalibrationMotor, motorEnabledState)       #habilita los motores
        steperMotor1.motor_go(True, "1/8", stepsP1_12, .005, False, 2) #Mov de 1 a 2
        GPIO.output(pin_enableCalibrationMotor, motorDisabledState)       #Modo seguro, motores inhabilitados

        ActivaPedal(servo_pin) #Baja palpador
        MedicionBloque = tomar_medicion(tiempoestabilizacion) #Tomar medición
        listaMediciones.append(MedicionBloque)  #Valor del calibrando en posición 2 (esquina)
        print(MedicionBloque)
        
        sleep(1)
        GPIO.output(pin_enableCalibrationMotor, motorEnabledState)       #habilita los motores
        steperMotor1.motor_go(False, "1/8", stepsP1_24v, .005, False, 2) #Mov1 de 2 a 3
        steperMotor2.motor_go(False, "Full", stepsP2_23h, .005, False, 1) #Mov2 de 2 a 3
        steperMotor1.motor_go(True, "1/8", stepsP1_43v, .005, False, 1) #Mov3 de 2 a 3
        GPIO.output(pin_enableCalibrationMotor, motorDisabledState)       #Modo seguro, motores inhabilitados

        ActivaPedal(servo_pin) #Baja palpador
        MedicionBloque = tomar_medicion(tiempoestabilizacion) #Tomar medición
        listaMediciones.append(MedicionBloque) #Valor del calibrando en posición 3 (esquina)
        print(MedicionBloque)
        
        sleep(1)
        GPIO.output(pin_enableCalibrationMotor, motorEnabledState)       #habilita los motores            
        steperMotor1.motor_go(False, "1/8", stepsP1_43v, .005, False, 2) #Mov de 3 a 4
        GPIO.output(pin_enableCalibrationMotor, motorDisabledState)       #Modo seguro, motores inhabilitados

        ActivaPedal(servo_pin) #Baja palpador
        MedicionBloque = tomar_medicion(tiempoestabilizacion) #Tomar medición
        listaMediciones.append(MedicionBloque)  #Valor del calibrando en posición 4 (esquina)
        print(MedicionBloque)
        
        sleep(1)
        GPIO.output(pin_enableCalibrationMotor, motorEnabledState)       #habilita los motores
        steperMotor2.motor_go(True, "Full", stepsP2_23h*2 , .005, False, 2) #Mov de 4 a 5
        GPIO.output(pin_enableCalibrationMotor, motorDisabledState)       #Modo seguro, motores inhabilitados

        ActivaPedal(servo_pin) #Baja palpador
        MedicionBloque = tomar_medicion(tiempoestabilizacion) #Tomar medición
        listaMediciones.append(MedicionBloque)  #Valor del calibrando en posición 5 (esquina)
        print(MedicionBloque)
        
        sleep(1)
        GPIO.output(pin_enableCalibrationMotor, motorEnabledState)       #habilita los motores
        steperMotor1.motor_go(True, "1/8", stepsP1_43v, .005, False, 2) #Mov de 5 a 6
        GPIO.output(pin_enableCalibrationMotor, motorDisabledState)       #Modo seguro, motores inhabilitados

        ActivaPedal(servo_pin) #Baja palpador
        MedicionBloque = tomar_medicion(tiempoestabilizacion) #Tomar medición
        listaMediciones.append(MedicionBloque)  #Valor del calibrando en posición 6 (esquina)
        print(MedicionBloque)
        
        sleep(1)
        GPIO.output(pin_enableCalibrationMotor, motorEnabledState)       #habilita los motores
        steperMotor1.motor_go(False, "1/8", stepsP1_43v, .005, False, 2) #Mov de 6 a 5
        steperMotor2.motor_go(False, "Full", stepsP2_23h, .005, False, 1) #Mov de 5 a Esp2
        steperMotor1.motor_go(False, "1/8", stepsP1_12-stepsP1_24v, .005, False, 1) #Mov de Esp2 a 1
        GPIO.output(pin_enableCalibrationMotor, motorDisabledState)       #Modo seguro, motores inhabilitados

        ActivaPedal(servo_pin) #Baja palpador
        MedicionBloque = tomar_medicion(tiempoestabilizacion) #Tomar medición
        listaMediciones.append(MedicionBloque)  #Valor del patrón en posición 1 (centro)
        print(MedicionBloque)
        sleep(5)
        ActivaPedal(servo_pin) #Baja palpador
        

    sleep(5)
    ActivaPedal(servo_pin) #Sube palpador
    sleep(5)
    GPIO.output(pin_enableCalibrationMotor, motorEnabledState)       #habilita los motores
    steperMotor1.motor_go(True, "1/8", round(stepsP1_12/2), .005, False, 2) #Mov de 1 a HOME
    GPIO.output(pin_enableCalibrationMotor, motorDisabledState)       #Modo seguro, motores inhabilitados

    ActivaPedal(servo_pin) #Baja palpador
    
    toc=time.perf_counter()                                 #Toma el tiempo final
    global tiempoCorrida
    tiempoCorrida=toc-tic                            #retorna el tiempo de corrida en segundos
    global t0
    t0=time.time()                                   #inicia el conteo de espera de bloques
    return listaMediciones, tiempoCorrida, t0






################## Creación de un archivo csv para Datos ##################

def CrearArchivoCSV(seleccionSecuencia, numCertificado):
	# Se crea un archivo csv, nombrado con una marca temporal:
	archivoDatos = "./Calibraciones en curso/" + numCertificado + ".csv" # Nombre del archivo para el almacenaje de datos
	open(archivoDatos, mode="w", newline="")	#Creación del Archivo

    # Se crean también para el registro de condiciones ambientales

	archivoDatosAmbientales = "./Calibraciones en curso/" + numCertificado + "-Ambientales.csv" # Nombre del archivo para el almacenaje de datos
	open(archivoDatosAmbientales, mode="w", newline="")	#Creación del Archivo


	return archivoDatos,archivoDatosAmbientales




################## Prueba de Movimientos de motores ##################

def SecuenciaPrueba():                                                                                    
    GPIO.output(pin_enableCalibrationMotor, motorDisabledState)       #Modo seguro, motores inhabilitados
    GPIO.output(pin_enableCalibrationMotor, motorEnabledState)
    steperMotor1.motor_go(True, "Half", 40, 0.001, True, 0)

    steperMotor2.motor_go(False, "Full", 200, 0.001, True, 0)
    sleep(1)
    steperMotor2.motor_go(True, "Full", 400, 0.001, True, 0)
    steperMotor2.motor_go(False, "Full", 200, 0.001, True, 0)
    sleep(1)
    steperMotor1.motor_go(False, "Half", 40, 0.001, True, 0)
    GPIO.output(pin_enableCalibrationMotor, motorDisabledState)

