################## Importación de librerías ##################
import RPi.GPIO as GPIO                                             # Biblioteca para el control de los motores a pasos y el servomotor
from RpiMotorLib import RpiMotorLib                                 # Biblioteca para motores a pasos
from RpiMotorLib import rpiservolib                                 # Biblioteca para servomotor
from time import sleep                                              # Biblioteca para sleep
import time
import serial                                                       # Biblioteca para configuración y adquisición de 
import warnings

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

sleepMot3 = 12                        #Pin para controlar el sleep del motor de ordenamiento
                                    #Si está en 1 está activo, en 0 está en sleep
pin_startRotationLimitSensor = 4               #Pin para el sensor infrarrojo de rotacion de angulo nicial
pin_endRotationLimitSensor = 3                 #Pin para el sensor infrarrojo de rotacion de angulo final

steperMotorPlate = RpiMotorLib.A4988Nema(direction3, step3, GPIO_pins3, "A4988") #Parámetros del motor

steperMotor1 = RpiMotorLib.A4988Nema(direction1, step1, GPIO_pins1, "A4988") #Parámetros del motor1
steperMotor2 = RpiMotorLib.A4988Nema(direction2, step2, GPIO_pins2, "A4988") #Parámetros del motor2

GPIO.setup(pin_enableCalibrationMotor, GPIO.OUT)     
                                                                                                                                           
GPIO.output(pin_enableCalibrationMotor, motorDisabledState)       #Modo seguro, motores inhabilitados

GPIO.setup(pin_enablePlateMotor, GPIO.OUT)     
                                                                                                                                           
GPIO.output(pin_enablePlateMotor, motorDisabledState)       #Modo seguro, motores de plato inhabilitados

GPIO.setup(sleepMot3, GPIO.OUT)                                                                                                                                                
GPIO.output(sleepMot3, GPIO.LOW)       #Sleep debe estar en LOW para deshabilitarse


GPIO.setmode(GPIO.BCM)              #Numeración Broadcom
GPIO.setup(pin_startRotationLimitSensor, GPIO.IN)    #Se define como entrada el sensor

posicionStep=0                      #Variable de posición angular del disco
required=0                          #Variable de pasos requeridos par llegar
                                    #a la posicion deseada
listo=0                             #Variable que determina cuando terminó

################## Movimiento Manual de los motores ##################

def moverManualInterfaz(event):
    try:
        GPIO.output(pin_enableCalibrationMotor, motorEnabledState)
        tecla = event.keysym
        print(tecla)
        
        
        if tecla == "Up":
            print("Motor 1 go up")
            steperMotor1.motor_go(False, "1/8", 2, 0.005, False, 0)		#COMFIGURACION CORRECTA

        elif tecla == "Down":
            print("Motor 1 go down")
            steperMotor1.motor_go(True, "1/8", 2, 0.005, False, 0)

        elif tecla == "Left":
            steperMotor2.motor_go(False, "Full", 4, 0.001, False, 0)		

        elif tecla == "Right":
            steperMotor2.motor_go(True, "Full", 4, 0.001, False, 0)

        GPIO.output(pin_enableCalibrationMotor, motorDisabledState)   # Apagar motores
    except:
        GPIO.output(pin_enableCalibrationMotor, motorDisabledState)   # Apagar motores
        

################## Servo Motor ##################

servo_pin = 26 #Pin que envía la señal al servomotor

def ActivaPedal(servo_pin=26): 

    myservotest = rpiservolib.SG90servo("servoone", 50, 2, 12) #Parámetros del servomotor

    myservotest.servo_move(servo_pin, 2.3, .5, False, .01)     #Movimiento a posición 2.3
    myservotest.servo_move(servo_pin, 7.5, .5, False, .01)     #Movimiento a posición 7.5


    
################## Ocultar advertencias en terminal ##################

def fxn():
    warnings.warn("deprecated", DeprecationWarning)

with warnings.catch_warnings():
    warnings.simplefilter("ignore")
    fxn()
