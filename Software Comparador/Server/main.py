from flask import Flask, jsonify, request
from FuncionesServer import *

app = Flask(__name__)

@app.route('/ejecutar_script', methods=['POST'])
def ejecutar_script():
    # Obtiene los datos en formato JSON de la petición
    data = request.json

    # Supongamos que envías un valor llamado "parametro" en tu JSON
    parametro = data.get('parametro', 'valor_por_defecto')

    # Lógica para ejecutar tu script usando el parámetro
    os.system(f"python3 /home/microv3/LCM/Pruebas/mi_script.py {parametro}")

    # Retorna una respuesta
    return jsonify(success=True, received_param=parametro), 200

@app.route('/moverDe0a1', methods=['POST'])
def moverDe0a1():
    moverDe0a1_()
    # Retorna una respuesta
    return jsonify(success=True), 200

@app.route('/moverPlato', methods=['POST'])
def moverPlato():
    data = request.json
    pos = data.get('posición', '0')
    if pos != '0':
        moverPlato_(pos)
    # Retorna una respuesta
        return jsonify(success=True), 200
    else:
        return jsonify(success=False, received_param=pos), 500

@app.route('/secuencias', methods=['POST', 'GET'])  # Ruta para la petición de ejecución de comandos
def secuencias():
    # Obtiene los datos en formato JSON de la petición
    data = request.json

    # Supongamos que envías un valor llamado "parametro" en tu JSON
    parametro = data.get('secuencia', '0')
    parametro = str(parametro).lower()
    print("Secuencia: ", parametro)

    # Distintos tipos de secuencias posibles
    if parametro == "desviación central":

        tiempoinicial = data.get("tiempoinicial", "Error")
        tiempoestabilizacion = data.get("tiempoestabilizacion", "Error")
        numRepeticiones = data.get("numRepeticiones", "Error")
        print("Tiempo de Estabilización: ", tiempoestabilizacion)
        print("Número de Repeticiones: ", numRepeticiones)
        
        output = Centros(tiempoinicial, tiempoestabilizacion, numRepeticiones)
        
        return jsonify(output)

    if parametro == "desviación central y planitud":
        
        tiempoinicial = data.get("tiempoinicial", "Error")
        tiempoestabilizacion = data.get("tiempoestabilizacion", "Error")
        numRepeticiones = data.get("numRepeticiones", "Error")
        print("Tiempo de Estabilización: ", tiempoestabilizacion)
        print("Número de Repeticiones: ", numRepeticiones)

        plantilla = data.get('plantilla','0')
        plantilla = str(plantilla).lower()

        if plantilla == 'pequeña':
            output = Completa1(tiempoestabilizacion, numRepeticiones)
        elif plantilla == 'grande':
            output = Completa2(tiempoestabilizacion, numRepeticiones)
        else:
            print("Error, parámetro de plantilla no recibido")
            return jsonify(success=False, received_param=parametro), 500

        
        return jsonify(output)
    
    if parametro == "prueba":
        print("Prueba de Motores")
        SecuenciaPrueba()
        return jsonify(success=True, received_param=parametro), 200

    else:
        print("Error, parámetro no recibido")
        return jsonify(success=False, received_param=parametro), 500


@app.route('/condicionesAmbientales', methods=['POST', 'GET'])
def condicionesAmbientales():
    # Función para forzar obtención de Condiciones Ambientales sin que falle el código
    def forzar_obtencion_CA(instrumento):
        try:
            if instrumento == "fluke":
                outputForzado = DatosFluke()

            elif instrumento == "vaisala":
                outputForzado = DatosVaisala()
            else:
                outputForzado = 0

        except:
            outputForzado = forzar_obtencion_CA(instrumento)
        
        return outputForzado

    print("Recibí Petición de Condiciones Ambientales")
    # Obtiene los datos en formato JSON de la petición
    data = request.json

    # Supongamos que envías un valor llamado "parametro" en tu JSON
    instrumento = data.get('instrumento', '0')
    instrumento = str(instrumento).lower()
    
    output = forzar_obtencion_CA(instrumento)

    print("Datos: ", output)
    return jsonify(output)


 
@app.route('/isUp', methods=['GET'])
def isUp():
    return jsonify(status="online"), 200

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000)