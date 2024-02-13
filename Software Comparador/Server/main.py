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
        tiempoestabilizacion = data.get("tiempoestabilizacion", "Error")
        numRepeticiones = data.get("numRepeticiones", "Error")
        print("Tiempo de Estabilización: ", tiempoestabilizacion)
        print("Número de Repeticiones: ", numRepeticiones)
        output = Centros(tiempoestabilizacion, numRepeticiones)
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

    print(data, instrumento)
    
    output = forzar_obtencion_CA(instrumento)

    print("Datos: ", output)
    return jsonify(output)


 
@app.route('/isUp', methods=['GET'])
def isUp():
    return jsonify(status="online"), 200

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000)