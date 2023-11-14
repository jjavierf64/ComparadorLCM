from flask import Flask, jsonify, request
#from FuncionesServer import *

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


@app.route('/secuencias', methods=['POST', 'GET'])
def secuencias():
    # Obtiene los datos en formato JSON de la petición
    data = request.json

    # Supongamos que envías un valor llamado "parametro" en tu JSON
    parametro = data.get('secuencia', '0')
    print("Secuencia: ", parametro)

    if parametro == "Centros":
        print("Tiempo de Estabilización: ", data.get("tiempoestabilizacion", "Error"))
        print("Número de Repeticiones: ", data.get("numRepeticiones", "Error"))
        return jsonify(success=True, received_param=parametro), 200


 
@app.route('/isUp', methods=['GET'])
def isUp():
    return jsonify(status="online"), 200

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000)