from flask import Flask, jsonify, request
import openpyxl

# ----------- ¿Qué hace la librería openpyxl? -----------
# La librería openpyxl se utiliza para leer y escribir en archivos Excel (.xlsx).
# Permite abrir un archivo Excel, manipular sus hojas y guardar cambios, como si fuera una base de datos.
# En este caso, usamos Excel para almacenar los datos de los usuarios en lugar de una base de datos tradicional.

# Crear una aplicación Flask
app = Flask(__name__)

# Imprimir mensaje en consola al iniciar el servidor
# Este mensaje se mostrará cuando inicies el servidor. Sirve como presentación del ejercicio.
print("Duoc UC, Ejemplo de API básica con Excel")
print("Profesor: Felipe Robinet")

# Ruta para obtener todos los datos del archivo Excel (GET /usuarios)
@app.route('/usuarios', methods=['GET'])
def obtener_usuarios():
    # Crear una lista vacía para almacenar los usuarios que leemos del archivo Excel
    usuarios = []

    # Abrir el archivo Excel existente 'datos.xlsx'
    # El método load_workbook abre el archivo Excel para que podamos leer su contenido.
    libro = openpyxl.load_workbook('datos.xlsx')

    # Seleccionamos la primera hoja activa del archivo
    hoja = libro.active

    # Iteramos sobre cada fila de la hoja de cálculo.
    # values_only=True nos permite obtener solo los valores (sin formato ni fórmulas).
    for fila in hoja.iter_rows(values_only=True):
        id, nombre, email = fila  # Se asignan las columnas a variables
        # Añadimos los datos a la lista de usuarios
        usuarios.append({'id': id, 'nombre': nombre, 'email': email})
    
    # Convertimos la lista de usuarios a formato JSON para enviar como respuesta
    return jsonify(usuarios)

# Ruta para agregar un nuevo usuario al archivo Excel (POST /usuarios)
@app.route('/usuarios', methods=['POST'])
def agregar_usuario():
    # Obtenemos los datos del nuevo usuario desde la solicitud (en formato JSON)
    nuevo_usuario = request.json

    # Abrimos el archivo Excel para escribir en él
    libro = openpyxl.load_workbook('datos.xlsx')
    hoja = libro.active

    # ----------- Verificación de duplicados -----------
    # Antes de agregar un usuario nuevo, verificamos si ya existe un usuario con el mismo ID
    for fila in hoja.iter_rows(min_row=2, values_only=True):
        id_existente = str(fila[0])  # Obtenemos el ID de la fila actual
        # Si el ID ya existe, devolvemos un mensaje de error
        if id_existente == str(nuevo_usuario['id']):
            return jsonify({'mensaje': 'Error: El usuario con ID ya existe'}), 400

    # ----------- Agregar el usuario al archivo Excel -----------
    # Si el ID no existe, agregamos el nuevo usuario al final de la hoja
    hoja.append([nuevo_usuario['id'], nuevo_usuario['nombre'], nuevo_usuario['email']])

    # Guardamos los cambios en el archivo Excel
    libro.save('datos.xlsx')

    # Enviamos una respuesta indicando que el usuario fue agregado exitosamente
    return jsonify({'mensaje': 'Usuario agregado'}), 201

# Ruta para obtener un usuario específico por ID desde Excel (GET /usuarios/<id>)
@app.route('/usuarios/<id>', methods=['GET'])
def obtener_usuario(id):
    # Abrimos el archivo Excel para leer los datos
    libro = openpyxl.load_workbook('datos.xlsx')
    hoja = libro.active

    # Recorremos todas las filas de la hoja para buscar al usuario con el ID solicitado
    for fila in hoja.iter_rows(values_only=True):
        id_usuario, nombre, email = fila
        # Si encontramos el ID, devolvemos la información del usuario en formato JSON
        if str(id_usuario) == id:
            return jsonify({'id': id_usuario, 'nombre': nombre, 'email': email})
    
    # Si no encontramos el usuario, devolvemos un mensaje de error
    return jsonify({'mensaje': 'Usuario no encontrado'}), 404

# Ruta para actualizar un usuario existente en el archivo Excel (PUT /usuarios/<id>)
@app.route('/usuarios/<id>', methods=['PUT'])
def actualizar_usuario(id):
    # Obtenemos los datos actualizados desde la solicitud en formato JSON
    datos_actualizados = request.json

    # Abrimos el archivo Excel para modificarlo
    libro = openpyxl.load_workbook('datos.xlsx')
    hoja = libro.active

    # Variable que indica si el usuario fue encontrado y actualizado
    actualizado = False

    # Recorremos cada fila de la hoja desde la segunda fila (min_row=2)
    # Buscamos el usuario con el ID solicitado
    for fila in hoja.iter_rows(min_row=2, values_only=False):
        id_usuario = fila[0].value  # Obtenemos el ID de la fila
        if str(id_usuario) == id:
            # Actualizamos los valores de nombre y email si se han proporcionado nuevos valores
            fila[1].value = datos_actualizados.get('nombre', fila[1].value)
            fila[2].value = datos_actualizados.get('email', fila[2].value)
            actualizado = True
            break  # Si actualizamos el usuario, salimos del bucle

    # Si el usuario fue actualizado, guardamos los cambios en el archivo Excel
    if actualizado:
        libro.save('datos.xlsx')
        return jsonify({'mensaje': 'Usuario actualizado'}), 200
    else:
        # Si no encontramos el ID, devolvemos un mensaje de error
        return jsonify({'mensaje': 'Usuario no encontrado'}), 404

# ----------- Ejecutar la aplicación -----------
# Esta línea inicia el servidor Flask. El servidor se ejecutará en modo de depuración (debug=True),
# lo que permite ver errores y cambios automáticamente.
if __name__ == '__main__':
    app.run(debug=True)
