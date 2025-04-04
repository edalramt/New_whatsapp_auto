import pandas as pd
import pywhatkit as kit
import time
import os

# Leer la lista de preinscritos
archivo_preinscritos = "preinscritos.xlsx"  # Nombre del archivo de entrada
df = pd.read_excel(archivo_preinscritos)

# Verificar si el archivo de "Contactos_enviados.xlsx" existe
archivo_enviados = "Contactos_enviados.xlsx"
if os.path.exists(archivo_enviados):
    df_enviados = pd.read_excel(archivo_enviados)
else:
    df_enviados = pd.DataFrame(columns=df.columns.tolist() + ["Revision"])

# Función para validar número de celular (Perú: 9 dígitos y empieza con 9)
def validar_numero(numero):
    if str(numero).isdigit() and len(str(numero)) == 9 and str(numero).startswith("9"):
        return "Válido"
    return "Inválido"

# Iterar sobre la lista de preinscritos
for index, row in df.iterrows():
    nombre = row["Nombre"]
    apellido_paterno = row["Apellido_Paterno"]
    apellido_materno = row["Apellido_Materno"]
    dni = row["DNI"]
    programa = row["Programa"]
    name_programa = row["Name_Programa"]
    celular = str(row["Celular"])  # Convertir a string para evitar errores
    correo = row["Correo"]

    # Validar número
    estado = validar_numero(celular)

    # Si ya se envió antes, lo marcamos como "Repetido"
    if celular in df_enviados["Celular"].astype(str).values:
        estado = "Repetido"

    # Si el número es válido y no es repetido, enviar mensaje
    if estado == "Válido":
        mensaje = f"Hola {nombre}, te saludamos desde la Escuela de Posgrado de la UNAC. Queremos recordarte que tu preinscripción al programa {name_programa} está en proceso. Si tienes dudas, contáctanos."
        
        try:
            kit.sendwhatmsg_instantly("+51" + celular, mensaje, wait_time=10, tab_close=True)
            estado = "Enviado"
            time.sleep(5)  # Espera entre envíos para evitar bloqueos
        except Exception as e:
            estado = f"Error: {str(e)}"

    # Agregar los datos al DataFrame de enviados
    df_enviados = pd.concat([df_enviados, pd.DataFrame([{
        "Nombre": nombre,
        "Apellido_Paterno": apellido_paterno,
        "Apellido_Materno": apellido_materno,
        "DNI": dni,
        "Programa": programa,
        "Name_Programa": name_programa,
        "Celular": celular,
        "Correo": correo,
        "Revision": estado
    }])], ignore_index=True)

# Guardar en "Contactos_enviados.xlsx"
df_enviados.to_excel(archivo_enviados, index=False)
print("Proceso finalizado. Los mensajes fueron enviados y registrados.")
