import pandas as pd
from tkinter import *
import unicodedata

# Cargamos el archivo Excel
df = pd.read_excel("C:/Users/EMA\\Desktop/1000-Registros-de-ventas.xlsx")

# Renombrando las columnas a números comenzando desde 1
df.columns = [str(i) for i in range(1, len(df.columns) + 1)]

# Mapeo de números de columnas a nombres deseados
mapeo_columnas = {
    '1': 'ID CLIENTE',
    '4': 'QUE VENDE',
    '6': 'URGENCIA',
    '7': 'PEDIDO',
    '9': 'ENVIO',
    '8': 'ID PEDIDO',
    '3': 'PAIS',
}

# Función para quitar acentos
# Función para normalizar texto: quitar acentos, convertir a mayusculas, quitar espacios extra y  prefijos
def normalizar_texto(cadena):
    cadena = unicodedata.normalize('NFKD', cadena).encode('ASCII', 'ignore').decode('ASCII')  # Quitamos acentos
    cadena = cadena.upper()  # Convertimos a Mayus
    cadena = cadena.strip()  # Quitamos espacios al principio y al final
    # Quitamos prefijos comunes
    prefijos = ['la c.', 'el c.', 'c.']
    for prefijo in prefijos:
        if cadena.startswith(prefijo):
            cadena = cadena[len(prefijo):].strip()
    return cadena

# Función que se ejecuta al realizar la búsqueda
def buscar():
    resultados.delete(0, END)  # Limpiar resultados previos
    texto_busqueda = normalizar_texto(entrada_texto.get())  # Normalizamos el texto ingresado
    columna_buscada = variable_columna.get()  # Obtenemos el valor del menú desplegable

    # Encontrar el número de columna correspondiente en el mapeo
    for num_columna, nombre_columna in mapeo_columnas.items():
        if nombre_columna == columna_buscada:
            columna_buscada = num_columna
            break

    if texto_busqueda:
        # Modificamos la condición de búsqueda para normalizar tanto el DataFrame como el texto ingresado
        filas = df.loc[df[columna_buscada].apply(lambda x: normalizar_texto(str(x))).str.contains(texto_busqueda, case=False, na=False)]
        if not filas.empty:
            for _, fila in filas.iterrows():
                for col in mapeo_columnas.keys():
                    resultados.insert(END, f"{mapeo_columnas[col]}: {fila[col]}")
                resultados.insert(END, "")  # Espacio en blanco entre resultados
        else:
            resultados.insert(END, "No se encontraron resultados.")

def on_enter(event):
    buscar()

# Creamos la ventana de la interfaz gráfica
ventana = Tk()
ventana.title('Buscador Excel')

frame_busqueda = Frame(ventana)
frame_busqueda.pack(fill=BOTH, expand=True, padx=10, pady=10)

# Texto "Ingresar:"
label_ingresar = Label(frame_busqueda, text="Ingresar:")
label_ingresar.pack(side=LEFT, padx=(0, 5))

entrada_texto = Entry(frame_busqueda)
entrada_texto.pack(side=LEFT, expand=True, fill=X)
entrada_texto.bind('<Return>', on_enter)

# Opciones para la búsqueda.
opciones_busqueda = ['ID CLIENTE', 'ID PEDIDO']  # Solo estas opciones se muestran

variable_columna = StringVar(ventana)
variable_columna.set(opciones_busqueda[0])  # Valor por defecto

opcion_columna = OptionMenu(frame_busqueda, variable_columna, *opciones_busqueda)
opcion_columna.pack(side=LEFT, padx=5)

boton_buscar = Button(frame_busqueda, text="Buscar", command=buscar)
boton_buscar.pack(side=LEFT, padx=5)

resultados = Listbox(ventana, width=100, height=20)
resultados.pack(padx=10, pady=10, fill=BOTH, expand=True)

ventana.mainloop()
















"""Los pasos serian los siguientes:

Intalación (PyInstaller)
pip install pyinstaller

Crear el archivo distribuido del ejectutable:
pyinstaller archivo.py

Opcionalmente podemos tener del segundo punto:

Desaparecer terminal de fondo (opcional). Para que desaparezca tenemos que indicar que es una aplicación en ventana, y eso lo hacemos de la siguiente forma al crear el ejecutable.
pyinstaller --windowed archivo.py

Ejecutable en un solo fichero (opcional). Por defecto Pyinstaller crea un directorio con un montón de ficheros. Podemos utilizar un comando para generar un solo fichero ejecutable que lo contenga todo, solo considera que este ocupara bastante más:
pyinstaller --windowed --onefile archivo.py

Cambiar el icono (opcional). También podemos cambiar el icono por defecto del ejecutable. Para ello necesitamos una imagen en formato .ico en el mismo directorio donde tenemos el script.
pyinstaller --windowed --onefile --icon=./hola.ico archivo.py """