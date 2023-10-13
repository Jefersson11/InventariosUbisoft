import csv
import openpyxl
from openpyxl import Workbook

# Nombre del archivo Excel
archivo_excel = "inventario.xlsx"

# Función para agregar un producto al inventario
def agregar_producto():
    nombre = input("Ingrese el nombre del producto: ")
    cantidad = int(input("Ingrese la cantidad: "))
    categoria = input("Ingrese la categoría: ")

    # Agregar el producto al diccionario
    producto = {
        'nombre': nombre,
        'cantidad': cantidad,
        'categoria': categoria
    }
    inventario[nombre] = producto

    # Guardar el inventario en el archivo Excel
    guardar_inventario_excel()

# Función para eliminar un producto del inventario
def eliminar_producto():
    nombre = input("Ingrese el nombre del producto que desea eliminar: ")
    if nombre in inventario:
        del inventario[nombre]
        print(f"Producto '{nombre}' eliminado del inventario.")
        guardar_inventario_excel()
    else:
        print(f"Producto '{nombre}' no encontrado en el inventario.")

# Función para buscar un producto en el inventario por nombre
def buscar_producto():
    nombre = input("Ingrese el nombre del producto que desea buscar: ")
    if nombre in inventario:
        producto = inventario[nombre]
        print(f"Nombre: {producto['nombre']}")
        print(f"Cantidad: {producto['cantidad']}")
        print(f"Categoría: {producto['categoria']}")
    else:
        print(f"Producto '{nombre}' no encontrado en el inventario.")

# Función para mostrar el inventario
def mostrar_inventario():
    if not inventario:
        print("El inventario está vacío.")
    else:
        print("\nInventario:")
        for nombre, producto in inventario.items():
            print(f"Nombre: {producto['nombre']}")
            print(f"Cantidad: {producto['cantidad']}")
            print(f"Categoría: {producto['categoria']}")
            print("-------------")

# Función para guardar el inventario en un archivo Excel
def guardar_inventario_excel():
    libro = Workbook()
    hoja = libro.active
    hoja.append(['Nombre', 'Cantidad', 'Categoría'])

    for producto in inventario.values():
        hoja.append([producto['nombre'], producto['cantidad'], producto['categoria']])

    libro.save(archivo_excel)
    print("Inventario guardado en el archivo Excel.")

# Función para cargar el inventario desde el archivo Excel
def cargar_inventario_excel():
    try:
        libro = openpyxl.load_workbook(archivo_excel)
        hoja = libro.active
        for fila in hoja.iter_rows(min_row=2, values_only=True):
            nombre, cantidad, categoria = fila
            producto = {
                'nombre': nombre,
                'cantidad': cantidad,
                'categoria': categoria
            }
            inventario[nombre] = producto
    except FileNotFoundError:
        print("El archivo Excel no existe. Se creará uno nuevo al guardar el inventario.")

# Función principal del programa
def main():
    cargar_inventario_excel()
    while True:
        print("\nMenú:")
        print("1. Agregar producto")
        print("2. Eliminar producto")
        print("3. Buscar producto")
        print("4. Mostrar inventario")
        print("5. Salir")
        opcion = input("Seleccione una opción: ")

        if opcion == '1':
            agregar_producto()
        elif opcion == '2':
            eliminar_producto()
        elif opcion == '3':
            buscar_producto()
        elif opcion == '4':
            mostrar_inventario()
        elif opcion == '5':
            break
        else:
            print("Opción no válida. Intente de nuevo.")

if _name_ == "_main_":
    inventario = {}
    main()
