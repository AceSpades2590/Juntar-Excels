def get_valid_column_number():
    while True:
        try:
            num_columnas = int(input("Ingrese el número de columnas a unir: "))
            return num_columnas
        except ValueError:
            print("Por favor, ingrese un número entero válido.")
