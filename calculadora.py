
# Calculadora interactiva para ahorro del pie de una vivienda en Chile

def calcular_ahorro():
    # Valor aproximado de la UF (puedes actualizarlo)
    valor_uf = 36000  # pesos por UF

    print("=== Calculadora de Ahorro para Pie de Vivienda ===")
    valor_vivienda_uf = float(input("Ingrese el valor de la vivienda en UF: "))
    porcentaje_pie = float(input("Ingrese el porcentaje de pie (ej. 20): "))
    plazo_anios = int(input("Ingrese el plazo en años para ahorrar: "))

    # Conversión a pesos
    valor_vivienda_pesos = valor_vivienda_uf * valor_uf
    pie_pesos = valor_vivienda_pesos * (porcentaje_pie / 100)
    meses = plazo_anios * 12
    ahorro_mensual = pie_pesos / meses

    print("\n=== Resultados ===")
    print(f"Valor vivienda: ${valor_vivienda_pesos:,.0f} pesos")
    print(f"Pie ({porcentaje_pie}%): ${pie_pesos:,.0f} pesos")
    print(f"Ahorro mensual necesario en {plazo_anios} años: ${ahorro_mensual:,.0f} pesos")

# Ejecutar la calculadora
calcular_ahorro()
