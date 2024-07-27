from flask import Flask, render_template, request
import pandas as pd
import os
from datetime import datetime

app = Flask(__name__)

@app.route('/', methods=['GET', 'POST'])
def index():
    mensaje = ''
    ruta_nueva = r'\\mlxgumvwfile01\Departamentos\Fakra\Pruebas\proyectos\aplicacionWeb\data\AutomatedLines.xlsx'#\\mlxgumvwfile01\Departamentos\Fakra\Pruebas\proyectos\aplicacionWeb
    filas_resaltadas = []

    if request.method == 'POST':
        numero_parte = request.form['numero_parte'].strip()
        df1 = pd.read_excel(ruta_nueva, sheet_name='Hoja1')
        df2 = pd.read_excel(ruta_nueva, sheet_name='Hoja2')
        df3 = pd.read_excel(ruta_nueva, sheet_name='Hoja3')

        # Convertir todas las celdas a string y eliminar espacios
        df1 = df1.applymap(lambda x: str(x).strip() if pd.notna(x) else x)
        df2 = df2.applymap(lambda x: str(x).strip() if pd.notna(x) else x)
        df3 = df3.applymap(lambda x: str(x).strip() if pd.notna(x) else x)

        # Buscar la fila correspondiente al número de parte ingresado
        fila_df1 = df1[df1.iloc[:, 0].astype(str).str.strip() == numero_parte]

        if fila_df1.empty:
            mensaje = f"No se encontró el número de parte '{numero_parte}' en la base de datos."
        else:
            resultados_ordenados = []  # Lista para almacenar los resultados

            for idx, row in fila_df1.iterrows():
                molex_pn = row[0]
                valores_buscados = row[1:].dropna().tolist()
                resultados_coincidencias = []

                # Agregar los valores buscados al inicio del mensaje
                mensaje += f"Housings: {valores_buscados}<br>"

                for index, row_df2 in df2.iterrows():
                    testers_encontrados = []
                    valores_encontrados = {val: 0 for val in valores_buscados}
                    
                    for value in row_df2.iloc[1:]:
                        if value in valores_buscados:
                            valores_encontrados[value] += 1

                    if all(valores_encontrados[val] >= valores_buscados.count(val) for val in valores_buscados):
                        testers_encontrados.append(row_df2['Line'])

                    if testers_encontrados:
                        resultados_coincidencias.extend(testers_encontrados)

                if resultados_coincidencias:
                    comunes_tester = set(resultados_coincidencias)

                    for tester in comunes_tester:
                        # Buscar el tester en la primera columna de Hoja3
                        fila_df3 = df3[df3.iloc[:, 0].astype(str).str.strip() == tester]
                        
                        if not fila_df3.empty:
                            # Verificar si el numero_parte está en la fila encontrada
                            encontrado = any(numero_parte in str(cell) for cell in fila_df3.iloc[0, :])
                            if encontrado:
                                resultados_ordenados.append(f"{molex_pn} - {tester}: <span style='color:green; font-weight:bold;'>Confirmado</span>")
                            else:
                                resultados_ordenados.append(f"{molex_pn} - {tester}: <span style='color:red; font-weight:bold;'>Confirmar con Ingeniería de Pruebas en piso</span>")
                        else:
                            resultados_ordenados.append(f"{molex_pn} - {tester}: Tester no habilitado")

                else:
                    mensaje += f"\nHousings: {valores_buscados}\nNo se encontraron valores de Tester comunes para Molex PN '{molex_pn}' en la fila {idx + 1}.\n\n"

            # Ordenar los resultados alfabéticamente por tester
            resultados_ordenados.sort()

            # Construir el mensaje final
            for resultado in resultados_ordenados:
                mensaje += f"<br>{resultado}<br>"

            # Comentar el bloque de código para resaltar las filas de df2
            """
            # Crear una versión resaltada de las filas de df2
            for index, row_df2 in df2.iterrows():
                resaltada = row_df2.copy()
                for col in row_df2.index[1:]:  # Excluye la primera columna (Line)
                    if row_df2[col] in valores_buscados:
                        resaltada[col] = f"<mark>{row_df2[col]}</mark>"
                filas_resaltadas.append(resaltada)
            """

    # Obtener la fecha y hora de modificación del archivo
    timestamp_modificacion = os.path.getmtime(ruta_nueva)
    fecha_modificacion = datetime.fromtimestamp(timestamp_modificacion).strftime('%Y-%m-%d  %H:%M:%S')
    #mensaje += f"<br><br><br><br>Última actualización: {fecha_modificacion}<br>"
    ultima_actualizacion = f"{fecha_modificacion}"

    return render_template('index.html', mensaje=mensaje, filas_resaltadas=filas_resaltadas, ultima_actualizacion=ultima_actualizacion)

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0')
