import csv
import os
import pandas as pd
from datetime import datetime
from collections import defaultdict

# --- CONFIGURACIÓN DE INTEGRIDAD ---
PASOS_INTEGRIDAD = [
    "1. Limpieza pasillo", "2. Orden y acomodo", 
    "3. Validación Fechas", "4. Etiqueta Vizpick", 
    "5. Picks Vizpick", "6. Llenado BIN/Rotación"
]
ANIO_ACTUAL = datetime.now().year
# Este es el archivo "Cerebro" que NUNCA se borra
ARCHIVO_CSV = f"Base_Datos_Excelencia_{ANIO_ACTUAL}.csv"
PIN_SEGURIDAD = "8819"

def preparar_archivo():
    """Crea el encabezado solo si el archivo no existe."""
    if not os.path.exists(ARCHIVO_CSV):
        with open(ARCHIVO_CSV, mode='w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f)
            headers = ["Fecha", "Mes", "Semana", "Asociado", "Puntaje_Diario", "Estatus"]
            writer.writerow(headers)

def registrar_dia():
    print(f"\n{'═'*45}\n  REGISTRO DIARIO - EXCELENCIA 8819\n{'═'*45}")
    nom = input("Nombre del Asociado: ").strip().upper()
    if not nom: return

    ahora = datetime.now()
    fecha, mes, sem = ahora.strftime('%Y-%m-%d'), ahora.strftime('%B'), ahora.isocalendar()[1]

    print(f"\n¿Estado de {nom}?\n1. Evaluación\n2. Descanso")
    op_tipo = input("Seleccione (1/2): ")

    if op_tipo == "2":
        # Usamos mode='a' (Append) para añadir al final sin borrar nada
        with open(ARCHIVO_CSV, mode='a', newline='', encoding='utf-8') as f:
            csv.writer(f).writerow([fecha, mes, sem, nom, "DESCANSO", "Día de Descanso"])
        print(f"✅ Descanso de {nom} guardado en la base histórica."); return

    # Lógica de resta de puntos
    penalizacion_pasos = 0
    pasos_perfectos = True
    for p in PASOS_INTEGRIDAD:
        if input(f"¿Cumplió {p}? (s/n): ").lower().strip() != 's':
            penalizacion_pasos += 10
            pasos_perfectos = False

    # El castigo del 5% si no hubo integridad (40% -> 35%)
    valor_auditoria = 40 if pasos_perfectos else 35
    puntaje_auditoria = 0
    
    print(f"\n--- Paso 7: Auditoría (Máximo: {valor_auditoria}%) ---")
    if input("¿Pasó Auditoría Final? (s/n): ").lower().strip() == 's':
        puntaje_auditoria = valor_auditoria
        estatus = "EVALUADO"
    else:
        estatus = "AUDITORÍA FALLIDA"

    nota_final = (60 - penalizacion_pasos) + puntaje_auditoria
    if nota_final < 0: nota_final = 0

    # GUARDADO FINAL (AÑADIR AL HISTORIAL)
    with open(ARCHIVO_CSV, mode='a', newline='', encoding='utf-8') as f:
        csv.writer(f).writerow([fecha, mes, sem, nom, nota_final, estatus])
    print(f"✅ REGISTRO ACUMULADO: {nom} | NOTA: {nota_final}% | TOTAL SEMANAS: {sem}")

def exportar_a_excel():
    if input("\n🔐 Ingrese PIN 8819 para exportar: ") != PIN_SEGURIDAD: return
    if not os.path.exists(ARCHIVO_CSV): 
        print("❌ No hay datos guardados aún."); return
    
    try:
        # Cargamos toda la historia acumulada
        df = pd.read_csv(ARCHIVO_CSV)
        
        # Filtramos para el promedio (ignoramos descansos)
        df_eval = df[df['Puntaje_Diario'] != 'DESCANSO'].copy()
        df_eval['Puntaje_Diario'] = df_eval['Puntaje_Diario'].astype(float)
        
        # Promedio de TODO el historial acumulado
        ranking = df_eval.groupby('Asociado')['Puntaje_Diario'].mean().reset_index()
        ranking = ranking.sort_values(by='Puntaje_Diario', ascending=False)
        
        nombre_excel = f"REPORTE_SEMAFORO_ACUMULADO_{ANIO_ACTUAL}.xlsx"
        
        with pd.ExcelWriter(nombre_excel, engine='openpyxl') as writer:
            ranking.to_excel(writer, index=False, sheet_name='Ranking_General')
            df.to_excel(writer, index=False, sheet_name='Historial_Completo_52_Sem')
            
            ws = writer.sheets['Ranking_General']
            from openpyxl.styles import PatternFill
            # Colores del semáforo Wilfred
            v = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid") # Verde
            a = PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid") # Amarillo
            r = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid") # Rojo
            
            for row in range(2, len(ranking) + 2):
                celda = ws.cell(row=row, column=2)
                val = celda.value
                if val >= 85: celda.fill = v
                elif 51 <= val <= 84: celda.fill = a
                else: celda.fill = r
                
        print(f"\n📊 EXCEL CREADO SIN BORRAR DATOS: {nombre_excel}")
        print("Recuerda: El archivo CSV original sigue guardando todo el año.")
    except Exception as e:
        print(f"⚠️ Error al procesar: {e}")

if __name__ == "__main__":
    preparar_archivo()
    while True:
        print(f"\n--- CONTROL DE EXCELENCIA BIN {ANIO_ACTUAL} ---")
        print("1. Realizar Evaluación o Marcar Descanso")
        print("2. Generar Reporte Excel (Semáforo)")
        print("3. Salir")
        op = input("Seleccione una opción: ")
        if op == "1": registrar_dia()
        elif op == "2": exportar_a_excel()
        elif op == "3": break