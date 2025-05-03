from flask import Flask, request, jsonify, render_template, send_file
import requests
import json
import os
import logging
from datetime import datetime, timedelta
from threading import Lock, Thread
import time
import schedule
from flask_cors import CORS
import pandas as pd
import numpy as np
import io
import xlsxwriter
import tempfile

# Configuración de logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

app = Flask(__name__)
CORS(app)

# Configuración de la API de Ollama (usando la misma de Eva)
OLLAMA_URL = os.environ.get("OLLAMA_URL", "https://evaenespanol.loca.lt")
MODEL_NAME = os.environ.get("MODEL_NAME", "llama3:8b")

# Contexto del sistema para ContaFin
ASSISTANT_CONTEXT = """
# ContaFin: Agente Contable-Financiero Especializado en PYMEs

## Rol Principal:
Eres ContaFin, un producto de Innovación Financiera (www.innovacionfinanciera.com), empresa líder en soluciones contables y financieras. Como asistente experto en contabilidad, finanzas operativas y cumplimiento fiscal para pequeñas y medianas empresas, representas los valores de precisión, claridad y cumplimiento normativo. Tu objetivo es proporcionar programas contables personalizados, análisis financieros detallados y herramientas prácticas que ayuden a las PYMEs a gestionar su negocio con claridad financiera y cumplimiento legal.

## Objetivos Principales:
1. Diseñar programas contables adaptados al sector y tamaño de la empresa (ej: flujo de caja, nóminas, inventario).
2. Generar estados financieros clave:
   - Estado de Pérdidas y Ganancias (estructurado por costos fijos/variables).
   - Balance General (activos, pasivos, patrimonio).
   - Flujo de Efectivo (operativo, inversión, financiamiento).
3. Calcular obligaciones fiscales (IVA, retenciones, impuestos corporativos) según jurisdicción.
4. Elaborar análisis estratégicos:
   - Punto de equilibrio (en unidades y valor monetario).
   - Margen de contribución.
   - Ratios de liquidez/rentabilidad (current ratio, ROE, ROA).
5. Automatizar procesos críticos:
   - Cálculo de nóminas con desglose de aportes.
   - Conciliaciones bancarias.
   - Proyecciones financieras a 12 meses.

## Instrucciones Detalladas:

### 1. Análisis de Necesidades Financieras:
- Identificar brechas contables y prioridades del negocio.
- Evaluar el nivel de madurez financiera de la empresa.
- Determinar requisitos específicos según sector y jurisdicción.

### 2. Aplicación de Metodologías Especializadas:
- Implementar frameworks de análisis financiero reconocidos.
- Aplicar técnicas de Financial Intelligence for Entrepreneurs para análisis de ratios.
- Utilizar metodologías estándar de contabilidad adaptadas a PYMEs.

### 3. Generación de Herramientas Prácticas:
- Diseñar plantillas Excel automatizadas con fórmulas validadas.
- Crear checklists de cumplimiento fiscal mensual/trimestral.
- Desarrollar dashboards visuales para KPIs financieros clave.

### 4. Formato de Respuestas:
Para cada consulta, debes estructurar tu respuesta de la siguiente manera:

# Análisis [Tema] - [Fecha]
## Powered by Innovación Financiera - Expertos en Soluciones Contables y Financieras

## 1. Resumen de la Situación:
- [Breve diagnóstico inicial]
- [Identificación de necesidades clave]

## 2. Solución Propuesta:
- [Descripción del programa/herramienta/análisis]
- [Metodología aplicada]
- [Beneficios esperados]

## 3. Implementación Paso a Paso:
1. [Paso 1 detallado]
2. [Paso 2 detallado]
3. [Paso 3 detallado]

## 4. Herramientas y Recursos:
- [Plantillas Excel generadas]
- [Fórmulas clave explicadas]
- [Enlaces a recursos adicionales]

## 5. Recomendaciones Estratégicas:
- [Oportunidad de mejora identificada]
- [Consideraciones fiscales importantes]
- [Próximos pasos sugeridos]

Recuerda siempre:
1. Incluir un disclaimer: "Consulte a un profesional certificado en su jurisdicción para validar estos cálculos".
2. Proporcionar ejemplos concretos adaptados al sector y tamaño de la empresa.
3. Enfatizar la simplicidad y usabilidad de las soluciones propuestas.
4. Destacar las ventajas competitivas de usar soluciones de Innovación Financiera.
5. Mantener un tono profesional pero accesible para usuarios no expertos en finanzas.
"""

# Almacenamiento de sesiones e informes
sessions = {}
sessions_lock = Lock()
financial_reports = []
financial_reports_lock = Lock()
excel_templates = {}
excel_templates_lock = Lock()

def call_ollama_api(prompt, session_id, max_retries=3):
    """Llamar a la API de Ollama con reintentos"""
    headers = {
        "Content-Type": "application/json"
    }
    
    # Construir el mensaje para la API
    messages = []
    
    # Preparar el contexto del sistema
    system_context = ASSISTANT_CONTEXT
    
    # Agregar el contexto del sistema como primer mensaje
    messages.append({
        "role": "system",
        "content": system_context
    })
    
    # Agregar historial de conversación si existe la sesión
    with sessions_lock:
        if session_id in sessions:
            messages.extend(sessions[session_id])
    
    # Agregar el nuevo mensaje del usuario
    messages.append({
        "role": "user",
        "content": prompt
    })
    
    # Preparar los datos para la API
    data = {
        "model": MODEL_NAME,
        "messages": messages,
        "stream": False,
        "options": {
            "temperature": 0.7
        }
    }
    
    # Intentar con reintentos
    for attempt in range(max_retries):
        try:
            logger.info(f"Conectando a {OLLAMA_URL}...")
            response = requests.post(f"{OLLAMA_URL}/api/chat", headers=headers, json=data, timeout=120)
            
            # Si hay un error, intentar mostrar el mensaje
            if response.status_code >= 400:
                try:
                    error_data = response.json()
                    logger.error(f"Error detallado: {error_data}")
                except:
                    logger.error(f"Contenido del error: {response.text[:500]}")
                
                # Si obtenemos un 403, intentar con una URL alternativa
                if response.status_code == 403 and attempt == 0:
                    logger.info("Error 403, probando URL alternativa...")
                    alt_url = "http://127.0.0.1:11434/api/chat"
                    response = requests.post(alt_url, headers=headers, json=data, timeout=120)
            
            response.raise_for_status()
            response_data = response.json()
            
            # Extraer la respuesta según el formato de Ollama
            if "message" in response_data and "content" in response_data["message"]:
                return response_data["message"]["content"]
            else:
                logger.error(f"Formato de respuesta inesperado: {response_data}")
                return "Lo siento, no pude generar una respuesta apropiada en este momento."
            
        except requests.exceptions.RequestException as e:
            logger.error(f"Error en intento {attempt+1}/{max_retries}: {str(e)}")
            if attempt < max_retries - 1:
                wait_time = 2 ** attempt  # Retroceso exponencial
                logger.info(f"Reintentando en {wait_time} segundos...")
                time.sleep(wait_time)
            else:
                return f"Lo siento, estoy experimentando problemas técnicos de comunicación. ¿Podríamos intentarlo más tarde?"
    
    return "No se pudo conectar al servicio. Por favor, inténtelo de nuevo más tarde."

def call_ollama_completion(prompt, session_id, max_retries=3):
    """Usar el endpoint de completion en lugar de chat (alternativa)"""
    headers = {
        "Content-Type": "application/json"
    }
    
    # Construir prompt completo con contexto e historial
    full_prompt = ASSISTANT_CONTEXT + "\n\n"
    
    full_prompt += "Historial de conversación:\n"
    
    with sessions_lock:
        if session_id in sessions:
            for msg in sessions[session_id]:
                role = "Usuario" if msg["role"] == "user" else "ContaFin"
                full_prompt += f"{role}: {msg['content']}\n"
    
    full_prompt += f"\nUsuario: {prompt}\nContaFin: "
    
    # Preparar datos para API de completion
    data = {
        "model": MODEL_NAME,
        "prompt": full_prompt,
        "stream": False,
        "options": {
            "temperature": 0.7
        }
    }
    
    completion_url = f"{OLLAMA_URL}/api/generate"
    
    # Intentar con reintentos
    for attempt in range(max_retries):
        try:
            logger.info(f"Conectando a {completion_url}...")
            response = requests.post(completion_url, headers=headers, json=data, timeout=120)
            
            response.raise_for_status()
            response_data = response.json()
            
            # Extraer respuesta del formato de completion
            if "response" in response_data:
                return response_data["response"]
            else:
                logger.error(f"Formato de respuesta inesperado: {response_data}")
                return "Lo siento, no pude generar una respuesta apropiada en este momento."
            
        except requests.exceptions.RequestException as e:
            logger.error(f"Error en intento {attempt+1}/{max_retries}: {str(e)}")
            if attempt < max_retries - 1:
                wait_time = 2 ** attempt
                logger.info(f"Reintentando en {wait_time} segundos...")
                time.sleep(wait_time)
            else:
                return f"Lo siento, estoy experimentando problemas técnicos de comunicación. ¿Podríamos intentarlo más tarde?"
    
    return "No se pudo conectar al servicio. Por favor, inténtelo de nuevo más tarde."

def generate_excel_template(template_type, data=None):
    """Generar plantillas Excel según el tipo solicitado"""
    output = io.BytesIO()
    workbook = xlsxwriter.Workbook(output)
    
    # Formatos comunes
    header_format = workbook.add_format({
        'bold': True,
        'font_color': 'white',
        'bg_color': '#0066cc',
        'border': 1
    })
    
    cell_format = workbook.add_format({
        'border': 1
    })
    
    number_format = workbook.add_format({
        'border': 1,
        'num_format': '#,##0.00'
    })
    
    formula_format = workbook.add_format({
        'border': 1,
        'num_format': '#,##0.00',
        'bg_color': '#e6f2ff'
    })
    
    # Generar plantilla según tipo
    if template_type == "flujo_caja":
        # Crear hoja de Flujo de Caja
        worksheet = workbook.add_worksheet("Flujo de Caja")
        
        # Configurar anchos de columna
        worksheet.set_column('A:A', 30)
        worksheet.set_column('B:G', 15)
        
        # Título
        worksheet.merge_range('A1:G1', 'PLANTILLA DE FLUJO DE CAJA', header_format)
        worksheet.merge_range('A2:G2', 'Período: Enero - Junio 2025', workbook.add_format({'bold': True, 'align': 'center'}))
        
        # Encabezados
        months = ['Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio']
        worksheet.write('A4', 'CONCEPTO', header_format)
        for i, month in enumerate(months):
            worksheet.write(3, i+1, month, header_format)
        
        # Secciones
        row = 4
        
        # Saldo inicial
        worksheet.write(row, 0, 'Saldo Inicial', workbook.add_format({'bold': True}))
        worksheet.write_formula(row, 1, '=0', number_format)
        for i in range(1, len(months)):
            worksheet.write_formula(row, i+1, f'=G{row+1+(i-1)*20}', number_format)
        
        row += 2
        
        # Ingresos
        worksheet.write(row, 0, 'INGRESOS', workbook.add_format({'bold': True, 'bg_color': '#c6efce'}))
        row += 1
        ingresos = ['Ventas de Contado', 'Cobro a Clientes', 'Préstamos', 'Otros Ingresos']
        for ingreso in ingresos:
            worksheet.write(row, 0, ingreso, cell_format)
            for i in range(len(months)):
                worksheet.write(row, i+1, 0, number_format)
            row += 1
        
        # Total ingresos
        worksheet.write(row, 0, 'Total Ingresos', workbook.add_format({'bold': True, 'bg_color': '#c6efce'}))
        for i in range(len(months)):
            worksheet.write_formula(row, i+1, f'=SUM(B{row-len(ingresos)}:B{row})', formula_format)
        
        row += 2
        
        # Egresos
        worksheet.write(row, 0, 'EGRESOS', workbook.add_format({'bold': True, 'bg_color': '#ffc7ce'}))
        row += 1
        egresos = ['Compras de Mercancía', 'Pago a Proveedores', 'Nómina', 'Impuestos', 'Servicios', 'Arrendamiento', 'Otros Gastos']
        for egreso in egresos:
            worksheet.write(row, 0, egreso, cell_format)
            for i in range(len(months)):
                worksheet.write(row, i+1, 0, number_format)
            row += 1
        
        # Total egresos
        worksheet.write(row, 0, 'Total Egresos', workbook.add_format({'bold': True, 'bg_color': '#ffc7ce'}))
        for i in range(len(months)):
            worksheet.write_formula(row, i+1, f'=SUM(B{row-len(egresos)}:B{row})', formula_format)
        
        row += 2
        
        # Flujo neto
        worksheet.write(row, 0, 'Flujo Neto del Mes', workbook.add_format({'bold': True}))
        inicio_ingresos = 7
        total_ingresos = inicio_ingresos + len(ingresos)
        total_egresos = row
        for i in range(len(months)):
            worksheet.write_formula(row, i+1, f'=B{total_ingresos}-B{total_egresos}', formula_format)
        
        row += 1
        
        # Saldo final
        worksheet.write(row, 0, 'Saldo Final', workbook.add_format({'bold': True}))
        for i in range(len(months)):
            worksheet.write_formula(row, i+1, f'=B5+B{row}', formula_format)
        
        # Agregar comentarios/instrucciones
        row += 3
        worksheet.merge_range(f'A{row}:G{row}', 'INSTRUCCIONES:', workbook.add_format({'bold': True}))
        row += 1
        instructions = [
            '1. Comience registrando su saldo inicial en la celda B5.',
            '2. Ingrese sus ingresos mensuales previstos en cada categoría.',
            '3. Registre todos los egresos esperados en cada mes.',
            '4. El flujo neto y saldo final se calcularán automáticamente.',
            '5. El saldo final de cada mes se traslada como saldo inicial del mes siguiente.',
            '6. Actualice mensualmente los datos reales para mantener control de su liquidez.'
        ]
        for instruction in instructions:
            worksheet.merge_range(f'A{row}:G{row}', instruction)
            row += 1
        
    elif template_type == "nomina":
        # Crear hoja de Nómina
        worksheet = workbook.add_worksheet("Nómina")
        
        # Configurar anchos de columna
        worksheet.set_column('A:A', 25)
        worksheet.set_column('B:L', 13)
        
        # Título
        worksheet.merge_range('A1:L1', 'PLANTILLA DE CÁLCULO DE NÓMINA', header_format)
        worksheet.merge_range('A2:L2', 'Mes: Mayo 2025', workbook.add_format({'bold': True, 'align': 'center'}))
        
        # Encabezados de empleados
        row = 4
        worksheet.write(row, 0, 'Empleado', header_format)
        worksheet.write(row, 1, 'Cargo', header_format)
        worksheet.write(row, 2, 'Salario Base', header_format)
        worksheet.write(row, 3, 'Días Trabajados', header_format)
        worksheet.write(row, 4, 'Horas Extra', header_format)
        worksheet.write(row, 5, 'Bonificaciones', header_format)
        worksheet.write(row, 6, 'Salario Total', header_format)
        worksheet.write(row, 7, 'Seguridad Social', header_format)
        worksheet.write(row, 8, 'Retención Renta', header_format)
        worksheet.write(row, 9, 'Otros Descuentos', header_format)
        worksheet.write(row, 10, 'Total Descuentos', header_format)
        worksheet.write(row, 11, 'Salario Neto', header_format)
        
        # Datos de ejemplo (5 empleados)
        cargos = ['Gerente', 'Analista', 'Asistente', 'Técnico', 'Operario']
        salarios = [5000000, 3000000, 2000000, 1800000, 1200000]
        
        row += 1
        for i in range(5):
            worksheet.write(row, 0, f'Empleado {i+1}', cell_format)
            worksheet.write(row, 1, cargos[i], cell_format)
            worksheet.write(row, 2, salarios[i], number_format)
            worksheet.write(row, 3, 30, cell_format)  # Días trabajados
            worksheet.write(row, 4, 0, cell_format)   # Horas extra
            worksheet.write(row, 5, 0, number_format) # Bonificaciones
            
            # Salario total (Base + (HE * valor hora * 1.5) + Bonificaciones)
            worksheet.write_formula(row, 6, 
                f'=C{row+1}*(D{row+1}/30)+((C{row+1}/240)*1.5*E{row+1})+F{row+1}', 
                formula_format)
            
            # Seguridad social (8% del salario base)
            worksheet.write_formula(row, 7, f'=C{row+1}*0.08', formula_format)
            
            # Retención en la fuente (asumiendo 10% simplificado)
            worksheet.write_formula(row, 8, f'=IF(C{row+1}>2500000,G{row+1}*0.1,0)', formula_format)
            
            # Otros descuentos
            worksheet.write(row, 9, 0, number_format)
            
            # Total descuentos
            worksheet.write_formula(row, 10, f'=SUM(H{row+1}:J{row+1})', formula_format)
            
            # Salario neto
            worksheet.write_formula(row, 11, f'=G{row+1}-K{row+1}', formula_format)
            
            row += 1
        
        # Totales
        row += 1
        worksheet.write(row, 0, 'TOTALES', workbook.add_format({'bold': True, 'bg_color': '#d9d9d9'}))
        worksheet.merge_range(f'B{row+1}:B{row+1}', '', workbook.add_format({'bg_color': '#d9d9d9'}))
        
        # Sumar cada columna numérica
        for col in range(2, 12):
            worksheet.write_formula(row, col, f'=SUM({chr(65+col)}6:{chr(65+col)}{row})', 
                                   workbook.add_format({'bold': True, 'border': 1, 'num_format': '#,##0.00', 'bg_color': '#d9d9d9'}))
        
        # Resumen para la empresa
        row += 3
        worksheet.merge_range(f'A{row}:L{row}', 'RESUMEN DE COSTOS PARA LA EMPRESA', 
                             workbook.add_format({'bold': True, 'bg_color': '#4472c4', 'font_color': 'white', 'align': 'center'}))
        
        row += 1
        worksheet.merge_range(f'A{row}:B{row}', 'Total Salarios Netos:', workbook.add_format({'bold': True}))
        worksheet.write_formula(f'C{row}', '=L11', number_format)
        
        row += 1
        worksheet.merge_range(f'A{row}:B{row}', 'Aportes Patronales (20%):', workbook.add_format({'bold': True}))
        worksheet.write_formula(f'C{row}', '=G11*0.2', number_format)
        
        row += 1
        worksheet.merge_range(f'A{row}:B{row}', 'Prima (1/12 salario anual):', workbook.add_format({'bold': True}))
        worksheet.write_formula(f'C{row}', '=G11/12', number_format)
        
        row += 1
        worksheet.merge_range(f'A{row}:B{row}', 'Cesantías (1/12 salario anual):', workbook.add_format({'bold': True}))
        worksheet.write_formula(f'C{row}', '=G11/12', number_format)
        
        row += 1
        worksheet.merge_range(f'A{row}:B{row}', 'Intereses Cesantías (12% anual):', workbook.add_format({'bold': True}))
        worksheet.write_formula(f'C{row}', '=C16*0.12/12', number_format)
        
        row += 1
        worksheet.merge_range(f'A{row}:B{row}', 'Vacaciones (1/24 salario anual):', workbook.add_format({'bold': True}))
        worksheet.write_formula(f'C{row}', '=G11/24', number_format)
        
        row += 1
        worksheet.merge_range(f'A{row}:B{row}', 'Costo Total Empresa:', 
                             workbook.add_format({'bold': True, 'bg_color': '#ffff00'}))
        worksheet.write_formula(f'C{row}', '=SUM(C13:C18)', 
                               workbook.add_format({'bold': True, 'border': 1, 'num_format': '#,##0.00', 'bg_color': '#ffff00'}))
        
        # Instrucciones
        row += 3
        worksheet.merge_range(f'A{row}:L{row}', 'INSTRUCCIONES:', workbook.add_format({'bold': True}))
        row += 1
        instructions = [
            '1. Ingrese el nombre de cada empleado y su cargo en las columnas A y B.',
            '2. Registre el salario base mensual en la columna C.',
            '3. Actualice los días trabajados, horas extra y bonificaciones según corresponda.',
            '4. Los cálculos de salario total, descuentos y neto se realizan automáticamente.',
            '5. El resumen muestra el costo total para la empresa incluyendo provisiones sociales.',
            '6. Ajuste los porcentajes según la legislación vigente en su país.'
        ]
        for instruction in instructions:
            worksheet.merge_range(f'A{row}:L{row}', instruction)
            row += 1
            
    elif template_type == "balance_general":
        # Crear hoja de Balance General
        worksheet = workbook.add_worksheet("Balance General")
        
        # Configurar anchos de columna
        worksheet.set_column('A:A', 40)
        worksheet.set_column('B:C', 18)
        
        # Título
        worksheet.merge_range('A1:C1', 'BALANCE GENERAL', header_format)
        worksheet.merge_range('A2:C2', 'Al 31 de Mayo de 2025', workbook.add_format({'bold': True, 'align': 'center'}))
        
        # ACTIVOS
        row = 4
        worksheet.merge_range(f'A{row}:C{row}', 'ACTIVOS', workbook.add_format({'bold': True, 'bg_color': '#4472c4', 'font_color': 'white', 'align': 'center'}))
        
        # Activos Corrientes
        row += 1
        worksheet.write(row, 0, 'ACTIVOS CORRIENTES', workbook.add_format({'bold': True, 'bg_color': '#d9d9d9'}))
        worksheet.merge_range(f'B{row+1}:C{row+1}', '', workbook.add_format({'bg_color': '#d9d9d9'}))
        
        row += 1
        activos_corrientes = ['Efectivo y Equivalentes', 'Inversiones Temporales', 'Cuentas por Cobrar', 'Inventarios', 'Gastos Pagados por Anticipado']
        for activo in activos_corrientes:
            worksheet.write(row, 0, activo, cell_format)
            worksheet.write(row, 1, 0, number_format)  # Valor inicial 0
            worksheet.write(row, 2, '', cell_format)   # Para notas o detalles
            row += 1
        
        # Total Activos Corrientes
        worksheet.write(row, 0, 'Total Activos Corrientes', workbook.add_format({'bold': True}))
        start_row = row - len(activos_corrientes) + 1
        worksheet.write_formula(row, 1, f'=SUM(B{start_row}:B{row})', formula_format)
        worksheet.write(row, 2, '', cell_format)
        
        # Activos No Corrientes
        row += 2
        worksheet.write(row, 0, 'ACTIVOS NO CORRIENTES', workbook.add_format({'bold': True, 'bg_color': '#d9d9d9'}))
        worksheet.merge_range(f'B{row+1}:C{row+1}', '', workbook.add_format({'bg_color': '#d9d9d9'}))
        
        row += 1
        activos_no_corrientes = ['Propiedad, Planta y Equipo', 'Depreciación Acumulada', 'Intangibles', 'Inversiones a Largo Plazo', 'Otros Activos']
        for i, activo in enumerate(activos_no_corrientes):
            # Depreciación acumulada va con valor negativo
            valor = 0
            if activo == 'Depreciación Acumulada':
                formula_format.set_num_format('#,##0.00_);(#,##0.00)')
            
            worksheet.write(row, 0, activo, cell_format)
            worksheet.write(row, 1, valor, number_format)
            worksheet.write(row, 2, '', cell_format)
            row += 1
        
        # Total Activos No Corrientes
        worksheet.write(row, 0, 'Total Activos No Corrientes', workbook.add_format({'bold': True}))
        start_row = row - len(activos_no_corrientes) + 1
        worksheet.write_formula(row, 1, f'=SUM(B{start_row}:B{row})', formula_format)
        worksheet.write(row, 2, '', cell_format)
        
        # TOTAL ACTIVOS
        row += 2
        worksheet.write(row, 0, 'TOTAL ACTIVOS', workbook.add_format({'bold': True, 'bg_color': '#4472c4', 'font_color': 'white'}))
        activos_corrientes_row = 11  # Ajustar según la posición real
        activos_no_corrientes_row = 19  # Ajustar según la posición real
        worksheet.write_formula(row, 1, f'=B{activos_corrientes_row}+B{activos_no_corrientes_row}', 
                               workbook.add_format({'bold': True, 'border': 1, 'num_format': '#,##0.00', 'bg_color': '#c6efce'}))
        
        # PASIVOS
        row += 2
        worksheet.merge_range(f'A{row}:C{row}', 'PASIVOS', workbook.add_format({'bold': True, 'bg_color': '#4472c4', 'font_color': 'white', 'align': 'center'}))
        
        # Pasivos Corrientes
        row += 1
        worksheet.write(row, 0, 'PASIVOS CORRIENTES', workbook.add_format({'bold': True, 'bg_color': '#d9d9d9'}))
        worksheet.merge_range(f'B{row+1}:C{row+1}', '', workbook.add_format({'bg_color': '#d9d9d9'}))
        
        row += 1
        pasivos_corrientes = ['Cuentas por Pagar', 'Obligaciones Financieras CP', 'Impuestos por Pagar', 'Obligaciones Laborales', 'Anticipos de Clientes']
        for pasivo in pasivos_corrientes:
            worksheet.write(row, 0, pasivo, cell_format)
            worksheet.write(row, 1, 0, number_format)  # Valor inicial 0
            worksheet.write(row, 2, '', cell_format)   # Para notas o detalles
            row += 1
        
        # Total Pasivos Corrientes
        worksheet.write(row, 0, 'Total Pasivos Corrientes', workbook.add_format({'bold': True}))
        start_row = row - len(pasivos_corrientes) + 1
        worksheet.write_formula(row, 1, f'=SUM(B{start_row}:B{row})', formula_format)
        worksheet.write(row, 2, '', cell_format)
        pasivos_corrientes_row = row
        
        # Pasivos No Corrientes
        row += 2
        worksheet.write(row, 0, 'PASIVOS NO CORRIENTES', workbook.add_format({'bold': True, 'bg_color': '#d9d9d9'}))
        worksheet.merge_range(f'B{row+1}:C{row+1}', '', workbook.add_format({'bg_color': '#d9d9d9'}))
        
        row += 1
        pasivos_no_corrientes = ['Obligaciones Financieras LP', 'Pasivos Estimados', 'Otros Pasivos LP']
        for pasivo in pasivos_no_corrientes:
            worksheet.write(row, 0, pasivo, cell_format)
            worksheet.write(row, 1, 0, number_format)  # Valor inicial 0
            worksheet.write(row, 2, '', cell_format)   # Para notas o detalles
            row += 1
        
        # Total Pasivos No Corrientes
        worksheet.write(row, 0, 'Total Pasivos No Corrientes', workbook.add_format({'bold': True}))
        start_row = row - len(pasivos_no_corrientes) + 1
        worksheet.write_formula(row, 1, f'=SUM(B{start_row}:B{row})', formula_format)
        worksheet.write(row, 2, '', cell_format)
        pasivos_no_corrientes_row = row
        
        # TOTAL PASIVOS
        row += 2
        worksheet.write(row, 0, 'TOTAL PASIVOS', workbook.add_format({'bold': True, 'bg_color': '#4472c4', 'font_color': 'white'}))
        worksheet.write_formula(row, 1, f'=B{pasivos_corrientes_row}+B{pasivos_no_corrientes_row}', 
                              workbook.add_format({'bold': True, 'border': 1, 'num_format': '#,##0.00', 'bg_color': '#ffc7ce'}))
        pasivos_totales_row = row
        
        # PATRIMONIO
        row += 2
        worksheet.merge_range(f'A{row}:C{row}', 'PATRIMONIO', workbook.add_format({'bold': True, 'bg_color': '#4472c4', 'font_color': 'white', 'align': 'center'}))
        
        row += 1
        worksheet.write(row, 0, 'Capital Social', cell_format)
        worksheet.write(row, 1, 0, number_format)
        worksheet.write(row, 2, '', cell_format)
        row += 1
        
        worksheet.write(row, 0, 'Reservas', cell_format)
        worksheet.write(row, 1, 0, number_format)
        worksheet.write(row, 2, '', cell_format)
        row += 1
        
        worksheet.write(row, 0, 'Utilidades Retenidas', cell_format)
        worksheet.write(row, 1, 0, number_format)
        worksheet.write(row, 2, '', cell_format)
        row += 1
        
        worksheet.write(row, 0, 'Utilidad del Ejercicio', cell_format)
        worksheet.write(row, 1, 0, number_format)
        worksheet.write(row, 2, '', cell_format)
        row += 1
        
        # Total Patrimonio
        worksheet.write(row, 0, 'TOTAL PATRIMONIO', workbook.add_format({'bold': True}))
        start_row = row - 4 + 1
        worksheet.write_formula(row, 1, f'=SUM(B{start_row}:B{row})', formula_format)
        worksheet.write(row, 2, '', cell_format)
        patrimonio_row = row
        
        # TOTAL PASIVO + PATRIMONIO
        row += 2
        worksheet.write(row, 0, 'TOTAL PASIVO + PATRIMONIO', workbook.add_format({'bold': True, 'bg_color': '#4472c4', 'font_color': 'white'}))
        worksheet.write_formula(row, 1, f'=B{pasivos_totales_row}+B{patrimonio_row}', 
                              workbook.add_format({'bold': True, 'border': 1, 'num_format': '#,##0.00', 'bg_color': '#c6efce'}))
        
        # Ecuación contable (verificación)
        row += 2
        worksheet.write(row, 0, 'VERIFICACIÓN ECUACIÓN CONTABLE', workbook.add_format({'bold': True}))
        activos_totales_row = 22  # Ajustar según la posición real
        pasivo_patrimonio_row = row - 2
        worksheet.write_formula(row, 1, f'=IF(B{activos_totales_row}=B{pasivo_patrimonio_row},"CORRECTO","ERROR")', 
                              workbook.add_format({'bold': True, 'border': 1}))
        
        # Instrucciones
        row += 3
        worksheet.merge_range(f'A{row}:C{row}', 'INSTRUCCIONES:', workbook.add_format({'bold': True}))
        row += 1
        instructions = [
            '1. Complete los valores de cada cuenta en la columna B.',
            '2. La columna C puede utilizarse para notas o detalles adicionales.',
            '3. Los totales y subtotales se calculan automáticamente.',
            '4. Verifique que la ecuación contable (Activo = Pasivo + Patrimonio) esté balanceada.',
            '5. Actualice la fecha del balance en la celda A2.',
            '6. Este balance es una plantilla básica, personalice según las necesidades de su empresa.'
        ]
        for instruction in instructions:
            worksheet.merge_range(f'A{row}:C{row}', instruction)
            row += 1
    
    elif template_type == "estado_resultados":
        # Crear hoja de Estado de Resultados
        worksheet = workbook.add_worksheet("Estado de Resultados")
        
        # Configurar anchos de columna
        worksheet.set_column('A:A', 40)
        worksheet.set_column('B:C', 15)
        
        # Título
        worksheet.merge_range('A1:C1', 'ESTADO DE RESULTADOS', header_format)
        worksheet.merge_range('A2:C2', 'Del 1 al 31 de Mayo de 2025', workbook.add_format({'bold': True, 'align': 'center'}))
        
        row = 4
        # Ingresos Operacionales
        worksheet.write(row, 0, 'INGRESOS OPERACIONALES', workbook.add_format({'bold': True, 'bg_color': '#4472c4', 'font_color': 'white'}))
        worksheet.merge_range(f'B{row+1}:C{row+1}', '', workbook.add_format({'bg_color': '#4472c4', 'font_color': 'white'}))
        
        row += 1
        ingresos = ['Ventas Brutas', 'Devoluciones en Ventas', 'Descuentos Comerciales']
        
        for i, ingreso in enumerate(ingresos):
            formato = number_format
            if i > 0:  # Devoluciones y descuentos son negativos
                formato = workbook.add_format({'border': 1, 'num_format': '#,##0.00_);(#,##0.00)'})
            
            worksheet.write(row, 0, ingreso, cell_format)
            worksheet.write(row, 1, 0, formato)
            worksheet.write(row, 2, '', cell_format)
            row += 1
        
        # Total Ingresos Netos
        worksheet.write(row, 0, 'TOTAL INGRESOS NETOS', workbook.add_format({'bold': True}))
        worksheet.write_formula(row, 1, '=B5-B6-B7', formula_format)  # Ajustar según posición real
        worksheet.write(row, 2, '', cell_format)
        ingresos_netos_row = row
        
        row += 2
        # Costo de Ventas
        worksheet.write(row, 0, 'COSTO DE VENTAS', workbook.add_format({'bold': True, 'bg_color': '#4472c4', 'font_color': 'white'}))
        worksheet.merge_range(f'B{row+1}:C{row+1}', '', workbook.add_format({'bg_color': '#4472c4', 'font_color': 'white'}))
        
        row += 1
        costos = ['Inventario Inicial', 'Compras', 'Fletes en Compras', 'Inventario Final']
        
        for i, costo in enumerate(costos):
            formato = number_format
            if costo == 'Inventario Final':  # Inventario final es negativo en el cálculo
                formato = workbook.add_format({'border': 1, 'num_format': '#,##0.00_);(#,##0.00)'})
            
            worksheet.write(row, 0, costo, cell_format)
            worksheet.write(row, 1, 0, formato)
            worksheet.write(row, 2, '', cell_format)
            row += 1
        
        # Total Costo de Ventas
        worksheet.write(row, 0, 'TOTAL COSTO DE VENTAS', workbook.add_format({'bold': True}))
        worksheet.write_formula(row, 1, '=B10+B11+B12-B13', formula_format)  # Ajustar según posición real
        worksheet.write(row, 2, '', cell_format)
        costo_ventas_row = row
        
        row += 2
        # Utilidad Bruta
        worksheet.write(row, 0, 'UTILIDAD BRUTA', workbook.add_format({'bold': True, 'bg_color': '#c6efce'}))
        worksheet.write_formula(row, 1, f'=B{ingresos_netos_row}-B{costo_ventas_row}', 
                              workbook.add_format({'bold': True, 'border': 1, 'num_format': '#,##0.00', 'bg_color': '#c6efce'}))
        worksheet.write(row, 2, '', workbook.add_format({'bg_color': '#c6efce'}))
        utilidad_bruta_row = row
        
        row += 2
        # Gastos Operacionales
        worksheet.write(row, 0, 'GASTOS OPERACIONALES', workbook.add_format({'bold': True, 'bg_color': '#4472c4', 'font_color': 'white'}))
        worksheet.merge_range(f'B{row+1}:C{row+1}', '', workbook.add_format({'bg_color': '#4472c4', 'font_color': 'white'}))
        
        row += 1
        # Gastos de Administración
        worksheet.write(row, 0, 'Gastos de Administración', workbook.add_format({'bold': True, 'italic': True}))
        worksheet.merge_range(f'B{row+1}:C{row+1}', '', cell_format)
        
        row += 1
        gastos_admin = ['Nómina Administrativa', 'Honorarios', 'Impuestos', 'Arrendamientos', 'Seguros', 'Servicios', 'Depreciación', 'Diversos']
        
        for gasto in gastos_admin:
            worksheet.write(row, 0, '  ' + gasto, cell_format)  # Indentación con espacios
            worksheet.write(row, 1, 0, number_format)
            worksheet.write(row, 2, '', cell_format)
            row += 1
        
        # Total Gastos Administración
        worksheet.write(row, 0, 'Total Gastos Administración', workbook.add_format({'bold': True}))
        start_row = row - len(gastos_admin) + 1
        worksheet.write_formula(row, 1, f'=SUM(B{start_row}:B{row})', formula_format)
        worksheet.write(row, 2, '', cell_format)
        gastos_admin_row = row
        
        row += 1
        # Gastos de Ventas
        worksheet.write(row, 0, 'Gastos de Ventas', workbook.add_format({'bold': True, 'italic': True}))
        worksheet.merge_range(f'B{row+1}:C{row+1}', '', cell_format)
        
        row += 1
        gastos_ventas = ['Nómina Ventas', 'Comisiones', 'Publicidad', 'Gastos de Viaje', 'Diversos']
        
        for gasto in gastos_ventas:
            worksheet.write(row, 0, '  ' + gasto, cell_format)  # Indentación con espacios
            worksheet.write(row, 1, 0, number_format)
            worksheet.write(row, 2, '', cell_format)
            row += 1
        
        # Total Gastos Ventas
        worksheet.write(row, 0, 'Total Gastos Ventas', workbook.add_format({'bold': True}))
        start_row = row - len(gastos_ventas) + 1
        worksheet.write_formula(row, 1, f'=SUM(B{start_row}:B{row})', formula_format)
        worksheet.write(row, 2, '', cell_format)
        gastos_ventas_row = row
        
        # Total Gastos Operacionales
        row += 1
        worksheet.write(row, 0, 'TOTAL GASTOS OPERACIONALES', workbook.add_format({'bold': True}))
        worksheet.write_formula(row, 1, f'=B{gastos_admin_row}+B{gastos_ventas_row}', formula_format)
        worksheet.write(row, 2, '', cell_format)
        gastos_operacionales_row = row
        
        row += 2
        # Utilidad Operacional
        worksheet.write(row, 0, 'UTILIDAD OPERACIONAL', workbook.add_format({'bold': True, 'bg_color': '#c6efce'}))
        worksheet.write_formula(row, 1, f'=B{utilidad_bruta_row}-B{gastos_operacionales_row}', 
                              workbook.add_format({'bold': True, 'border': 1, 'num_format': '#,##0.00', 'bg_color': '#c6efce'}))
        worksheet.write(row, 2, '', workbook.add_format({'bg_color': '#c6efce'}))
        utilidad_operacional_row = row
        
        row += 2
        # Ingresos y Gastos No Operacionales
        worksheet.write(row, 0, 'INGRESOS NO OPERACIONALES', workbook.add_format({'bold': True, 'bg_color': '#4472c4', 'font_color': 'white'}))
        worksheet.merge_range(f'B{row+1}:C{row+1}', '', workbook.add_format({'bg_color': '#4472c4', 'font_color': 'white'}))
        
        row += 1
        ingresos_no_op = ['Ingresos Financieros', 'Otros Ingresos']
        
        for ingreso in ingresos_no_op:
            worksheet.write(row, 0, ingreso, cell_format)
            worksheet.write(row, 1, 0, number_format)
            worksheet.write(row, 2, '', cell_format)
            row += 1
        
        # Total Ingresos No Operacionales
        worksheet.write(row, 0, 'Total Ingresos No Operacionales', workbook.add_format({'bold': True}))
        start_row = row - len(ingresos_no_op) + 1
        worksheet.write_formula(row, 1, f'=SUM(B{start_row}:B{row})', formula_format)
        worksheet.write(row, 2, '', cell_format)
        ingresos_no_op_row = row
        
        row += 1
        # Gastos No Operacionales
        worksheet.write(row, 0, 'GASTOS NO OPERACIONALES', workbook.add_format({'bold': True, 'bg_color': '#4472c4', 'font_color': 'white'}))
        worksheet.merge_range(f'B{row+1}:C{row+1}', '', workbook.add_format({'bg_color': '#4472c4', 'font_color': 'white'}))
        
        row += 1
        gastos_no_op = ['Gastos Financieros', 'Otros Gastos']
        
        for gasto in gastos_no_op:
            worksheet.write(row, 0, gasto, cell_format)
            worksheet.write(row, 1, 0, number_format)
            worksheet.write(row, 2, '', cell_format)
            row += 1
        
        # Total Gastos No Operacionales
        worksheet.write(row, 0, 'Total Gastos No Operacionales', workbook.add_format({'bold': True}))
        start_row = row - len(gastos_no_op) + 1
        worksheet.write_formula(row, 1, f'=SUM(B{start_row}:B{row})', formula_format)
        worksheet.write(row, 2, '', cell_format)
        gastos_no_op_row = row
        
        row += 2
        # Utilidad Antes de Impuestos
        worksheet.write(row, 0, 'UTILIDAD ANTES DE IMPUESTOS', workbook.add_format({'bold': True, 'bg_color': '#c6efce'}))
        worksheet.write_formula(row, 1, f'=B{utilidad_operacional_row}+B{ingresos_no_op_row}-B{gastos_no_op_row}', 
                              workbook.add_format({'bold': True, 'border': 1, 'num_format': '#,##0.00', 'bg_color': '#c6efce'}))
        worksheet.write(row, 2, '', workbook.add_format({'bg_color': '#c6efce'}))
        utilidad_antes_imp_row = row
        
        row += 1
        # Provisión Impuesto de Renta
        worksheet.write(row, 0, 'Provisión Impuesto de Renta (30%)', cell_format)
        worksheet.write_formula(row, 1, f'=IF(B{utilidad_antes_imp_row}>0,B{utilidad_antes_imp_row}*0.3,0)', number_format)
        worksheet.write(row, 2, '', cell_format)
        
        row += 2
        # Utilidad Neta
        worksheet.write(row, 0, 'UTILIDAD NETA DEL EJERCICIO', workbook.add_format({'bold': True, 'bg_color': '#ffff00'}))
        worksheet.write_formula(row, 1, f'=B{utilidad_antes_imp_row}-B{row-1}', 
                              workbook.add_format({'bold': True, 'border': 1, 'num_format': '#,##0.00', 'bg_color': '#ffff00'}))
        worksheet.write(row, 2, '', workbook.add_format({'bg_color': '#ffff00'}))
        
        # Instrucciones
        row += 3
        worksheet.merge_range(f'A{row}:C{row}', 'INSTRUCCIONES:', workbook.add_format({'bold': True}))
        row += 1
        instructions = [
            '1. Complete los valores de cada cuenta en la columna B.',
            '2. La columna C puede utilizarse para notas o análisis vertical (%).',
            '3. Los totales y subtotales se calculan automáticamente.',
            '4. Actualice la fecha del estado de resultados en la celda A2.',
            '5. Ajuste el porcentaje de impuesto de renta según la tasa vigente en su país.',
            '6. Esta plantilla es configurable según las necesidades de su empresa.'
        ]
        for instruction in instructions:
            worksheet.merge_range(f'A{row}:C{row}', instruction)
            row += 1
    
    elif template_type == "punto_equilibrio":
        # Crear hoja de Punto de Equilibrio
        worksheet = workbook.add_worksheet("Punto de Equilibrio")
        
        # Configurar anchos de columna
        worksheet.set_column('A:A', 30)
        worksheet.set_column('B:D', 15)
        
        # Título
        worksheet.merge_range('A1:D1', 'ANÁLISIS DE PUNTO DE EQUILIBRIO', header_format)
        worksheet.merge_range('A2:D2', 'Evaluación de Rentabilidad Operativa', workbook.add_format({'bold': True, 'align': 'center'}))
        
        row = 4
        # Sección de datos de entrada
        worksheet.merge_range(f'A{row}:D{row}', 'DATOS DE ENTRADA', workbook.add_format({'bold': True, 'bg_color': '#4472c4', 'font_color': 'white', 'align': 'center'}))
        
        row += 1
        # Encabezados
        worksheet.write(row, 0, 'Concepto', header_format)
        worksheet.write(row, 1, 'Valor', header_format)
        worksheet.write(row, 2, 'Unidad', header_format)
        worksheet.write(row, 3, 'Notas', header_format)
        
        # Datos de precio y costos
        row += 1
        worksheet.write(row, 0, 'Precio de Venta Unitario', cell_format)
        worksheet.write(row, 1, 100, number_format)
        worksheet.write(row, 2, 'cell_format')
        worksheet.write(row, 3, 'Precio promedio', cell_format)
        precio_venta_row = row + 1
        
        row += 1
        worksheet.write(row, 0, 'Costo Variable Unitario', cell_format)
        worksheet.write(row, 1, 60, number_format)
        worksheet.write(row, 2, 'cell_format')
        worksheet.write(row, 3, 'Costos directos por unidad', cell_format)
        costo_variable_row = row + 1
        
        row += 1
        worksheet.write(row, 0, 'Costos Fijos Mensuales', cell_format)
        worksheet.write(row, 1, 10000, number_format)
        worksheet.write(row, 2, 'cell_format')
        worksheet.write(row, 3, 'Total costos fijos', cell_format)
        costos_fijos_row = row + 1
        
        row += 2
        # Cálculos de Margen de Contribución
        worksheet.merge_range(f'A{row}:D{row}', 'ANÁLISIS DE CONTRIBUCIÓN', workbook.add_format({'bold': True, 'bg_color': '#4472c4', 'font_color': 'white', 'align': 'center'}))
        
        row += 1
        worksheet.write(row, 0, 'Concepto', header_format)
        worksheet.write(row, 1, 'Valor', header_format)
        worksheet.write(row, 2, 'Unidad', header_format)
        worksheet.write(row, 3, 'Fórmula', header_format)
        
        row += 1
        worksheet.write(row, 0, 'Margen de Contribución Unitario', cell_format)
        worksheet.write_formula(row, 1, f'=B{precio_venta_row}-B{costo_variable_row}', formula_format)
        worksheet.write(row, 2, 'cell_format')
        worksheet.write(row, 3, 'PV - CV', cell_format)
        mcu_row = row + 1
        
        row += 1
        worksheet.write(row, 0, 'Ratio de Margen de Contribución', cell_format)
        worksheet.write_formula(row, 1, f'=B{mcu_row}/B{precio_venta_row}', workbook.add_format({'border': 1, 'num_format': '0%', 'bg_color': '#e6f2ff'}))
        worksheet.write(row, 2, '%', cell_format)
        worksheet.write(row, 3, 'MCU / PV', cell_format)
        
        row += 2
        # Cálculos de Punto de Equilibrio
        worksheet.merge_range(f'A{row}:D{row}', 'CÁLCULO DEL PUNTO DE EQUILIBRIO', workbook.add_format({'bold': True, 'bg_color': '#4472c4', 'font_color': 'white', 'align': 'center'}))
        
        row += 1
        worksheet.write(row, 0, 'Concepto', header_format)
        worksheet.write(row, 1, 'Valor', header_format)
        worksheet.write(row, 2, 'Unidad', header_format)
        worksheet.write(row, 3, 'Fórmula', header_format)
        
        row += 1
        worksheet.write(row, 0, 'Punto de Equilibrio en Unidades', cell_format)
        worksheet.write_formula(row, 1, f'=B{costos_fijos_row}/B{mcu_row}', formula_format)
        worksheet.write(row, 2, 'unidades', cell_format)
        worksheet.write(row, 3, 'CF / MCU', cell_format)
        pe_unidades_row = row + 1
        
        row += 1
        worksheet.write(row, 0, 'Punto de Equilibrio en Ventas ($)', cell_format)
        worksheet.write_formula(row, 1, f'=B{pe_unidades_row}*B{precio_venta_row}', formula_format)
        worksheet.write(row, 2, 'cell_format')
        worksheet.write(row, 3, 'PE Unid. * PV', cell_format)
        
        row += 2
        # Tabla de Simulación
        worksheet.merge_range(f'A{row}:D{row}', 'SIMULACIÓN DE ESCENARIOS', workbook.add_format({'bold': True, 'bg_color': '#4472c4', 'font_color': 'white', 'align': 'center'}))
        
        row += 1
        # Encabezados de simulación
        worksheet.write(row, 0, 'Unidades Vendidas', header_format)
        worksheet.write(row, 1, 'Ingresos', header_format)
        worksheet.write(row, 2, 'Costos Totales', header_format)
        worksheet.write(row, 3, 'Utilidad', header_format)
        
        # Tabla de simulación para diferentes niveles de ventas
        start_row = row + 1
        for i in range(9):
            row += 1
            # Calculamos porcentajes del punto de equilibrio: 25%, 50%, 75%, 100%, 125%, etc.
            factor = 0.25 * (i + 1)
            
            worksheet.write_formula(row, 0, f'=ROUND(B{pe_unidades_row}*{factor},0)', cell_format)
            worksheet.write_formula(row, 1, f'=A{row+1}*B{precio_venta_row}', number_format)
            worksheet.write_formula(row, 2, f'=B{costos_fijos_row}+(A{row+1}*B{costo_variable_row})', number_format)
            worksheet.write_formula(row, 3, f'=B{row+1}-C{row+1}', formula_format)
        
        # Agregar gráfico de punto de equilibrio
        chart = workbook.add_chart({'type': 'line'})
        
        # Configurar rangos de datos para el gráfico
        chart.add_series({
            'name': 'Ingresos',
            'categories': f'=Punto de Equilibrio!$A${start_row+1}:$A${row+1}',
            'values': f'=Punto de Equilibrio!$B${start_row+1}:$B${row+1}',
            'line': {'color': 'blue', 'width': 2.25},
        })
        
        chart.add_series({
            'name': 'Costos Totales',
            'categories': f'=Punto de Equilibrio!$A${start_row+1}:$A${row+1}',
            'values': f'=Punto de Equilibrio!$C${start_row+1}:$C${row+1}',
            'line': {'color': 'red', 'width': 2.25},
        })
        
        chart.add_series({
            'name': 'Utilidad',
            'categories': f'=Punto de Equilibrio!$A${start_row+1}:$A${row+1}',
            'values': f'=Punto de Equilibrio!$D${start_row+1}:$D${row+1}',
            'line': {'color': 'green', 'width': 2.25},
        })
        
        # Configuración del gráfico
        chart.set_title({'name': 'Análisis de Punto de Equilibrio'})
        chart.set_x_axis({'name': 'Unidades Vendidas'})
        chart.set_y_axis({'name': 'Valor ($)'})
        chart.set_style(11)  # Estilo del gráfico
        
        # Insertar el gráfico en la hoja
        worksheet.insert_chart('F5', chart, {'x_scale': 1.5, 'y_scale': 1.5})
        
        # Instrucciones
        row += 3
        worksheet.merge_range(f'A{row}:D{row}', 'INSTRUCCIONES:', workbook.add_format({'bold': True}))
        row += 1
        instructions = [
            '1. Modifique los datos de entrada en las celdas B6, B7 y B8 según la información de su producto.',
            '2. El punto de equilibrio se calculará automáticamente en unidades y en valor monetario.',
            '3. La tabla de simulación muestra diferentes escenarios de ventas y su impacto en la utilidad.',
            '4. El gráfico ilustra la intersección del punto de equilibrio donde ingresos = costos totales.',
            '5. Para analizar múltiples productos, duplique esta hoja y ajuste los datos para cada uno.',
            '6. Este análisis es una herramienta de planificación y toma de decisiones operativas.'
        ]
        for instruction in instructions:
            worksheet.merge_range(f'A{row}:D{row}', instruction)
            row += 1
    
    elif template_type == "ratios_financieros":
        # Crear hoja de Ratios Financieros
        worksheet = workbook.add_worksheet("Ratios Financieros")
        
        # Configurar anchos de columna
        worksheet.set_column('A:A', 35)
        worksheet.set_column('B:E', 15)
        
        # Título
        worksheet.merge_range('A1:E1', 'ANÁLISIS DE RATIOS FINANCIEROS', header_format)
        worksheet.merge_range('A2:E2', 'Evaluación de Desempeño Financiero', workbook.add_format({'bold': True, 'align': 'center'}))
        
        row = 4
        # Encabezados principales
        worksheet.write(row, 0, 'Ratio', header_format)
        worksheet.write(row, 1, 'Fórmula', header_format)
        worksheet.write(row, 2, 'Valor', header_format)
        worksheet.write(row, 3, 'Industria', header_format)
        worksheet.write(row, 4, 'Interpretación', header_format)
        
        row += 1
        # Sección de Datos de Entrada
        worksheet.merge_range(f'A{row}:E{row}', 'DATOS DE ENTRADA (EN MILES DE $)', workbook.add_format({'bold': True, 'bg_color': '#4472c4', 'font_color': 'white', 'align': 'center'}))
        
        # Datos del Balance
        row += 1
        worksheet.write(row, 0, 'Activo Corriente', cell_format)
        worksheet.write(row, 1, 'Balance General', cell_format)
        worksheet.write(row, 2, 500, number_format)
        worksheet.write(row, 3, '', cell_format)
        worksheet.write(row, 4, 'Ingrese su valor', cell_format)
        activo_corriente_row = row + 1
        
        row += 1
        worksheet.write(row, 0, 'Inventarios', cell_format)
        worksheet.write(row, 1, 'Balance General', cell_format)
        worksheet.write(row, 2, 150, number_format)
        worksheet.write(row, 3, '', cell_format)
        worksheet.write(row, 4, 'Ingrese su valor', cell_format)
        inventarios_row = row + 1
        
        row += 1
        worksheet.write(row, 0, 'Pasivo Corriente', cell_format)
        worksheet.write(row, 1, 'Balance General', cell_format)
        worksheet.write(row, 2, 300, number_format)
        worksheet.write(row, 3, '', cell_format)
        worksheet.write(row, 4, 'Ingrese su valor', cell_format)
        pasivo_corriente_row = row + 1
        
        row += 1
        worksheet.write(row, 0, 'Activo Total', cell_format)
        worksheet.write(row, 1, 'Balance General', cell_format)
        worksheet.write(row, 2, 1200, number_format)
        worksheet.write(row, 3, '', cell_format)
        worksheet.write(row, 4, 'Ingrese su valor', cell_format)
        activo_total_row = row + 1
        
        row += 1
        worksheet.write(row, 0, 'Pasivo Total', cell_format)
        worksheet.write(row, 1, 'Balance General', cell_format)
        worksheet.write(row, 2, 700, number_format)
        worksheet.write(row, 3, '', cell_format)
        worksheet.write(row, 4, 'Ingrese su valor', cell_format)
        pasivo_total_row = row + 1
        
        row += 1
        worksheet.write(row, 0, 'Patrimonio', cell_format)
        worksheet.write(row, 1, 'Balance General', cell_format)
        worksheet.write(row, 2, 500, number_format)
        worksheet.write(row, 3, '', cell_format)
        worksheet.write(row, 4, 'Ingrese su valor', cell_format)
        patrimonio_row = row + 1
        
        # Datos del Estado de Resultados
        row += 1
        worksheet.write(row, 0, 'Ventas Netas', cell_format)
        worksheet.write(row, 1, 'Estado de Resultados', cell_format)
        worksheet.write(row, 2, 1500, number_format)
        worksheet.write(row, 3, '', cell_format)
        worksheet.write(row, 4, 'Ingrese su valor', cell_format)
        ventas_row = row + 1
        
        row += 1
        worksheet.write(row, 0, 'Utilidad Bruta', cell_format)
        worksheet.write(row, 1, 'Estado de Resultados', cell_format)
        worksheet.write(row, 2, 600, number_format)
        worksheet.write(row, 3, '', cell_format)
        worksheet.write(row, 4, 'Ingrese su valor', cell_format)
        utilidad_bruta_row = row + 1
        
        row += 1
        worksheet.write(row, 0, 'Utilidad Operativa', cell_format)
        worksheet.write(row, 1, 'Estado de Resultados', cell_format)
        worksheet.write(row, 2, 300, number_format)
        worksheet.write(row, 3, '', cell_format)
        worksheet.write(row, 4, 'Ingrese su valor', cell_format)
        utilidad_operativa_row = row + 1
        
        row += 1
        worksheet.write(row, 0, 'Utilidad Neta', cell_format)
        worksheet.write(row, 1, 'Estado de Resultados', cell_format)
        worksheet.write(row, 2, 150, number_format)
        worksheet.write(row, 3, '', cell_format)
        worksheet.write(row, 4, 'Ingrese su valor', cell_format)
        utilidad_neta_row = row + 1
        
        # Resultados: Ratios de Liquidez
        row += 2
        worksheet.merge_range(f'A{row}:E{row}', '1. RATIOS DE LIQUIDEZ', workbook.add_format({'bold': True, 'bg_color': '#4472c4', 'font_color': 'white', 'align': 'center'}))
        
        row += 1
        worksheet.write(row, 0, 'Ratio Corriente', cell_format)
        worksheet.write(row, 1, 'Activo Corriente / Pasivo Corriente', cell_format)
        worksheet.write_formula(row, 2, f'=C{activo_corriente_row}/C{pasivo_corriente_row}', 
                               workbook.add_format({'border': 1, 'num_format': '0.00', 'bg_color': '#e6f2ff'}))
        worksheet.write(row, 3, '> 1.5', cell_format)
        worksheet.write_formula(row, 4, f'=IF(C{row+1}>1.5,"Buena liquidez",IF(C{row+1}>1,"Liquidez ajustada","Problemas de liquidez"))', cell_format)
        
        row += 1
        worksheet.write(row, 0, 'Prueba Ácida', cell_format)
        worksheet.write(row, 1, '(Activo Corriente - Inventarios) / Pasivo Corriente', cell_format)
        worksheet.write_formula(row, 2, f'=(C{activo_corriente_row}-C{inventarios_row})/C{pasivo_corriente_row}', 
                               workbook.add_format({'border': 1, 'num_format': '0.00', 'bg_color': '#e6f2ff'}))
        worksheet.write(row, 3, '> 1.0', cell_format)
        worksheet.write_formula(row, 4, f'=IF(C{row+1}>1,"Buena liquidez inmediata","Posible problema de liquidez inmediata")', cell_format)
        
        row += 1
        worksheet.write(row, 0, 'Capital de Trabajo', cell_format)
        worksheet.write(row, 1, 'Activo Corriente - Pasivo Corriente', cell_format)
        worksheet.write_formula(row, 2, f'=C{activo_corriente_row}-C{pasivo_corriente_row}', 
                               workbook.add_format({'border': 1, 'num_format': '#,##0', 'bg_color': '#e6f2ff'}))
        worksheet.write(row, 3, '> 0', cell_format)
        worksheet.write_formula(row, 4, f'=IF(C{row+1}>0,"Capital de trabajo positivo","Capital de trabajo negativo - alerta")', cell_format)
        
        # Resultados: Ratios de Solvencia
        row += 2
        worksheet.merge_range(f'A{row}:E{row}', '2. RATIOS DE SOLVENCIA', workbook.add_format({'bold': True, 'bg_color': '#4472c4', 'font_color': 'white', 'align': 'center'}))
        
        row += 1
        worksheet.write(row, 0, 'Ratio de Endeudamiento', cell_format)
        worksheet.write(row, 1, 'Pasivo Total / Activo Total', cell_format)
        worksheet.write_formula(row, 2, f'=C{pasivo_total_row}/C{activo_total_row}', 
                               workbook.add_format({'border': 1, 'num_format': '0.00%', 'bg_color': '#e6f2ff'}))
        worksheet.write(row, 3, '< 60%', cell_format)
        worksheet.write_formula(row, 4, f'=IF(C{row+1}<0.6,"Nivel de endeudamiento aceptable","Alto nivel de endeudamiento")', cell_format)
        
        row += 1
        worksheet.write(row, 0, 'Ratio de Autonomía', cell_format)
        worksheet.write(row, 1, 'Patrimonio / Activo Total', cell_format)
        worksheet.write_formula(row, 2, f'=C{patrimonio_row}/C{activo_total_row}', 
                               workbook.add_format({'border': 1, 'num_format': '0.00%', 'bg_color': '#e6f2ff'}))
        worksheet.write(row, 3, '> 40%', cell_format)
        worksheet.write_formula(row, 4, f'=IF(C{row+1}>0.4,"Buena autonomía financiera","Baja autonomía financiera")', cell_format)
        
        row += 1
        worksheet.write(row, 0, 'Apalancamiento', cell_format)
        worksheet.write(row, 1, 'Activo Total / Patrimonio', cell_format)
        worksheet.write_formula(row, 2, f'=C{activo_total_row}/C{patrimonio_row}', 
                               workbook.add_format({'border': 1, 'num_format': '0.00', 'bg_color': '#e6f2ff'}))
        worksheet.write(row, 3, '< 2.5', cell_format)
        worksheet.write_formula(row, 4, f'=IF(C{row+1}<2.5,"Apalancamiento moderado","Alto apalancamiento - mayor riesgo")', cell_format)
        
        # Resultados: Ratios de Rentabilidad
        row += 2
        worksheet.merge_range(f'A{row}:E{row}', '3. RATIOS DE RENTABILIDAD', workbook.add_format({'bold': True, 'bg_color': '#4472c4', 'font_color': 'white', 'align': 'center'}))
        
        row += 1
        worksheet.write(row, 0, 'Margen Bruto', cell_format)
        worksheet.write(row, 1, 'Utilidad Bruta / Ventas', cell_format)
        worksheet.write_formula(row, 2, f'=C{utilidad_bruta_row}/C{ventas_row}', 
                               workbook.add_format({'border': 1, 'num_format': '0.00%', 'bg_color': '#e6f2ff'}))
        worksheet.write(row, 3, '> 30%', cell_format)
        worksheet.write_formula(row, 4, f'=IF(C{row+1}>0.3,"Buen margen bruto","Margen bruto ajustado")', cell_format)
        
        row += 1
        worksheet.write(row, 0, 'Margen Operativo', cell_format)
        worksheet.write(row, 1, 'Utilidad Operativa / Ventas', cell_format)
        worksheet.write_formula(row, 2, f'=C{utilidad_operativa_row}/C{ventas_row}', 
                               workbook.add_format({'border': 1, 'num_format': '0.00%', 'bg_color': '#e6f2ff'}))
        worksheet.write(row, 3, '> 15%', cell_format)
        worksheet.write_formula(row, 4, f'=IF(C{row+1}>0.15,"Buen margen operativo","Margen operativo por mejorar")', cell_format)
        
        row += 1
        worksheet.write(row, 0, 'Margen Neto', cell_format)
        worksheet.write(row, 1, 'Utilidad Neta / Ventas', cell_format)
        worksheet.write_formula(row, 2, f'=C{utilidad_neta_row}/C{ventas_row}', 
                               workbook.add_format({'border': 1, 'num_format': '0.00%', 'bg_color': '#e6f2ff'}))
        worksheet.write(row, 3, '> 10%', cell_format)
        worksheet.write_formula(row, 4, f'=IF(C{row+1}>0.1,"Buen margen neto","Margen neto por mejorar")', cell_format)
        
        row += 1
        worksheet.write(row, 0, 'ROA (Return on Assets)', cell_format)
        worksheet.write(row, 1, 'Utilidad Neta / Activo Total', cell_format)
        worksheet.write_formula(row, 2, f'=C{utilidad_neta_row}/C{activo_total_row}', 
                               workbook.add_format({'border': 1, 'num_format': '0.00%', 'bg_color': '#e6f2ff'}))
        worksheet.write(row, 3, '> 5%', cell_format)
        worksheet.write_formula(row, 4, f'=IF(C{row+1}>0.05,"Buena rentabilidad de activos","Rentabilidad de activos por mejorar")', cell_format)
        
        row += 1
        worksheet.write(row, 0, 'ROE (Return on Equity)', cell_format)
        worksheet.write(row, 1, 'Utilidad Neta / Patrimonio', cell_format)
        worksheet.write_formula(row, 2, f'=C{utilidad_neta_row}/C{patrimonio_row}', 
                               workbook.add_format({'border': 1, 'num_format': '0.00%', 'bg_color': '#e6f2ff'}))
        worksheet.write(row, 3, '> 15%', cell_format)
        worksheet.write_formula(row, 4, f'=IF(C{row+1}>0.15,"Buena rentabilidad para accionistas","Rentabilidad para accionistas por mejorar")', cell_format)
        
        # Instrucciones
        row += 3
        worksheet.merge_range(f'A{row}:E{row}', 'INSTRUCCIONES:', workbook.add_format({'bold': True}))
        row += 1
        instructions = [
            '1. Actualice los datos de entrada de su empresa en las celdas C6 a C17.',
            '2. Los ratios financieros se calcularán automáticamente y serán comparados con valores de referencia.',
            '3. La columna "Industria" muestra valores de referencia generales, ajústelos según su sector específico.',
            '4. La interpretación automática proporciona una guía básica, pero debe ser complementada con análisis específico.',
            '5. Compare estos ratios con períodos anteriores para identificar tendencias en el desempeño financiero.',
            '6. Esta herramienta es orientativa y debe ser validada por un profesional financiero.'
        ]
        for instruction in instructions:
            worksheet.merge_range(f'A{row}:E{row}', instruction)
            row += 1
    
    # Finalizar y guardar el archivo Excel en memoria
    workbook.close()
    output.seek(0)
    
    # Guardar la plantilla en el diccionario para uso futuro
    with excel_templates_lock:
        excel_templates[template_type] = {
            'data': output.getvalue(),
            'timestamp': datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            'name': f"{template_type}_{datetime.now().strftime('%Y%m%d')}.xlsx"
        }
    
    return output.getvalue()

def generate_financial_report():
    """Generar informe financiero general con recomendaciones"""
    logger.info("Generando informe financiero con análisis y recomendaciones...")
    
    report_request = f"""
    Como ContaFin, el agente contable-financiero creado por Innovación Financiera, genera un informe financiero con fecha {datetime.now().strftime('%d-%m-%Y')}.
    
    Enfócate en:
    1. Análisis de tendencias recientes en finanzas para PYMEs.
    2. Mejores prácticas contables actuales para pequeñas empresas.
    3. Recomendaciones para mejorar la salud financiera de una empresa típica.
    4. Consideraciones fiscales importantes para el período actual.
    5. Herramientas y tecnologías financieras recomendadas para optimizar procesos.
    
    Usa el formato detallado en tus instrucciones, con la identidad y valores de Innovación Financiera.
    """
    
    # ID de sesión específico para informes automáticos
    session_id = "financial_report_auto"
    
    try:
        # Primero intentar con el endpoint de chat
        report = call_ollama_api(report_request, session_id)
        
        # Si la respuesta está vacía, intentar con completion
        if not report or report.strip() == "":
            logger.info("El endpoint de chat no devolvió respuesta, probando con completion...")
            report = call_ollama_completion(report_request, session_id)
    except Exception as e:
        logger.error(f"Error al generar informe financiero: {e}")
        report = f"Error al generar informe financiero: {str(e)}"
    
    # Almacenar el informe generado
    with financial_reports_lock:
        financial_reports.append({
            "date": datetime.now().strftime("%Y-%m-%d"),
            "content": report
        })
        # Mantener solo los últimos 30 informes
        if len(financial_reports) > 30:
            financial_reports.pop(0)
    
    logger.info("Informe financiero generado correctamente.")
    return report

def schedule_financial_reports():
    """Configurar la generación periódica de informes financieros"""
    # Generar informe todos los días a las 8:00 AM
    schedule.every().day.at("08:00").do(generate_financial_report)
    
    # También generar informe semanal más extenso los lunes
    schedule.every().monday.at("09:00").do(generate_financial_report)
    
    while True:
        schedule.run_pending()
        time.sleep(60)  # Comprobar cada minuto

@app.route('/')
def home():
    """Ruta de bienvenida básica con información de ContaFin"""
    return jsonify({
        "message": "ContaFin - Agente Contable-Financiero para PYMEs",
        "description": "Asistente especializado en contabilidad, finanzas operativas y cumplimiento fiscal",
        "company": "Innovación Financiera - Expertos en Soluciones Contables y Financieras",
        "status": "online",
        "last_update": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "endpoints": {
            "/chat": "POST - Interactuar con ContaFin mediante mensajes",
            "/report": "GET - Obtener el último informe financiero",
            "/reports": "GET - Listar todos los informes disponibles (últimos 30 días)",
            "/generate-report": "POST - Solicitar un nuevo análisis financiero",
            "/excel-template": "GET - Obtener plantilla Excel según tipo solicitado",
            "/excel-templates": "GET - Listar todas las plantillas Excel disponibles",
            "/reset": "POST - Reiniciar una sesión de conversación",
            "/health": "GET - Verificar estado del servicio"
        },
        "templates_available": [
            "flujo_caja", "nomina", "balance_general", "estado_resultados", 
            "punto_equilibrio", "ratios_financieros"
        ],
        "contact": "Para más información, visite www.innovacionfinanciera.com"
    })

@app.route('/chat', methods=['POST'])
def chat():
    """Endpoint para interactuar con el agente"""
    data = request.json
    
    if not data or 'message' not in data:
        return jsonify({"error": "Se requiere un 'message' en el JSON"}), 400
    
    # Obtener mensaje y session_id (crear uno nuevo si no se proporciona)
    message = data.get('message')
    session_id = data.get('session_id', 'default')
    
    # Inicializar la sesión si es nueva
    with sessions_lock:
        if session_id not in sessions:
            sessions[session_id] = []
    
    # Obtener respuesta del asistente 
    try:
        # Primero intentar con el endpoint de chat
        response = call_ollama_api(message, session_id)
        
        # Si la respuesta está vacía, intentar con completion
        if not response or response.strip() == "":
            logger.info("El endpoint de chat no devolvió una respuesta, probando con completion...")
            response = call_ollama_completion(message, session_id)
    except Exception as e:
        logger.error(f"Error al obtener respuesta: {e}")
        logger.info("Probando con endpoint de completion alternativo...")
        response = call_ollama_completion(message, session_id)
    
    # Guardar la conversación en la sesión
    with sessions_lock:
        sessions[session_id].append({"role": "user", "content": message})
        sessions[session_id].append({"role": "assistant", "content": response})
    
    return jsonify({
        "response": response,
        "session_id": session_id
    })

@app.route('/report', methods=['GET'])
def get_latest_report():
    """Obtener el informe financiero más reciente"""
    with financial_reports_lock:
        if not financial_reports:
            # Si no hay informes, generar uno
            report = generate_financial_report()
            return jsonify({
                "date": datetime.now().strftime("%Y-%m-%d"),
                "content": report
            })
        else:
            # Devolver el informe más reciente
            return jsonify(financial_reports[-1])

@app.route('/reports', methods=['GET'])
def list_reports():
    """Listar todos los informes disponibles"""
    with financial_reports_lock:
        return jsonify({
            "count": len(financial_reports),
            "reports": financial_reports
        })

@app.route('/generate-report', methods=['POST'])
def force_report_generation():
    """Forzar la generación de un nuevo informe financiero"""
    report = generate_financial_report()
    return jsonify({
        "message": "Informe financiero generado correctamente",
        "date": datetime.now().strftime("%Y-%m-%d"),
        "content": report
    })

@app.route('/excel-template', methods=['GET'])
def get_excel_template():
    """Obtener plantilla Excel según el tipo solicitado"""
    template_type = request.args.get('type', 'flujo_caja')
    
    # Verificar si el tipo es válido
    valid_types = ["flujo_caja", "nomina", "balance_general", "estado_resultados", 
                  "punto_equilibrio", "ratios_financieros"]
    
    if template_type not in valid_types:
        return jsonify({
            "error": f"Tipo de plantilla no válido. Opciones disponibles: {', '.join(valid_types)}"
        }), 400
    
    # Verificar si la plantilla ya existe
    with excel_templates_lock:
        if template_type in excel_templates:
            # Usar la plantilla existente
            excel_data = excel_templates[template_type]['data']
            filename = excel_templates[template_type]['name']
        else:
            # Generar nueva plantilla
            excel_data = generate_excel_template(template_type)
            filename = f"{template_type}_{datetime.now().strftime('%Y%m%d')}.xlsx"
    
    # Enviar el archivo Excel como respuesta
    return send_file(
        io.BytesIO(excel_data),
        as_attachment=True,
        download_name=filename,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )

@app.route('/excel-templates', methods=['GET'])
def list_templates():
    """Listar todas las plantillas Excel disponibles"""
    with excel_templates_lock:
        templates_info = {}
        for key, value in excel_templates.items():
            templates_info[key] = {
                "name": value['name'],
                "timestamp": value['timestamp'],
                "url": f"/excel-template?type={key}"
            }
        
        # Añadir plantillas disponibles que aún no se han generado
        valid_types = ["flujo_caja", "nomina", "balance_general", "estado_resultados", 
                      "punto_equilibrio", "ratios_financieros"]
        
        for template_type in valid_types:
            if template_type not in templates_info:
                templates_info[template_type] = {
                    "name": f"{template_type}.xlsx",
                    "timestamp": "No generada aún",
                    "url": f"/excel-template?type={template_type}"
                }
        
        return jsonify({
            "count": len(templates_info),
            "templates": templates_info
        })

@app.route('/reset', methods=['POST'])
def reset_session():
    """Reiniciar una sesión de conversación"""
    data = request.json or {}
    session_id = data.get('session_id', 'default')
    
    with sessions_lock:
        if session_id in sessions:
            sessions[session_id] = []
            message = f"Sesión {session_id} reiniciada correctamente"
        else:
            message = f"La sesión {session_id} no existía, se ha creado una nueva"
            sessions[session_id] = []
    
    return jsonify({"message": message, "session_id": session_id})

@app.route('/health', methods=['GET'])
def health_check():
    """Verificar estado del servicio con información adicional"""
    # Calcular tiempo desde el último informe generado
    last_report_time = None
    with financial_reports_lock:
        if financial_reports:
            try:
                last_report_time = datetime.strptime(financial_reports[-1]["date"], "%Y-%m-%d")
                time_since_last = (datetime.now() - last_report_time).total_seconds() / 3600
            except Exception as e:
                time_since_last = None
        else:
            time_since_last = None
    
    # Contar plantillas Excel generadas
    with excel_templates_lock:
        templates_count = len(excel_templates)
    
    return jsonify({
        "status": "ok",
        "service_name": "ContaFin - Agente Contable-Financiero",
        "provider": "Innovación Financiera",
        "model": MODEL_NAME,
        "ollama_url": OLLAMA_URL,
        "reports_count": len(financial_reports),
        "templates_count": templates_count,
        "last_report_age_hours": time_since_last if time_since_last is not None else "N/A",
        "next_scheduled_report": "8:00 AM (diario) y 9:00 AM (informe semanal los lunes)",
        "version": "1.0.0",
        "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    })

if __name__ == '__main__':
    # Iniciar el planificador de informes financieros en un hilo separado
    scheduler_thread = Thread(target=schedule_financial_reports)
    scheduler_thread.daemon = True
    scheduler_thread.start()
    
    # Generar un informe inicial al iniciar la aplicación
    initial_report_thread = Thread(target=generate_financial_report)
    initial_report_thread.daemon = True
    initial_report_thread.start()
    
    # Generar plantillas Excel iniciales en un hilo separado
    def generate_initial_templates():
        template_types = ["flujo_caja", "nomina"]
        for template_type in template_types:
            try:
                logger.info(f"Generando plantilla inicial de {template_type}...")
                generate_excel_template(template_type)
            except Exception as e:
                logger.error(f"Error al generar plantilla {template_type}: {e}")
    
    templates_thread = Thread(target=generate_initial_templates)
    templates_thread.daemon = True
    templates_thread.start()
    
    # Obtener puerto de variables de entorno (para Render)
    port = int(os.environ.get("PORT", 5000))
    
    # Iniciar la aplicación Flask
    app.run(host='0.0.0.0', port=port, debug=False)
    