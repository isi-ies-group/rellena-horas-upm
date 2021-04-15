# -*- coding: utf-8 -*-
"""
Created on Tue Apr 13 14:09:09 2021

@author: Ruben
"""
import openpyxl
import openpyxl.utils
from datetime import datetime
import numpy as np

# %% ENTRADAS
nombre_archivo = 'fichero_prueba.xlsx'

# horas anuales por tarea
HORAS_TASK = np.array([255, 80, 70, 65, 30, 0])

# meses en los cueles se imputan horas a las tareas
MESES_IMPUTAR = [False, True, True, True, True, True,
                 True, False, True, True, True, True]

# fijar la semilla del random permite reproducir siempre el mismo patrón aleatorio
np.random.seed(123)

# %% CONSTANTES
COLOR_FESTIVO = 52

HORAS_DIA_MAX = 7.5
MINIMA_CARGA_HORARIA = 0.5

COL_INI_TAREA = 'B'
COL_FECHA = 'A'
FILA_CABECERAS_TAREAS = '3'
FILA_INI_TAREAS = '4'

wb = openpyxl.load_workbook(nombre_archivo)

hoja = wb[wb.sheetnames[0]]

NUM_FILAS = len(hoja[COL_INI_TAREA])

for celda in hoja[FILA_CABECERAS_TAREAS]:
    if celda.value == 'Otras actividades':
        col_otras_idx = celda.col_idx

COL_OTRAS_TAREAS = openpyxl.utils.get_column_letter(col_otras_idx+2)

COL_FIN_TAREA = openpyxl.utils.get_column_letter(col_otras_idx-1)

# %% Lee los días laborables
lista_dias = []
lista_filas = []
lista_horas_otros_proyectos = []

for dia in hoja[f'{COL_INI_TAREA}{FILA_INI_TAREAS}:B{NUM_FILAS}']:
    for celda in dia:
        celda_fecha = hoja[COL_FECHA+str(celda.row)]
        if celda.fill.fgColor.indexed != COLOR_FESTIVO and celda_fecha.value is not None:
            if MESES_IMPUTAR[datetime.strptime(celda_fecha.value, '%d/%m/%Y').date().month - 1]:
                lista_filas.append(celda_fecha.row)
                lista_dias.append(celda_fecha.value)
                lista_horas_otros_proyectos.append(
                    hoja[f'{COL_OTRAS_TAREAS}{celda_fecha.row}'].value)

if len(lista_horas_otros_proyectos)*HORAS_DIA_MAX - sum(lista_horas_otros_proyectos) < sum(HORAS_TASK):
    raise SystemError('Más horas en tareas que horas disponibles!')

# %% Inicializa a 0 todas las tareas de todos los días laborables
offset_columna_tarea = openpyxl.utils.column_index_from_string(COL_INI_TAREA)

for celda_dia in hoja[f'{COL_INI_TAREA}{FILA_INI_TAREAS}:B{NUM_FILAS}']:
    dia = celda_dia[0].row
    for tarea in range(offset_columna_tarea, col_otras_idx):
        hoja[f'{openpyxl.utils.get_column_letter(tarea)}{dia}'].value = 0

# %% Rellena la hoja con horas en las tareas
horas_task_ahora = HORAS_TASK.copy()

lista_filas_random = np.array(lista_filas)
# OJO hace el shuffle in_place!
np.random.shuffle(lista_filas_random)

while horas_task_ahora.sum() > 0:
    for dia in lista_filas_random:
        for tarea in range(offset_columna_tarea, col_otras_idx):

            horas_proyectos_ahora = hoja[f'{COL_INI_TAREA}{dia}:{COL_FIN_TAREA}{dia}']
            total_proyectos_ahora = sum(
                [c.value for c in list(horas_proyectos_ahora[0])])

            total_otros_proyectos = hoja[f'{COL_OTRAS_TAREAS}{dia}'].value

            horas_libres_dia = HORAS_DIA_MAX - total_proyectos_ahora - total_otros_proyectos

            if horas_libres_dia > 0 and horas_task_ahora[tarea-offset_columna_tarea] > 0:
                hoja[f'{openpyxl.utils.get_column_letter(tarea)}{dia}'].value += MINIMA_CARGA_HORARIA
                horas_task_ahora[tarea -
                                 offset_columna_tarea] -= MINIMA_CARGA_HORARIA

# %% Guarda en copia
nombre_partido = nombre_archivo.split('.')
nombre_archivo_relleno = nombre_partido[0] + \
    '_relleno' + '.' + nombre_partido[1]

wb.save(nombre_archivo_relleno)

print('Generado fichero relleno:', nombre_archivo_relleno)
