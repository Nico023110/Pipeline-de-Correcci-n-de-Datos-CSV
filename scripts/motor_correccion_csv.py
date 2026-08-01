# -*- coding: utf-8 -*-
"""
Motor de Corrección Múltiple de RIPS / Facturación (CSV y JSON)
Adaptado para procesamiento en lote con rutas dinámicas y CLI.

@author: analisisdedatos
"""

import os
import json
import copy
import shutil
import calendar
import argparse
import pandas as pd
import numpy as np

# ------------------------------------------------------------------------
# RUTAS BASE DEL PROYECTO
# ------------------------------------------------------------------------
DIR_SCRIPTS = os.path.dirname(os.path.abspath(__file__))
DIR_PROYECTO = os.path.dirname(DIR_SCRIPTS)
DIR_FUENTE = os.path.join(DIR_PROYECTO, 'fuente')
DIR_DATA = os.path.join(DIR_PROYECTO, 'data')
DIR_CORREGIDO = os.path.join(DIR_PROYECTO, 'factura_corregida')
DIR_LOGS = os.path.join(DIR_PROYECTO, 'log_errores')
DIR_DESCARGAS = os.path.join(os.path.expanduser('~'), 'Downloads')

# ------------------------------------------------------------------------
# LOG DE CORRECCIONES
# ------------------------------------------------------------------------
log_correcciones = []


def registrar_correccion(
    df,
    idx,
    funcion,
    columna,
    valor_anterior,
    valor_nuevo
):
    log_correcciones.append({
        'funcion': funcion,
        'tipoDocumentoIdentificacion': (
            df.at[idx, 'tipoDocumentoIdentificacion']
            if 'tipoDocumentoIdentificacion' in df.columns
            else ''
        ),
        'numDocumentoIdentificacion': (
            df.at[idx, 'numDocumentoIdentificacion']
            if 'numDocumentoIdentificacion' in df.columns
            else ''
        ),
        'tipoRegistro': (
            df.at[idx, 'tipoRegistro']
            if 'tipoRegistro' in df.columns
            else ''
        ),
        'fechaInicioAtencion': (
            df.at[idx, 'fechaInicioAtencion']
            if 'fechaInicioAtencion' in df.columns
            else ''
        ),
        'fechaSuministroTecnologia': (
            df.at[idx, 'fechaSuministroTecnologia']
            if 'fechaSuministroTecnologia' in df.columns
            else ''
        ),
        'fechaDispensAdmon': (
            df.at[idx, 'fechaDispensAdmon']
            if 'fechaDispensAdmon' in df.columns
            else ''
        ),
        'columna': columna,
        'valor_anterior': valor_anterior,
        'valor_nuevo': valor_nuevo
    })


# ------------------------------------------------------------------------
# FUNCIONES DE CARGA Y DESESTRUCTURACIÓN (CSV / JSON)
# ------------------------------------------------------------------------

def cargar_archivo_csv(ruta):
    """Carga y parsea un archivo RIPS CSV delimitado por punto y coma."""
    if not os.path.exists(ruta):
        raise FileNotFoundError(f"No se encontró el archivo en la ruta: {ruta}")

    with open(ruta, encoding='utf-8', errors='ignore') as f:
        lineas = f.read().splitlines()

    df_raw = pd.DataFrame(lineas)
    df = df_raw[0].str.split(';', expand=True)

    # Tomar la primera fila como nombres de columnas
    df.columns = df.iloc[0].astype(str).str.replace('"', '', regex=False)

    # Eliminar la fila de encabezados
    df = df.iloc[1:].reset_index(drop=True)

    # Quitar comillas de los datos
    df = df.replace(r'^"(.*)"$', r'\1', regex=True)

    return df


def cargar_json_rips(ruta):
    """
    Convierte un JSON RIPS en un DataFrame plano sin perder la
    estructura original para permitir la posterior reconstrucción.
    """
    with open(ruta, "r", encoding="utf-8-sig") as f:
        json_original = json.load(f)

    registros = []

    for i_usuario, usuario in enumerate(json_original.get("usuarios", [])):
        datos_usuario = usuario.copy()
        servicios = datos_usuario.pop("servicios", {})

        for tipo_servicio, lista_registros in servicios.items():
            if not isinstance(lista_registros, list):
                continue

            for i_registro, registro in enumerate(lista_registros):
                fila = {}

                # Datos generales factura
                fila["numDocumentoIdObligado"] = json_original.get("numDocumentoIdObligado")
                fila["numFactura"] = json_original.get("numFactura")
                fila["tipoNota"] = json_original.get("tipoNota")
                fila["numNota"] = json_original.get("numNota")

                # Datos usuario
                fila.update(datos_usuario)

                # Datos servicio
                fila.update(registro)

                # Información para reconstrucción
                fila["_usuario"] = i_usuario
                fila["_tipoServicio"] = tipo_servicio
                fila["_registro"] = i_registro

                registros.append(fila)

    df = pd.DataFrame(registros)
    return json_original, df


def reconstruir_json_rips(json_original, df):
    """Reconstruye el JSON RIPS original aplicando los datos modificados del DataFrame."""
    nuevo_json = copy.deepcopy(json_original)

    for _, fila in df.iterrows():
        usuario = int(fila["_usuario"])
        tipo = fila["_tipoServicio"]
        registro = int(fila["_registro"])

        destino = nuevo_json["usuarios"][usuario]["servicios"][tipo][registro]

        for columna in df.columns:
            if columna.startswith("_"):
                continue

            if columna in [
                "numDocumentoIdObligado",
                "numFactura",
                "tipoNota",
                "numNota"
            ]:
                continue

            if columna in destino:
                destino[columna] = fila[columna]

    return nuevo_json


# ------------------------------------------------------------------------
# REGLAS DE CORRECCIÓN
# ------------------------------------------------------------------------

def corregir_causa_motivo_40(df):
    """Cambia a causaMotivoAtencion 40 las atenciones de finalidad 11 y causa 38."""
    columnas_requeridas = [
        'tipoRegistro',
        'finalidadTecnologiaSalud',
        'causaMotivoAtencion'
    ]

    if not all(col in df.columns for col in columnas_requeridas):
        print('No existen todas las columnas requeridas para corregir_causa_motivo_40.')
        return df

    mascara = (
        (df['tipoRegistro'] == 'consultas') &
        (df['finalidadTecnologiaSalud'] == '11') &
        (df['causaMotivoAtencion'] == '38')
    )

    cantidad = mascara.sum()

    for idx in df.index[mascara]:
        registrar_correccion(
            df,
            idx,
            'corregir_causa_motivo_40',
            'causaMotivoAtencion',
            df.at[idx, 'causaMotivoAtencion'],
            '40'
        )
        df.at[idx, 'causaMotivoAtencion'] = '40'

    print('\n' + '-' * 70)
    print('FUNCIÓN: corregir_causa_motivo_40')
    print('-' * 70)
    print(f'Registros encontrados: {cantidad}')

    return df


def corregir_codtecnologia_sin_ceros(df):
    """Normaliza codTecnologiaSalud eliminando ceros no significativos a la izquierda."""
    print('\n' + '-' * 70)
    print('FUNCIÓN: corregir_codtecnologia_sin_ceros')
    print('-' * 70)

    if 'codTecnologiaSalud' not in df.columns:
        print('No existe la columna codTecnologiaSalud')
        return df

    def limpiar_codigo(valor):
        if pd.isna(valor):
            return valor

        valor = str(valor)

        # Separar parte antes y después del guion si existe
        if '-' in valor:
            parte1, parte2 = valor.split('-', 1)
            parte1 = str(int(parte1)) if parte1.isdigit() else parte1
            parte2 = str(int(parte2)) if parte2.isdigit() else parte2
            return f"{parte1}-{parte2}"
        else:
            if valor.isdigit():
                return str(int(valor))

        return valor

    mascara = df['codTecnologiaSalud'].notna() & (df['codTecnologiaSalud'] != '')
    antes = df.loc[mascara, 'codTecnologiaSalud'].copy()

    df.loc[mascara, 'codTecnologiaSalud'] = df.loc[mascara, 'codTecnologiaSalud'].apply(limpiar_codigo)
    despues = df.loc[mascara, 'codTecnologiaSalud']

    for idx in despues.index:
        if antes.loc[idx] != despues.loc[idx]:
            registrar_correccion(
                df,
                idx,
                'corregir_codtecnologia_sin_ceros',
                'codTecnologiaSalud',
                antes.loc[idx],
                despues.loc[idx]
            )

    total_modificados = (antes != despues).sum()
    print(f'Registros corregidos: {total_modificados}')
    print('FINALIZADO: corregir_codtecnologia_sin_ceros')
    print('-' * 70)

    return df


def corregir_ium_por_cum(df, ruta_equivalencias=None):
    """Reemplaza IUM por CUM utilizando un Excel con columnas IUM y CUM."""
    columnas_requeridas = ['codTecnologiaSalud']
    if not all(col in df.columns for col in columnas_requeridas):
        print('No existen todas las columnas requeridas para corregir_ium_por_cum.')
        return df

    if ruta_equivalencias is None:
        ruta_equivalencias = os.path.join(DIR_DATA, 'equivalencias_ium_cum.xlsx')

    if not os.path.exists(ruta_equivalencias):
        print(f'No se encontró el archivo: {ruta_equivalencias}')
        return df

    equivalencias = pd.read_excel(ruta_equivalencias, dtype=str)
    columnas_excel = ['IUM', 'CUM']

    if not all(col in equivalencias.columns for col in columnas_excel):
        print('El Excel debe contener las columnas IUM y CUM.')
        return df

    equivalencias_ium_cum = dict(zip(equivalencias['IUM'], equivalencias['CUM']))
    mascara = df['codTecnologiaSalud'].isin(equivalencias_ium_cum.keys())
    cantidad = mascara.sum()

    print('\n' + '-' * 70)
    print('FUNCIÓN: corregir_ium_por_cum')
    print('-' * 70)
    print(f'Registros encontrados: {cantidad}')

    for idx in df.index[mascara]:
        valor_anterior = df.at[idx, 'codTecnologiaSalud']
        valor_nuevo = equivalencias_ium_cum.get(valor_anterior, valor_anterior)

        registrar_correccion(
            df,
            idx,
            'corregir_ium_por_cum',
            'codTecnologiaSalud',
            valor_anterior,
            valor_nuevo
        )

        df.at[idx, 'codTecnologiaSalud'] = valor_nuevo

    print(f'Registros modificados: {cantidad}')
    return df


def reubicar_fechas_mes_objetivo(df, anio_objetivo, mes_objetivo):
    """Normaliza y reubica las fechas del archivo al año y mes objetivo especificado."""
    print('\n' + '-' * 70)
    print('FUNCIÓN: reubicar_fechas_mes_objetivo')
    print('-' * 70)

    columnas_requeridas = ['fechaInicioAtencion', 'fechaEgreso']
    if not all(col in df.columns for col in columnas_requeridas):
        print('No existen las columnas requeridas (fechaInicioAtencion, fechaEgreso).')
        return df

    total_inicio_egreso = 0
    total_suministro = 0
    total_dispensacion = 0

    ultimo_dia_mes = calendar.monthrange(anio_objetivo, mes_objetivo)[1]

    inicio_original = pd.to_datetime(df['fechaInicioAtencion'], errors='coerce')
    egreso_original = pd.to_datetime(df['fechaEgreso'], errors='coerce')

    mascara_hospitalizacion = inicio_original.notna() & egreso_original.notna()

    for idx in df.index[mascara_hospitalizacion]:
        fecha_inicio = inicio_original.loc[idx]
        fecha_egreso = egreso_original.loc[idx]

        diferencia_horas = (fecha_egreso - fecha_inicio).total_seconds() / 3600

        ya_en_mes_objetivo = (
            fecha_inicio.year == anio_objetivo and
            fecha_inicio.month == mes_objetivo and
            fecha_egreso.year == anio_objetivo and
            fecha_egreso.month == mes_objetivo
        )

        if ya_en_mes_objetivo and diferencia_horas < 8:
            nueva_fecha_egreso = fecha_inicio + pd.Timedelta(hours=8)

            registrar_correccion(
                df,
                idx,
                'reubicar_fechas_mes_objetivo',
                'fechaEgreso',
                df.at[idx, 'fechaEgreso'],
                nueva_fecha_egreso.strftime('%Y-%m-%d %H:%M')
            )

            df.at[idx, 'fechaEgreso'] = nueva_fecha_egreso.strftime('%Y-%m-%d %H:%M')
            total_inicio_egreso += 1
            continue

        if ya_en_mes_objetivo:
            continue

        duracion_dias = (fecha_egreso.date() - fecha_inicio.date()).days

        nueva_fecha_inicio = pd.Timestamp(
            year=anio_objetivo,
            month=mes_objetivo,
            day=1,
            hour=fecha_inicio.hour,
            minute=fecha_inicio.minute,
            second=fecha_inicio.second
        )

        if diferencia_horas < 8:
            nueva_fecha_egreso = nueva_fecha_inicio + pd.Timedelta(hours=8)
        else:
            dia_egreso = min(1 + max(duracion_dias, 0), ultimo_dia_mes)
            nueva_fecha_egreso = pd.Timestamp(
                year=anio_objetivo,
                month=mes_objetivo,
                day=dia_egreso,
                hour=fecha_egreso.hour,
                minute=fecha_egreso.minute,
                second=fecha_egreso.second
            )

        registrar_correccion(
            df,
            idx,
            'reubicar_fechas_mes_objetivo',
            'fechaInicioAtencion',
            df.at[idx, 'fechaInicioAtencion'],
            nueva_fecha_inicio.strftime('%Y-%m-%d %H:%M')
        )
        df.at[idx, 'fechaInicioAtencion'] = nueva_fecha_inicio.strftime('%Y-%m-%d %H:%M')

        registrar_correccion(
            df,
            idx,
            'reubicar_fechas_mes_objetivo',
            'fechaEgreso',
            df.at[idx, 'fechaEgreso'],
            nueva_fecha_egreso.strftime('%Y-%m-%d %H:%M')
        )
        df.at[idx, 'fechaEgreso'] = nueva_fecha_egreso.strftime('%Y-%m-%d %H:%M')
        total_inicio_egreso += 1

    print(f'Inicio/Egreso corregidos: {total_inicio_egreso}')

    # Fechas de atención sin egreso
    mascara_solo_inicio = inicio_original.notna() & egreso_original.isna()
    total_solo_inicio = 0

    for idx in df.index[mascara_solo_inicio]:
        fecha_inicio = inicio_original.loc[idx]
        ya_en_mes_objetivo = (
            fecha_inicio.year == anio_objetivo and
            fecha_inicio.month == mes_objetivo
        )

        if ya_en_mes_objetivo:
            continue

        nueva_fecha_inicio = pd.Timestamp(
            year=anio_objetivo,
            month=mes_objetivo,
            day=1,
            hour=fecha_inicio.hour,
            minute=fecha_inicio.minute,
            second=fecha_inicio.second
        )

        registrar_correccion(
            df,
            idx,
            'reubicar_fechas_mes_objetivo',
            'fechaInicioAtencion',
            df.at[idx, 'fechaInicioAtencion'],
            nueva_fecha_inicio.strftime('%Y-%m-%d %H:%M')
        )

        df.at[idx, 'fechaInicioAtencion'] = nueva_fecha_inicio.strftime('%Y-%m-%d %H:%M')
        total_solo_inicio += 1

    print(f'Solo fechaInicioAtencion corregidas: {total_solo_inicio}')

    # Fecha suministro tecnología
    if 'fechaSuministroTecnologia' in df.columns:
        fechas = pd.to_datetime(df['fechaSuministroTecnologia'], errors='coerce')
        mascara = fechas.notna()

        for idx in df.index[mascara]:
            fecha = fechas.loc[idx]
            if fecha.year == anio_objetivo and fecha.month == mes_objetivo:
                continue

            nueva_fecha = pd.Timestamp(
                year=anio_objetivo,
                month=mes_objetivo,
                day=1,
                hour=fecha.hour,
                minute=fecha.minute,
                second=fecha.second
            )

            registrar_correccion(
                df,
                idx,
                'reubicar_fechas_mes_objetivo',
                'fechaSuministroTecnologia',
                df.at[idx, 'fechaSuministroTecnologia'],
                nueva_fecha.strftime('%Y-%m-%d %H:%M')
            )
            df.at[idx, 'fechaSuministroTecnologia'] = nueva_fecha.strftime('%Y-%m-%d %H:%M')
            total_suministro += 1

    print(f'fechaSuministroTecnologia corregidas: {total_suministro}')

    # Fecha dispensación / administración
    if 'fechaDispensAdmon' in df.columns:
        fechas = pd.to_datetime(df['fechaDispensAdmon'], errors='coerce')
        mascara = fechas.notna()

        for idx in df.index[mascara]:
            fecha = fechas.loc[idx]
            if fecha.year == anio_objetivo and fecha.month == mes_objetivo:
                continue

            nueva_fecha = pd.Timestamp(
                year=anio_objetivo,
                month=mes_objetivo,
                day=1,
                hour=fecha.hour,
                minute=fecha.minute,
                second=fecha.second
            )

            registrar_correccion(
                df,
                idx,
                'reubicar_fechas_mes_objetivo',
                'fechaDispensAdmon',
                df.at[idx, 'fechaDispensAdmon'],
                nueva_fecha.strftime('%Y-%m-%d %H:%M')
            )
            df.at[idx, 'fechaDispensAdmon'] = nueva_fecha.strftime('%Y-%m-%d %H:%M')
            total_dispensacion += 1

    print(f'fechaDispensAdmon corregidas: {total_dispensacion}')
    print('-' * 70)

    return df


def corregir_finalidades_por_diagnostico(df, ruta_equivalencias_cups=None):
    """Corrige finalidades para procedimientos con finalidad 11 según el diagnóstico principal (CIE10) y lista de equivalencias CUPS."""
    columnas_requeridas = [
        'tipoRegistro',
        'finalidadTecnologiaSalud',
        'codDiagnosticoPrincipal',
        'codProcedimiento'
    ]

    if not all(col in df.columns for col in columnas_requeridas):
        print('No existen todas las columnas requeridas para corregir_finalidades_por_diagnostico.')
        return df

    total_modificados = 0
    print('\n' + '-' * 70)
    print('FUNCIÓN: corregir_finalidades_por_diagnostico')
    print('-' * 70)

    # Z300 - Z309 -> Finalidad 19
    diagnosticos_z30 = [f'Z30{i}' for i in range(10)]
    mascara_z30 = (
        (df['tipoRegistro'] == 'procedimientos') &
        (df['finalidadTecnologiaSalud'] == '11') &
        (df['codDiagnosticoPrincipal'].isin(diagnosticos_z30))
    )
    cantidad_z30 = mascara_z30.sum()

    for idx in df.index[mascara_z30]:
        registrar_correccion(
            df,
            idx,
            'corregir_finalidades_por_diagnostico',
            'finalidadTecnologiaSalud',
            df.at[idx, 'finalidadTecnologiaSalud'],
            '19'
        )
    df.loc[mascara_z30, 'finalidadTecnologiaSalud'] = '19'
    total_modificados += cantidad_z30
    print(f'Z300-Z309 -> Finalidad 19: {cantidad_z30}')

    # Z310 - Z319 -> Finalidad 22
    diagnosticos_z31 = [f'Z31{i}' for i in range(10)]
    mascara_z31 = (
        (df['tipoRegistro'] == 'procedimientos') &
        (df['finalidadTecnologiaSalud'] == '11') &
        (df['codDiagnosticoPrincipal'].isin(diagnosticos_z31))
    )
    cantidad_z31 = mascara_z31.sum()

    for idx in df.index[mascara_z31]:
        registrar_correccion(
            df,
            idx,
            'corregir_finalidades_por_diagnostico',
            'finalidadTecnologiaSalud',
            df.at[idx, 'finalidadTecnologiaSalud'],
            '22'
        )
    df.loc[mascara_z31, 'finalidadTecnologiaSalud'] = '22'
    total_modificados += cantidad_z31
    print(f'Z310-Z319 -> Finalidad 22: {cantidad_z31}')

    # Z320 - Z369 -> Finalidad 23
    diagnosticos_z32_z36 = [f'Z{i}' for i in range(320, 370)]
    mascara_z32_z36 = (
        (df['tipoRegistro'] == 'procedimientos') &
        (df['finalidadTecnologiaSalud'] == '11') &
        (df['codDiagnosticoPrincipal'].isin(diagnosticos_z32_z36))
    )
    cantidad_z32_z36 = mascara_z32_z36.sum()

    for idx in df.index[mascara_z32_z36]:
        registrar_correccion(
            df,
            idx,
            'corregir_finalidades_por_diagnostico',
            'finalidadTecnologiaSalud',
            df.at[idx, 'finalidadTecnologiaSalud'],
            '23'
        )
    df.loc[mascara_z32_z36, 'finalidadTecnologiaSalud'] = '23'
    total_modificados += cantidad_z32_z36
    print(f'Z320-Z369 -> Finalidad 23: {cantidad_z32_z36}')

    # Procedimiento inicia en 990... -> Finalidad 40
    mascara_proc990 = (
        (df['tipoRegistro'] == 'procedimientos') &
        (df['finalidadTecnologiaSalud'] == '11') &
        (df['codProcedimiento'].fillna('').str.startswith('990'))
    )
    cantidad_proc990 = mascara_proc990.sum()

    for idx in df.index[mascara_proc990]:
        registrar_correccion(
            df,
            idx,
            'corregir_finalidades_por_diagnostico',
            'finalidadTecnologiaSalud',
            df.at[idx, 'finalidadTecnologiaSalud'],
            '40'
        )
    df.loc[mascara_proc990, 'finalidadTecnologiaSalud'] = '40'
    total_modificados += cantidad_proc990
    print(f'Procedimiento 990* -> Finalidad 40: {cantidad_proc990}')

    # Procedimiento inicia en 992... o 997... -> Finalidad 14
    mascara_proc992_997 = (
        (df['tipoRegistro'] == 'procedimientos') &
        (df['finalidadTecnologiaSalud'] == '11') &
        (df['codProcedimiento'].fillna('').str.startswith(('992', '997')))
    )
    cantidad_proc992_997 = mascara_proc992_997.sum()

    for idx in df.index[mascara_proc992_997]:
        registrar_correccion(
            df,
            idx,
            'corregir_finalidades_por_diagnostico',
            'finalidadTecnologiaSalud',
            df.at[idx, 'finalidadTecnologiaSalud'],
            '14'
        )
    df.loc[mascara_proc992_997, 'finalidadTecnologiaSalud'] = '14'
    total_modificados += cantidad_proc992_997
    print(f'Procedimiento 992* o 997* -> Finalidad 14: {cantidad_proc992_997}')

    # Procedimiento inicia en 90.... y diagnóstico inicia en Z... -> Finalidad 15
    mascara_proc90_z = (
        (df['tipoRegistro'] == 'procedimientos') &
        (df['finalidadTecnologiaSalud'] == '11') &
        (df['codProcedimiento'].fillna('').str.startswith('90')) &
        (df['codDiagnosticoPrincipal'].fillna('').str.startswith('Z'))
    )
    cantidad_proc90_z = mascara_proc90_z.sum()

    for idx in df.index[mascara_proc90_z]:
        registrar_correccion(
            df,
            idx,
            'corregir_finalidades_por_diagnostico',
            'finalidadTecnologiaSalud',
            df.at[idx, 'finalidadTecnologiaSalud'],
            '15'
        )
    df.loc[mascara_proc90_z, 'finalidadTecnologiaSalud'] = '15'
    total_modificados += cantidad_proc90_z
    print(f'Procedimiento 90* + Diagnóstico Z* -> Finalidad 15: {cantidad_proc90_z}')

    # Procedimiento inicia en 90.... y diagnóstico NO inicia en Z -> Finalidad 12
    mascara_proc90_no_z = (
        (df['tipoRegistro'] == 'procedimientos') &
        (df['finalidadTecnologiaSalud'] == '11') &
        (df['codProcedimiento'].fillna('').str.startswith('90')) &
        (~df['codDiagnosticoPrincipal'].fillna('').str.startswith('Z')) &
        (df['codDiagnosticoPrincipal'].fillna('').ne(''))
    )
    cantidad_proc90_no_z = mascara_proc90_no_z.sum()

    for idx in df.index[mascara_proc90_no_z]:
        registrar_correccion(
            df,
            idx,
            'corregir_finalidades_por_diagnostico',
            'finalidadTecnologiaSalud',
            df.at[idx, 'finalidadTecnologiaSalud'],
            '12'
        )
    df.loc[mascara_proc90_no_z, 'finalidadTecnologiaSalud'] = '12'
    total_modificados += cantidad_proc90_no_z
    print(f'Procedimiento 90* + Diagnóstico NO Z* -> Finalidad 12: {cantidad_proc90_no_z}')

    # CUPS específicos -> Finalidades desde Excel
    if ruta_equivalencias_cups is None:
        ruta_equivalencias_cups = os.path.join(DIR_DATA, 'equivalencias_cups_finalidad.xlsx')

    if not os.path.exists(ruta_equivalencias_cups):
        print(f'No se encontró el archivo: {ruta_equivalencias_cups}')
    else:
        equivalencias = pd.read_excel(ruta_equivalencias_cups, dtype=str)
        columnas_excel = ['CUPS', 'FINALIDAD']

        if not all(col in equivalencias.columns for col in columnas_excel):
            print('El Excel debe contener las columnas CUPS y FINALIDAD.')
        else:
            equivalencias_cups = dict(
                zip(
                    equivalencias['CUPS'].astype(str).str.upper(),
                    equivalencias['FINALIDAD'].astype(str)
                )
            )
            codigos_proc = df['codProcedimiento'].fillna('').astype(str).str.upper()
            mascara_cups = (
                (df['tipoRegistro'] == 'procedimientos') &
                (df['finalidadTecnologiaSalud'] == '11') &
                (codigos_proc.isin(equivalencias_cups.keys()))
            )

            cantidad = mascara_cups.sum()
            print(f'CUPS encontrados: {cantidad}')

            for idx in df.index[mascara_cups]:
                cups = codigos_proc.loc[idx]
                finalidad = equivalencias_cups[cups]
                registrar_correccion(
                    df,
                    idx,
                    f'CUPS_{cups}',
                    'finalidadTecnologiaSalud',
                    df.at[idx, 'finalidadTecnologiaSalud'],
                    finalidad
                )
                df.at[idx, 'finalidadTecnologiaSalud'] = finalidad

            total_modificados += cantidad

    # Finalidad 25 -> Causa 42
    if 'causaMotivoAtencion' in df.columns:
        mascara_finalidad25 = (
            (df['finalidadTecnologiaSalud'] == '25') &
            (df['causaMotivoAtencion'] != '42')
        )
        cantidad_finalidad25 = mascara_finalidad25.sum()

        for idx in df.index[mascara_finalidad25]:
            registrar_correccion(
                df,
                idx,
                'finalidad25_causa42',
                'causaMotivoAtencion',
                df.at[idx, 'causaMotivoAtencion'],
                '42'
            )
        df.loc[mascara_finalidad25, 'causaMotivoAtencion'] = '42'
        total_modificados += cantidad_finalidad25
        print(f'Finalidad 25 -> Causa 42: {cantidad_finalidad25}')

    print('-' * 70)
    print(f'Total registros modificados: {total_modificados}')
    print('-' * 70)

    return df


def corregir_sexo_por_reglas(df):
    """Ajusta sexo a 'F' en finalidades 23 o diagnósticos específicos como Z124."""
    print('\n' + '-' * 70)
    print('FUNCIÓN: corregir_sexo_por_reglas')
    print('-' * 70)

    columnas_requeridas = ['finalidadTecnologiaSalud', 'codSexo']
    if not all(col in df.columns for col in columnas_requeridas):
        print('No existen todas las columnas requeridas.')
        return df

    mascara = (
        (df['finalidadTecnologiaSalud'].fillna('') == '23') &
        (df['codSexo'].fillna('') != 'F')
    )
    cantidad = mascara.sum()

    for idx in df.index[mascara]:
        registrar_correccion(
            df,
            idx,
            'corregir_sexo_finalidad_23',
            'codSexo',
            df.at[idx, 'codSexo'],
            'F'
        )
    df.loc[mascara, 'codSexo'] = 'F'
    print(f'Registros con finalidad 23 corregidos a sexo F: {cantidad}')

    # DX Z124 -> SEXO F
    if 'codDiagnosticoPrincipal' in df.columns:
        mascara_z124 = (
            (df['codDiagnosticoPrincipal'].fillna('') == 'Z124') &
            (df['codSexo'].fillna('') != 'F')
        )
        cantidad_z124 = mascara_z124.sum()

        for idx in df.index[mascara_z124]:
            registrar_correccion(
                df,
                idx,
                'corregir_sexo_dx_z124',
                'codSexo',
                df.at[idx, 'codSexo'],
                'F'
            )
        df.loc[mascara_z124, 'codSexo'] = 'F'
        print(f'DX Z124 corregidos a sexo F: {cantidad_z124}')

    print('-' * 70)
    return df


def corregir_pais_origen(df):
    """Asigna codPaisOrigen por defecto '170' (Colombia) si está vacío o es nulo."""
    print('\n' + '-' * 70)
    print('FUNCIÓN: corregir_pais_origen')
    print('-' * 70)

    if 'codPaisOrigen' not in df.columns:
        print('No existe la columna codPaisOrigen')
        return df

    mascara = df['codPaisOrigen'].isna() | (df['codPaisOrigen'].astype(str).str.strip().eq(''))
    cantidad = mascara.sum()

    for idx in df.index[mascara]:
        registrar_correccion(
            df,
            idx,
            'corregir_pais_origen',
            'codPaisOrigen',
            df.at[idx, 'codPaisOrigen'],
            '170'
        )

    df.loc[mascara, 'codPaisOrigen'] = '170'
    print(f'Registros corregidos: {cantidad}')
    print('-' * 70)
    return df


def corregir_condicion_destino_egreso(df):
    """Establece condicionDestinoUsuarioEgreso a '01' si es diferente de '01' o no nula."""
    print('\n' + '-' * 70)
    print('FUNCIÓN: corregir_condicion_destino_egreso')
    print('-' * 70)

    if 'condicionDestinoUsuarioEgreso' not in df.columns:
        print('No existe la columna condicionDestinoUsuarioEgreso')
        return df

    val = df['condicionDestinoUsuarioEgreso'].astype(str).str.strip()
    mascara = df['condicionDestinoUsuarioEgreso'].notna() & (val != '') & (val != '01')
    cantidad = mascara.sum()

    for idx in df.index[mascara]:
        registrar_correccion(
            df,
            idx,
            'corregir_condicion_destino_egreso',
            'condicionDestinoUsuarioEgreso',
            df.at[idx, 'condicionDestinoUsuarioEgreso'],
            '01'
        )

    df.loc[mascara, 'condicionDestinoUsuarioEgreso'] = '01'
    print(f'Registros corregidos: {cantidad}')
    print('-' * 70)
    return df


def corregir_diagnosticos_duplicados(df):
    """Reemplaza diagnósticos relacionados repetidos por .null."""
    print('\n' + '-' * 70)
    print('FUNCIÓN: corregir_diagnosticos_duplicados')
    print('-' * 70)

    if 'codDiagnosticoPrincipal' not in df.columns:
        print('No existe la columna codDiagnosticoPrincipal')
        return df

    columnas_relacionadas = [col for col in df.columns if 'codDiagnosticoRelacionado' in col]

    if not columnas_relacionadas:
        print('No se encontraron columnas de diagnósticos relacionados.')
        return df

    total_modificados = 0

    for idx in df.index:
        principal = str(df.at[idx, 'codDiagnosticoPrincipal']).strip()
        if principal in ['', '.null.', 'nan']:
            continue

        diagnosticos_vistos = {principal}

        for columna in columnas_relacionadas:
            valor = df.at[idx, columna]
            if pd.isna(valor):
                continue

            valor = str(valor).strip()
            if valor in ['', '.null.']:
                continue

            if valor in diagnosticos_vistos:
                registrar_correccion(
                    df,
                    idx,
                    'corregir_diagnosticos_duplicados',
                    columna,
                    valor,
                    '.null.'
                )
                df.at[idx, columna] = '.null.'
                total_modificados += 1
            else:
                diagnosticos_vistos.add(valor)

    print(f'Diagnósticos relacionados corregidos: {total_modificados}')
    print('-' * 70)
    return df


def corregir_num_consultas_prenatal(df):
    """Convierte numConsultasCPrenatal a valor entero estándar."""
    print('\n' + '-' * 70)
    print('FUNCIÓN: corregir_num_consultas_prenatal')
    print('-' * 70)

    if 'numConsultasCPrenatal' not in df.columns:
        print('No existe la columna numConsultasCPrenatal')
        return df

    total_modificados = 0

    for idx in df.index:
        valor_original = str(df.at[idx, 'numConsultasCPrenatal']).strip()
        if valor_original in ['', '.null.', 'nan', 'None']:
            continue

        try:
            valor_nuevo = str(int(float(valor_original)))
        except (ValueError, TypeError):
            continue

        if valor_original != valor_nuevo:
            registrar_correccion(
                df,
                idx,
                'corregir_num_consultas_prenatal',
                'numConsultasCPrenatal',
                valor_original,
                valor_nuevo
            )
            df.at[idx, 'numConsultasCPrenatal'] = valor_nuevo
            total_modificados += 1

    print(f'Registros corregidos: {total_modificados}')
    print('-' * 70)
    return df


# ------------------------------------------------------------------------
# EXPORTACIÓN DE RESULTADOS
# ------------------------------------------------------------------------

def guardar_archivo(df, ruta_entrada, json_original=None, dir_corregido=DIR_CORREGIDO, dir_logs=DIR_LOGS):
    """Exporta el archivo corregido (CSV o JSON) y su correspondiente log de auditoría."""
    os.makedirs(dir_corregido, exist_ok=True)
    os.makedirs(dir_logs, exist_ok=True)

    nombre = os.path.basename(ruta_entrada)
    nombre_sin_extension = os.path.splitext(nombre)[0]
    extension = os.path.splitext(nombre)[1].lower()

    if extension == ".csv":
        nombre_salida = f"{nombre_sin_extension}_CORREGIDO.CSV"
        ruta_salida = os.path.join(dir_corregido, nombre_salida)
        df.to_csv(ruta_salida, sep=";", index=False, encoding="utf-8")

    elif extension == ".json":
        nombre_salida = f"{nombre_sin_extension}_CORREGIDO.json"
        ruta_salida = os.path.join(dir_corregido, nombre_salida)
        if json_original is not None:
            json_corregido = reconstruir_json_rips(json_original, df)
            with open(ruta_salida, "w", encoding="utf-8") as f:
                json.dump(json_corregido, f, ensure_ascii=False, indent=4)
        else:
            df.to_json(ruta_salida, orient="records", force_ascii=False, indent=4)

    print(f"\nArchivo corregido exportado a:\n{ruta_salida}")

    if len(log_correcciones) > 0:
        ruta_log = os.path.join(dir_logs, f"LOG_ERRORES_{nombre_sin_extension}.xlsx")
        df_log = pd.DataFrame(log_correcciones)
        try:
            df_log.to_excel(ruta_log, index=False)
            print(f"Log de auditoría exportado a:\n{ruta_log}")
        except PermissionError:
            ruta_log_alt = os.path.join(dir_logs, f"LOG_ERRORES_{nombre_sin_extension}_NUEVO.xlsx")
            df_log.to_excel(ruta_log_alt, index=False)
            print(f"[!] Advertencia: {ruta_log} está abierto/bloqueado. Se exportó log a:\n{ruta_log_alt}")
        except Exception as e:
            print(f"[!] No se pudo exportar el log de auditoría {ruta_log}: {e}")
    else:
        print("No se registraron cambios en el log de auditoría.")

    return ruta_salida


# ------------------------------------------------------------------------
# ANÁLISIS DINÁMICO DE MES Y AÑO OBJETIVO POR DOCUMENTO
# ------------------------------------------------------------------------

def determinar_mes_y_anio_objetivo(df, anio_defecto=2026, mes_defecto=6):
    """
    Analiza las fechas del documento para determinar dinámicamente cuál es el mes
    y año que más se repite (la moda de las fechas del archivo).
    Ese mes detectado en el propio documento se utiliza como el mes_objetivo para la corrección.
    """
    columnas_fechas = [
        'fechaInicioAtencion',
        'fechaDispensAdmon',
        'fechaSuministroTecnologia',
        'fechaEgreso'
    ]

    todas_las_fechas = []
    for col in columnas_fechas:
        if col in df.columns:
            fechas_parsed = pd.to_datetime(df[col], errors='coerce').dropna()
            if not fechas_parsed.empty:
                todas_las_fechas.append(fechas_parsed)

    if todas_las_fechas:
        fechas_concatenadas = pd.concat(todas_las_fechas, ignore_index=True)
        if not fechas_concatenadas.empty:
            mes_mas_repetido = int(fechas_concatenadas.dt.month.mode()[0])
            anio_mas_repetido = int(fechas_concatenadas.dt.year.mode()[0])
            nombres_meses = [
                "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
                "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"
            ]
            nombre_mes = nombres_meses[mes_mas_repetido - 1]
            
            # El mes objetivo es EXACTAMENTE el mes predominante dentro de la factura
            mes_objetivo_calculado = mes_mas_repetido
            anio_objetivo_calculado = anio_mas_repetido

            return anio_objetivo_calculado, mes_objetivo_calculado, f"Detectado {nombre_mes} {anio_mas_repetido} -> Objetivo: {mes_objetivo_calculado}/{anio_objetivo_calculado}"

    return anio_defecto, mes_defecto, f"Por Defecto ({mes_defecto}/{anio_defecto})"


# ------------------------------------------------------------------------
# ORQUESTADOR / PIPELINE MULTIPLE
# ------------------------------------------------------------------------

def procesar_un_archivo(ruta_entrada, anio_objetivo=None, mes_objetivo=None):
    """Ejecuta el pipeline completo de corrección sobre cualquier archivo de factura (CSV o JSON)."""
    global log_correcciones
    log_correcciones.clear()

    nombre_archivo = os.path.basename(ruta_entrada)
    extension = os.path.splitext(ruta_entrada)[1].lower()
    json_original = None

    if extension == ".csv":
        df = cargar_archivo_csv(ruta_entrada)
    elif extension == ".json":
        json_original, df = cargar_json_rips(ruta_entrada)
    else:
        print(f"Formato no soportado: {extension} en {ruta_entrada}")
        return None

    df = df.fillna("")

    if mes_objetivo is None or anio_objetivo is None:
        anio_calc, mes_calc, desc_mes = determinar_mes_y_anio_objetivo(df)
        anio_final = anio_objetivo if anio_objetivo is not None else anio_calc
        mes_final = mes_objetivo if mes_objetivo is not None else mes_calc
    else:
        anio_final = anio_objetivo
        mes_final = mes_objetivo
        desc_mes = f"Manual ({mes_final}/{anio_final})"

    print(f"\n[+] Procesando: {nombre_archivo} | Mes Objetivo: {mes_final} (Año: {anio_final}) [{desc_mes}]")

    # Pipeline de correcciones dinámico
    df = corregir_causa_motivo_40(df)
    df = corregir_codtecnologia_sin_ceros(df)
    df = corregir_ium_por_cum(df)
    #df = reubicar_fechas_mes_objetivo(df, anio_final, mes_final)
    df = corregir_finalidades_por_diagnostico(df)
    df = corregir_sexo_por_reglas(df)
    df = corregir_pais_origen(df)
    df = corregir_condicion_destino_egreso(df)
    df = corregir_diagnosticos_duplicados(df)
    df = corregir_num_consultas_prenatal(df)

    ruta_salida = guardar_archivo(df, ruta_entrada, json_original=json_original)
    return ruta_salida


def buscar_e_importar_facturas(carpetas_origen=None, carpeta_destino=DIR_FUENTE):
    """
    Busca recursivamente cualquier archivo CSV/JSON de factura en las carpetas de origen
    y los copia a la carpeta fuente.
    """
    if carpetas_origen is None:
        carpetas_origen = [DIR_DESCARGAS, DIR_FUENTE]

    os.makedirs(carpeta_destino, exist_ok=True)

    encontrados = 0
    copiados = 0

    print("=" * 80)
    print("BÚSQUEDA Y RECOLECCIÓN DE FACTURAS RIPS (GENERAL)")
    print("=" * 80)

    for ruta_base in carpetas_origen:
        if not os.path.exists(ruta_base):
            continue
        print(f"Escaneando directorio: {ruta_base} ...")
        for root, _, files in os.walk(ruta_base):
            for file in files:
                if file.lower().endswith(('.csv', '.json')):
                    src_path = os.path.join(root, file)
                    dest_path = os.path.join(carpeta_destino, file)
                    encontrados += 1
                    if os.path.abspath(src_path) != os.path.abspath(dest_path):
                        shutil.copy2(src_path, dest_path)
                        copiados += 1
                        print(f"  [+] Factura recolectada: {file} -> {carpeta_destino}")

    print(f"Recolectadas: {encontrados} facturas | Nuevas copiadas a fuente/: {copiados}\n")
    return copiados


def generar_consolidado_log_errores(dir_logs=DIR_LOGS, nombre_salida='CONSOLIDADO_LOG_ERRORES.xlsx'):
    """
    Consolida todos los archivos de auditoría Excel (LOG_ERRORES_*.xlsx)
    en un único archivo maestro de resumen.
    """
    if not os.path.exists(dir_logs):
        print(f"No existe el directorio de logs: {dir_logs}")
        return None

    archivos_log = [
        f for f in os.listdir(dir_logs)
        if f.lower().startswith('log_errores_')
        and f.lower().endswith('.xlsx')
        and not f.startswith('~$')
        and f != nombre_salida
    ]

    if not archivos_log:
        print("No se encontraron archivos de log individuales para consolidar.")
        return None

    dfs = []
    print("\n" + "=" * 80)
    print(f"GENERANDO CONSOLIDADO DE AUDITORÍA DE ERRORES ({len(archivos_log)} ARCHIVOS DE LOG)")
    print("=" * 80)

    for f in archivos_log:
        ruta_file = os.path.join(dir_logs, f)
        try:
            df_temp = pd.read_excel(ruta_file)
            if not df_temp.empty:
                nombre_factura = f.replace("LOG_ERRORES_", "").replace(".xlsx", "")
                df_temp.insert(0, "archivo_log", f)
                df_temp.insert(1, "factura_id", nombre_factura)
                dfs.append(df_temp)
        except Exception as e:
            print(f"  [!] Error leyendo log {f}: {e}")

    if dfs:
        df_consolidado = pd.concat(dfs, ignore_index=True)
        ruta_salida = os.path.join(dir_logs, nombre_salida)
        try:
            df_consolidado.to_excel(ruta_salida, index=False)
            print("=" * 80)
            print(f"CONSOLIDADO GENERADO CON ÉXITO: {len(df_consolidado)} registros totales en:")
            print(f"{ruta_salida}")
            print("=" * 80)
            return ruta_salida
        except PermissionError:
            ruta_salida_alt = os.path.join(dir_logs, f"CONSOLIDADO_LOG_ERRORES_NUEVO.xlsx")
            df_consolidado.to_excel(ruta_salida_alt, index=False)
            print(f"[!] Advertencia: {ruta_salida} está bloqueado. Se guardó consolidado en:\n{ruta_salida_alt}")
            return ruta_salida_alt
    else:
        print("No se encontraron registros en los logs de auditoría para consolidar.")
        return None


def procesar_multiples_archivos(carpeta_fuente=DIR_FUENTE, anio_objetivo=None, mes_objetivo=None, buscar_en_descargas=False):
    """Procesa todos los archivos CSV y JSON presentes exclusivamente en la carpeta fuente."""
    if buscar_en_descargas:
        buscar_e_importar_facturas(carpeta_destino=carpeta_fuente)

    if not os.path.exists(carpeta_fuente):
        raise FileNotFoundError(f"No existe la carpeta fuente especificada: {carpeta_fuente}")

    archivos = [
        os.path.join(carpeta_fuente, f)
        for f in os.listdir(carpeta_fuente)
        if f.lower().endswith(('.csv', '.json'))
    ]

    if not archivos:
        print(f"No se encontraron archivos CSV o JSON en la carpeta fuente: {carpeta_fuente}")
        return

    print("=" * 80)
    print(f"INICIANDO PROCESAMIENTO MÚLTIPLE DE {len(archivos)} ARCHIVOS EN FUENTE")
    print(f"Carpeta Fuente: {carpeta_fuente}")
    if mes_objetivo is not None:
        print(f"Mes Objetivo Fijo: {mes_objetivo}")
    else:
        print("Mes Objetivo: Calculado dinámicamente según el mes predominante de cada documento")
    print("=" * 80)

    procesados_exito = 0

    for ruta in archivos:
        resultado = procesar_un_archivo(ruta, anio_objetivo=anio_objetivo, mes_objetivo=mes_objetivo)
        if resultado:
            procesados_exito += 1

    print("\n" + "=" * 80)
    print(f"PROCESO DE LOTE FINALIZADO CON ÉXITO: {procesados_exito}/{len(archivos)} archivos procesados.")
    print("=" * 80)

    # Generar el consolidado maestro de logs de auditoría
    generar_consolidado_log_errores()


# ------------------------------------------------------------------------
# PUNTO DE ENTRADA CLI / IDE
# ------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(
        description="Motor de Corrección Múltiple de RIPS (CSV / JSON) con Análisis Dinámico de Fechas para Antigravity IDE"
    )
    parser.add_argument(
        '--fuente',
        type=str,
        default=DIR_FUENTE,
        help="Ruta al directorio de entrada con los archivos CSV/JSON (defecto: carpeta fuente/)"
    )
    parser.add_argument(
        '--anio',
        type=int,
        default=None,
        help="Año objetivo explícito para la reubicación de fechas (defecto: se analiza del documento)"
    )
    parser.add_argument(
        '--mes',
        type=int,
        default=None,
        help="Mes objetivo explícito (defecto: se calcula la moda de las fechas del propio documento)"
    )
    parser.add_argument(
        '--buscar_descargas',
        action='store_true',
        help="Si se especifica, busca e importa facturas adicionales desde la carpeta Descargas"
    )

    args = parser.parse_args()

    procesar_multiples_archivos(
        carpeta_fuente=args.fuente,
        anio_objetivo=args.anio,
        mes_objetivo=args.mes,
        buscar_en_descargas=args.buscar_descargas
    )


if __name__ == '__main__':
    main()


