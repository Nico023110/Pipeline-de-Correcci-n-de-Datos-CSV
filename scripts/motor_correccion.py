# -*- coding: utf-8 -*-
"""
Motor de Corrección de RIPS / Facturación CSV
Adaptado para Antigravity IDE / VS Code / Consola Python

@author: analisisdedatos
"""

import os
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
# FUNCIÓN DE CARGA
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
        print('No existen todas las columnas requeridas.')
        print('\nColumnas encontradas:')
        print(df.columns.tolist())
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
        print('No existen todas las columnas requeridas.')
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
        print('No existen las columnas requeridas.')
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
    """Corrige finalidades para procedimientos con finalidad 11 según el diagnóstico principal (CIE10)."""
    columnas_requeridas = [
        'tipoRegistro',
        'finalidadTecnologiaSalud',
        'codDiagnosticoPrincipal',
        'codProcedimiento'
    ]

    if not all(col in df.columns for col in columnas_requeridas):
        print('No existen todas las columnas requeridas.')
        print('\nColumnas encontradas:')
        print(df.columns.tolist())
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
    print('FUNCIÓN: corregir_sexo_finalidad_23')
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
# ORQUESTADOR / PIPELINE PRINCIPAL
# ------------------------------------------------------------------------

def procesar_archivo(ruta_entrada, anio_objetivo=2026, mes_objetivo=6):
    """Ejecuta el pipeline completo de corrección sobre un archivo RIPS CSV."""
    global log_correcciones
    log_correcciones.clear()

    print('=' * 70)
    print(f'INICIANDO PROCESO DE CORRECCIÓN RIPS: {ruta_entrada}')
    print(f'AÑO OBJETIVO: {anio_objetivo} | MES OBJETIVO: {mes_objetivo}')
    print('=' * 70)

    # Asegurar que existan los directorios de salida
    os.makedirs(DIR_CORREGIDO, exist_ok=True)
    os.makedirs(DIR_LOGS, exist_ok=True)

    # 1. Cargar archivo
    df = cargar_archivo_csv(ruta_entrada)

    # 2. Secuencia de correcciones
    df = corregir_causa_motivo_40(df)
    df = corregir_codtecnologia_sin_ceros(df)
    df = corregir_ium_por_cum(df)
    df = reubicar_fechas_mes_objetivo(df, anio_objetivo, mes_objetivo)
    df = corregir_finalidades_por_diagnostico(df)
    df = corregir_sexo_por_reglas(df)
    df = corregir_pais_origen(df)
    df = corregir_condicion_destino_egreso(df)
    df = corregir_diagnosticos_duplicados(df)
    df = corregir_num_consultas_prenatal(df)

    # 3. Exportar archivo corregido
    nombre_base = os.path.splitext(os.path.basename(ruta_entrada))[0]
    nombre_salida = f'{nombre_base}_CORREGIDO.CSV'
    ruta_salida = os.path.join(DIR_CORREGIDO, nombre_salida)

    df.to_csv(ruta_salida, sep=';', index=False, encoding='utf-8')

    # 4. Exportar log de correcciones
    if len(log_correcciones) > 0:
        df_log = pd.DataFrame(log_correcciones)
        ruta_log = os.path.join(DIR_LOGS, f'LOG_ERRORES_{nombre_base}.xlsx')
        df_log.to_excel(ruta_log, index=False)
        print(f'\nLog de auditoría generado correctamente:\n{ruta_log}')
    else:
        print('\nNo se registraron cambios en la auditoría.')

    print('\n' + '=' * 70)
    print('PROCESO FINALIZADO CON ÉXITO')
    print(f'Archivo exportado a:\n{ruta_salida}')
    print('=' * 70)

    return ruta_salida


# ------------------------------------------------------------------------
# PUNTO DE ENTRADA CLI / IDE
# ------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(description="Motor de Corrección de RIPS CSV para Antigravity IDE")
    parser.add_argument(
        '--archivo',
        type=str,
        default=None,
        help="Ruta al archivo CSV de entrada (por defecto busca en la carpeta fuente/)"
    )
    parser.add_argument(
        '--anio',
        type=int,
        default=2026,
        help="Año objetivo para la reubicación de fechas (defecto: 2026)"
    )
    parser.add_argument(
        '--mes',
        type=int,
        default=6,
        help="Mes objetivo para la reubicación de fechas (defecto: 6)"
    )

    args = parser.parse_args()

    ruta_archivo = args.archivo

    # Si no se especifica archivo por línea de comandos, buscar el más reciente o por defecto en fuente/
    if not ruta_archivo:
        if os.path.exists(DIR_FUENTE):
            archivos_csv = [
                os.path.join(DIR_FUENTE, f)
                for f in os.listdir(DIR_FUENTE)
                if f.lower().endswith('.csv')
            ]
            if archivos_csv:
                # Tomar el último modificado
                ruta_archivo = max(archivos_csv, key=os.path.getmtime)

    if not ruta_archivo or not os.path.exists(ruta_archivo):
        # Fallback al archivo estático por defecto
        ruta_archivo = os.path.join(DIR_FUENTE, 'FEV394477.CSV')

    procesar_archivo(ruta_archivo, anio_objetivo=args.anio, mes_objetivo=args.mes)


if __name__ == '__main__':
    main()