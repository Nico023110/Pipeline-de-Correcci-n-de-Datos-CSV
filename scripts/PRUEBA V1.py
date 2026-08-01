# -*- coding: utf-8 -*-
"""
Created on Wed Jun  3 07:43:29 2026

@author: analisisdedatos
"""

import pandas as pd
import numpy as np
import os

#------------------------------------------------------------------------
# LOG DE CORRECCIONES
#------------------------------------------------------------------------

#------------------------------------------------------------------------
# LOG DE CORRECCIONES
#------------------------------------------------------------------------

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
        'tipoDocumentoIdentificacion':
            df.at[idx, 'tipoDocumentoIdentificacion']
            if 'tipoDocumentoIdentificacion' in df.columns
            else '',
        'numDocumentoIdentificacion':
            df.at[idx, 'numDocumentoIdentificacion']
            if 'numDocumentoIdentificacion' in df.columns
            else '',
        'tipoRegistro':
            df.at[idx, 'tipoRegistro']
            if 'tipoRegistro' in df.columns
            else '',
        'fechaInicioAtencion':
            df.at[idx, 'fechaInicioAtencion']
            if 'fechaInicioAtencion' in df.columns
            else '',
        'columna': columna,
        'valor_anterior': valor_anterior,
        'valor_nuevo': valor_nuevo
    })

#------------------------------------------------------------------------

#-------------------------- Cargar archivo -----------------------------

ruta = r'C:\proyecto\FEV394388.CSV'

if not os.path.exists(ruta):
    raise FileNotFoundError(
        f"No se encontró el archivo en la ruta: {ruta}"
    )

# Leer archivo como texto puro
with open(
    ruta,
    encoding='utf-8',
    errors='ignore'
) as f:

    lineas = f.read().splitlines()

df_raw = pd.DataFrame(lineas)

# Separar la única columna por ;
analisis_codigos = df_raw[0].str.split(';', expand=True)

# Tomar la primera fila como nombres de columnas
analisis_codigos.columns = analisis_codigos.iloc[0]

# Quitar comillas de los encabezados
analisis_codigos.columns = (
    analisis_codigos.columns
    .astype(str)
    .str.replace('"', '', regex=False)
)

# Eliminar la fila de encabezados
analisis_codigos = analisis_codigos.iloc[1:].reset_index(drop=True)

# Quitar comillas de los datos
analisis_codigos = analisis_codigos.replace(
    r'^"(.*)"$',
    r'\1',
    regex=True
)

#------------------------------------------------------------------------

# Esta función se encarga de cambiar a causaMotivoAtencion 40
# las atenciones de finalidad 11 y causa 38

def corregir_causa_motivo_40(df):

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
    
        df.at[
            idx,
            'causaMotivoAtencion'
        ] = '40'
        

    print('\n' + '-' * 70)
    print('FUNCIÓN: corregir_causa_motivo_40')
    print('-' * 70)
    print(f'Registros encontrados: {cantidad}')

    return df


#-------------------------- Ejecutar corrección -------------------------

analisis_codigos = corregir_causa_motivo_40(
    analisis_codigos
)

#------------------------------------------------------------------------

def corregir_codtecnologia_sin_ceros(df):

    print('\n' + '-' * 70)
    print('FUNCIÓN: corregir_codtecnologia_sin_ceros')
    print('-' * 70)

    if 'codTecnologiaSalud' not in df.columns:
        print('No existe la columna codTecnologiaSalud')
        return df

    total_modificados = 0

    def limpiar_codigo(valor):

        if pd.isna(valor):
            return valor

        valor = str(valor)

        # separar parte antes y después del guion si existe
        if '-' in valor:
            parte1, parte2 = valor.split('-', 1)

            # eliminar ceros a la izquierda en ambas partes
            parte1 = str(int(parte1)) if parte1.isdigit() else parte1
            parte2 = str(int(parte2)) if parte2.isdigit() else parte2

            return f"{parte1}-{parte2}"

        else:
            # solo número con ceros iniciales
            if valor.isdigit():
                return str(int(valor))

        return valor

    mascara = (
        df['codTecnologiaSalud']
        .notna()
        & (df['codTecnologiaSalud'] != '')
    )

    antes = df.loc[mascara, 'codTecnologiaSalud'].copy()

    df.loc[mascara, 'codTecnologiaSalud'] = (
        df.loc[mascara, 'codTecnologiaSalud']
        .apply(limpiar_codigo)
    )
    
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

#-------------------------- Ejecutar corrección -------------------------

analisis_codigos = corregir_codtecnologia_sin_ceros(analisis_codigos)

#------------------------------------------------------------------------

# Esta función reemplaza IUM por CUM utilizando un Excel
# con columnas IUM y CUM

def corregir_ium_por_cum(df):

    columnas_requeridas = [
        'codTecnologiaSalud'
    ]

    if not all(col in df.columns for col in columnas_requeridas):
        print('No existen todas las columnas requeridas.')
        return df

    ruta_equivalencias = (
        r'C:\proyecto\equivalencias_ium_cum.xlsx'
    )

    if not os.path.exists(ruta_equivalencias):
        print(
            f'No se encontró el archivo: '
            f'{ruta_equivalencias}'
        )
        return df

    equivalencias = pd.read_excel(
        ruta_equivalencias,
        dtype=str
    )

    columnas_excel = ['IUM', 'CUM']

    if not all(
        col in equivalencias.columns
        for col in columnas_excel
    ):
        print(
            'El Excel debe contener '
            'las columnas IUM y CUM.'
        )
        return df

    equivalencias_ium_cum = dict(
        zip(
            equivalencias['IUM'],
            equivalencias['CUM']
        )
    )

    mascara = df[
        'codTecnologiaSalud'
    ].isin(
        equivalencias_ium_cum.keys()
    )

    cantidad = mascara.sum()

    print('\n' + '-' * 70)
    print('FUNCIÓN: corregir_ium_por_cum')
    print('-' * 70)
    print(
        f'Registros encontrados: '
        f'{cantidad}'
)

    df.loc[
        mascara,
        'codTecnologiaSalud'
    ] = df.loc[
        mascara,
        'codTecnologiaSalud'
    ].map(
        equivalencias_ium_cum
    )

    print(
        f'Registros modificados: '
        f'{cantidad}'
    )

    return df


#-------------------------- Ejecutar corrección -------------------------

analisis_codigos = corregir_ium_por_cum(
    analisis_codigos
)

#------------------------------------------------------------------------

print('\nProceso finalizado correctamente.')

#------------------------------------------------------------------------

import calendar

def reubicar_fechas_mes_objetivo(
    df,
    anio_objetivo,
    mes_objetivo
):

    print('\n' + '-' * 70)
    print('FUNCIÓN: reubicar_fechas_mes_objetivo')
    print('-' * 70)

    columnas_requeridas = [
        'fechaInicioAtencion',
        'fechaEgreso'
    ]

    if not all(col in df.columns for col in columnas_requeridas):
        print('No existen las columnas requeridas.')
        return df

    total_inicio_egreso = 0
    total_suministro = 0
    total_dispensacion = 0

    ultimo_dia_mes = calendar.monthrange(
        anio_objetivo,
        mes_objetivo
    )[1]

    inicio_original = pd.to_datetime(
        df['fechaInicioAtencion'],
        errors='coerce'
    )

    egreso_original = pd.to_datetime(
        df['fechaEgreso'],
        errors='coerce'
    )

    mascara_hospitalizacion = (
        inicio_original.notna() &
        egreso_original.notna()
    )

    for idx in df.index[mascara_hospitalizacion]:

        fecha_inicio = inicio_original.loc[idx]
        fecha_egreso = egreso_original.loc[idx]

        # ----------------------------------------------------
        # VALIDAR MÍNIMO 8 HORAS
        # ----------------------------------------------------

        diferencia_horas = (
            fecha_egreso - fecha_inicio
        ).total_seconds() / 3600

        ya_en_mes_objetivo = (
            fecha_inicio.year == anio_objetivo and
            fecha_inicio.month == mes_objetivo and
            fecha_egreso.year == anio_objetivo and
            fecha_egreso.month == mes_objetivo
        )

        if ya_en_mes_objetivo and diferencia_horas < 8:

            nueva_fecha_egreso = (
                fecha_inicio +
                pd.Timedelta(hours=8)
            )

            df.at[
                idx,
                'fechaEgreso'
            ] = nueva_fecha_egreso.strftime(
                '%Y-%m-%d %H:%M'
            )

            total_inicio_egreso += 1

            continue

        # ----------------------------------------------------
        # SI YA ESTÁ EN EL MES OBJETIVO
        # ----------------------------------------------------

        if ya_en_mes_objetivo:
            continue

        # ----------------------------------------------------
        # REUBICAR AL MES OBJETIVO
        # ----------------------------------------------------

        duracion_dias = (
            fecha_egreso.date() -
            fecha_inicio.date()
        ).days

        nueva_fecha_inicio = pd.Timestamp(
            year=anio_objetivo,
            month=mes_objetivo,
            day=1,
            hour=fecha_inicio.hour,
            minute=fecha_inicio.minute,
            second=fecha_inicio.second
        )

        # ----------------------------------------
        # MISMO DÍA O MENOS DE 8 HORAS
        # ----------------------------------------

        if diferencia_horas < 8:

            nueva_fecha_egreso = (
                nueva_fecha_inicio +
                pd.Timedelta(hours=8)
            )

        else:

            dia_egreso = min(
                1 + max(duracion_dias, 0),
                ultimo_dia_mes
            )

            nueva_fecha_egreso = pd.Timestamp(
                year=anio_objetivo,
                month=mes_objetivo,
                day=dia_egreso,
                hour=fecha_egreso.hour,
                minute=fecha_egreso.minute,
                second=fecha_egreso.second
            )

        df.at[
            idx,
            'fechaInicioAtencion'
        ] = nueva_fecha_inicio.strftime(
            '%Y-%m-%d %H:%M'
        )

        df.at[
            idx,
            'fechaEgreso'
        ] = nueva_fecha_egreso.strftime(
            '%Y-%m-%d %H:%M'
        )

        total_inicio_egreso += 1

    print(
        f'Inicio/Egreso corregidos: '
        f'{total_inicio_egreso}'
    )

    # ----------------------------------------------------
    # FECHA SUMINISTRO TECNOLOGÍA
    # ----------------------------------------------------

    if 'fechaSuministroTecnologia' in df.columns:

        fechas = pd.to_datetime(
            df['fechaSuministroTecnologia'],
            errors='coerce'
        )

        mascara = fechas.notna()

        for idx in df.index[mascara]:

            fecha = fechas.loc[idx]

            if (
                fecha.year == anio_objetivo and
                fecha.month == mes_objetivo
            ):
                continue

            nueva_fecha = pd.Timestamp(
                year=anio_objetivo,
                month=mes_objetivo,
                day=1,
                hour=fecha.hour,
                minute=fecha.minute,
                second=fecha.second
            )

            df.at[
                idx,
                'fechaSuministroTecnologia'
            ] = nueva_fecha.strftime(
                '%Y-%m-%d %H:%M'
            )

            total_suministro += 1

    print(
        f'fechaSuministroTecnologia corregidas: '
        f'{total_suministro}'
    )

    # ----------------------------------------------------
    # FECHA DISPENSACIÓN / ADMINISTRACIÓN
    # ----------------------------------------------------

    if 'fechaDispensAdmon' in df.columns:

        fechas = pd.to_datetime(
            df['fechaDispensAdmon'],
            errors='coerce'
        )

        mascara = fechas.notna()

        for idx in df.index[mascara]:

            fecha = fechas.loc[idx]

            if (
                fecha.year == anio_objetivo and
                fecha.month == mes_objetivo
            ):
                continue

            nueva_fecha = pd.Timestamp(
                year=anio_objetivo,
                month=mes_objetivo,
                day=1,
                hour=fecha.hour,
                minute=fecha.minute,
                second=fecha.second
            )

            df.at[
                idx,
                'fechaDispensAdmon'
            ] = nueva_fecha.strftime(
                '%Y-%m-%d %H:%M'
            )

            total_dispensacion += 1

    print(
        f'fechaDispensAdmon corregidas: '
        f'{total_dispensacion}'
    )

    print('-' * 70)

    return df

#-------------------------- Ejecutar corrección -------------------------

analisis_codigos = reubicar_fechas_mes_objetivo(
    analisis_codigos,
    2026,  # año objetivo
    4      # mes objetivo
)

#------------------------------------------------------------------------

# Esta función corrige finalidades para procedimientos
# con finalidad 11 según el diagnóstico principal (CIE10)

def corregir_finalidades_por_diagnostico(df):

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

    # ----------------------------------------------------
    # Z300 - Z309 -> Finalidad 19
    # ----------------------------------------------------

    diagnosticos_z30 = [
        'Z300', 'Z301', 'Z302', 'Z303', 'Z304',
        'Z305', 'Z306', 'Z307', 'Z308', 'Z309'
    ]

    mascara_z30 = (
        (df['tipoRegistro'] == 'procedimientos') &
        (df['finalidadTecnologiaSalud'] == '11') &
        (
            df['codDiagnosticoPrincipal'].isin(
                diagnosticos_z30
            )
        )
    )

    cantidad_z30 = mascara_z30.sum()

    df.loc[
        mascara_z30,
        'finalidadTecnologiaSalud'
    ] = '19'

    total_modificados += cantidad_z30

    print(
        f'Z300-Z309 -> Finalidad 19: '
        f'{cantidad_z30}'
    )

    # ----------------------------------------------------
    # Z310 - Z319 -> Finalidad 22
    # ----------------------------------------------------

    diagnosticos_z31 = [
        'Z310', 'Z311', 'Z312', 'Z313', 'Z314',
        'Z315', 'Z316', 'Z317', 'Z318', 'Z319'
    ]

    mascara_z31 = (
        (df['tipoRegistro'] == 'procedimientos') &
        (df['finalidadTecnologiaSalud'] == '11') &
        (
            df['codDiagnosticoPrincipal'].isin(
                diagnosticos_z31
            )
        )
    )

    cantidad_z31 = mascara_z31.sum()

    df.loc[
        mascara_z31,
        'finalidadTecnologiaSalud'
    ] = '22'

    total_modificados += cantidad_z31

    print(
        f'Z310-Z319 -> Finalidad 22: '
        f'{cantidad_z31}'
    )

    # ----------------------------------------------------
    # Z320 - Z369 -> Finalidad 23
    # ----------------------------------------------------

    diagnosticos_z32_z36 = [
        f'Z{i}'
        for i in range(320, 370)
    ]

    mascara_z32_z36 = (
        (df['tipoRegistro'] == 'procedimientos') &
        (df['finalidadTecnologiaSalud'] == '11') &
        (
            df['codDiagnosticoPrincipal'].isin(
                diagnosticos_z32_z36
            )
        )
    )

    cantidad_z32_z36 = mascara_z32_z36.sum()

    df.loc[
        mascara_z32_z36,
        'finalidadTecnologiaSalud'
    ] = '23'

    total_modificados += cantidad_z32_z36

    print(
        f'Z320-Z369 -> Finalidad 23: '
        f'{cantidad_z32_z36}'
   )
    
    # ----------------------------------------------------
    # Procedimiento inicia en 990...
    # -> Finalidad 40
    # ----------------------------------------------------
    
    mascara_proc990 = (
        (df['tipoRegistro'] == 'procedimientos') &
        (df['finalidadTecnologiaSalud'] == '11') &
        (
            df['codProcedimiento']
            .fillna('')
            .str.startswith('990')
        )
    )
    
    cantidad_proc990 = mascara_proc990.sum()
    
    df.loc[
        mascara_proc990,
        'finalidadTecnologiaSalud'
    ] = '40'
    
    total_modificados += cantidad_proc990
    
    print(
        'Procedimiento 990* '
        f'-> Finalidad 40: {cantidad_proc990}'
    )
    
    # ----------------------------------------------------
    # Procedimiento inicia en 992... o 997...
    # -> Finalidad 14
    # ----------------------------------------------------
    
    mascara_proc992_997 = (
        (df['tipoRegistro'] == 'procedimientos') &
        (df['finalidadTecnologiaSalud'] == '11') &
        (
            df['codProcedimiento']
            .fillna('')
            .str.startswith(('992', '997'))
        )
    )
    
    cantidad_proc992_997 = mascara_proc992_997.sum()
    
    df.loc[
        mascara_proc992_997,
        'finalidadTecnologiaSalud'
    ] = '14'
    
    total_modificados += cantidad_proc992_997
    
    print(
        'Procedimiento 992* o 997* '
        f'-> Finalidad 14: {cantidad_proc992_997}'
    )

    # ----------------------------------------------------
    # Procedimiento inicia en 90.... y diagnóstico inicia en Z...
    # -> Finalidad 15
    # ----------------------------------------------------

    mascara_proc90_z = (
        (df['tipoRegistro'] == 'procedimientos') &
        (df['finalidadTecnologiaSalud'] == '11') &
        (
            df['codProcedimiento']
            .fillna('')
            .str.startswith('90')
        ) &
        (
            df['codDiagnosticoPrincipal']
            .fillna('')
            .str.startswith('Z')
        )
    )

    cantidad_proc90_z = mascara_proc90_z.sum()

    df.loc[
        mascara_proc90_z,
        'finalidadTecnologiaSalud'
    ] = '15'

    total_modificados += cantidad_proc90_z

    print(
        'Procedimiento 90* + Diagnóstico Z* '
        f'-> Finalidad 15: {cantidad_proc90_z}'
    )
    
    # ----------------------------------------------------
    # Procedimiento inicia en 90.... y diagnóstico NO inicia en Z
    # -> Finalidad 12
    # ----------------------------------------------------
    
    mascara_proc90_no_z = (
        (df['tipoRegistro'] == 'procedimientos') &
        (df['finalidadTecnologiaSalud'] == '11') &
        (
            df['codProcedimiento']
            .fillna('')
            .str.startswith('90')
        ) &
        (
            ~df['codDiagnosticoPrincipal']
            .fillna('')
            .str.startswith('Z')
        ) &
        (
            df['codDiagnosticoPrincipal']
            .fillna('')
            .ne('')
        )
    )
    
    cantidad_proc90_no_z = mascara_proc90_no_z.sum()
    
    df.loc[
        mascara_proc90_no_z,
        'finalidadTecnologiaSalud'
    ] = '12'
    
    total_modificados += cantidad_proc90_no_z
    
    print(
        'Procedimiento 90* + Diagnóstico NO Z* '
        f'-> Finalidad 12: {cantidad_proc90_no_z}'
    )
    
    # ----------------------------------------------------
    # CUPS ESPECÍFICOS -> FINALIDADES FIJAS
    # ----------------------------------------------------
    
    reglas_cups = {
        "892901": "12",
        "896101": "12",
        "898001": "12",
        "870454": "15",
        "870455": "15",
        "911009": "15",
        "911015": "15",
        "911018": "15",
        "869501": "17",
        "a20002": "16"
    }
    
    total_cups = 0
    
    codigos_proc = (
        df['codProcedimiento']
        .fillna('')
        .astype(str)
    )
    
    for cups, finalidad in reglas_cups.items():
    
        mascara_cups = (
            (df['tipoRegistro'] == 'procedimientos') &
            (codigos_proc.str.lower() == cups.lower()) &
            (df['finalidadTecnologiaSalud'] == '11')
        )
    
        cantidad = mascara_cups.sum()
    
        df.loc[
            mascara_cups,
            'finalidadTecnologiaSalud'
        ] = finalidad
    
        total_cups += cantidad
    
        print(f'CUPS {cups} -> Finalidad {finalidad}: {cantidad}')
    
    total_modificados += total_cups

    print('-' * 70)
    print(
        f'Total registros modificados: '
        f'{total_modificados}'
    )
    print('-' * 70)
    
    
    
    return df

#-------------------------- Ejecutar corrección -------------------------

analisis_codigos = corregir_finalidades_por_diagnostico(
    analisis_codigos
)

#------------------------------------------------------------------------

#------------------------------------------------------------------------

def corregir_sexo_finalidad_23(df):

    print('\n' + '-' * 70)
    print('FUNCIÓN: corregir_sexo_finalidad_23')
    print('-' * 70)

    columnas_requeridas = [
        'finalidadTecnologiaSalud',
        'codSexo'
    ]

    if not all(col in df.columns for col in columnas_requeridas):
        print('No existen todas las columnas requeridas.')
        return df

    mascara = (
        (df['finalidadTecnologiaSalud'].fillna('') == '23')
        &
        (df['codSexo'].fillna('') != 'F')
    )

    cantidad = mascara.sum()

    df.loc[
        mascara,
        'codSexo'
    ] = 'F'

    print(
        f'Registros con finalidad 23 corregidos a sexo F: '
        f'{cantidad}'
    )

    print('-' * 70)

    return df


#-------------------------- Ejecutar corrección -------------------------

analisis_codigos = corregir_sexo_finalidad_23(
    analisis_codigos
)

#------------------------------------------------------------------------

#------------------------------------------------------------------------

def corregir_pais_origen(df):

    print('\n' + '-' * 70)
    print('FUNCIÓN: corregir_pais_origen')
    print('-' * 70)

    if 'codPaisOrigen' not in df.columns:
        print('No existe la columna codPaisOrigen')
        return df

    mascara = (
        df['codPaisOrigen'].isna()
        |
        (
            df['codPaisOrigen']
            .astype(str)
            .str.strip()
            .eq('')
        )
    )

    cantidad = mascara.sum()

    df.loc[
        mascara,
        'codPaisOrigen'
    ] = '170'

    print(
        f'Registros corregidos: {cantidad}'
    )

    print('-' * 70)

    return df

#------------------------------------------------------------------------

#-------------------------- Ejecutar corrección -------------------------

analisis_codigos = corregir_pais_origen(
    analisis_codigos
)

#------------------------------------------------------------------------

#------------------------------------------------------------------------

def corregir_condicion_destino_egreso(df):

    print('\n' + '-' * 70)
    print('FUNCIÓN: corregir_condicion_destino_egreso')
    print('-' * 70)

    if 'condicionDestinoUsuarioEgreso' not in df.columns:
        print(
            'No existe la columna '
            'condicionDestinoUsuarioEgreso'
        )
        return df

    mascara = (
        df['condicionDestinoUsuarioEgreso']
        .notna()
        &
        (
            df['condicionDestinoUsuarioEgreso']
            .astype(str)
            .str.strip()
            != ''
        )
        &
        (
            df['condicionDestinoUsuarioEgreso']
            .astype(str)
            .str.strip()
            != '01'
        )
    )

    cantidad = mascara.sum()

    df.loc[
        mascara,
        'condicionDestinoUsuarioEgreso'
    ] = '01'

    print(
        f'Registros corregidos: {cantidad}'
    )

    print('-' * 70)

    return df

#------------------------------------------------------------------------

#-------------------------- Ejecutar corrección -------------------------

analisis_codigos = corregir_condicion_destino_egreso(
    analisis_codigos
)

#------------------------------------------------------------------------

#------------------------------------------------------------------------

def corregir_diagnosticos_duplicados(df):

    print('\n' + '-' * 70)
    print('FUNCIÓN: corregir_diagnosticos_duplicados')
    print('-' * 70)

    if 'codDiagnosticoPrincipal' not in df.columns:
        print(
            'No existe la columna '
            'codDiagnosticoPrincipal'
        )
        return df

    columnas_relacionadas = [
        col
        for col in df.columns
        if 'codDiagnosticoRelacionado' in col
    ]

    if not columnas_relacionadas:
        print(
            'No se encontraron columnas '
            'de diagnósticos relacionados.'
        )
        return df

    total_modificados = 0

    for idx in df.index:

        principal = str(
            df.at[idx, 'codDiagnosticoPrincipal']
        ).strip()

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

                df.at[idx, columna] = '.null.'
                total_modificados += 1

            else:

                diagnosticos_vistos.add(valor)

    print(
        f'Diagnósticos relacionados corregidos: '
        f'{total_modificados}'
    )

    print('-' * 70)

    return df

#------------------------------------------------------------------------
#-------------------------- Ejecutar corrección -------------------------

analisis_codigos = corregir_diagnosticos_duplicados(
    analisis_codigos
)

#------------------------------------------------------------------------
#-------------------------- Exportar archivo ----------------------------


nombre_salida = (
    os.path.splitext(
        os.path.basename(ruta)
    )[0]
    + '_CORREGIDO.CSV'
)

ruta_salida = os.path.join(
    r'C:\proyecto\factura_corregida',
    nombre_salida
)

analisis_codigos.to_csv(
    ruta_salida,
    sep=';',
    index=False,
    encoding='utf-8'
)

print('\n' + '=' * 70)
print('PROCESO FINALIZADO')
print('=' * 70)

print(
    f'Archivo exportado correctamente:\n'
    f'{ruta_salida}'
)

print('=' * 70)