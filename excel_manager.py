"""
=========================================================
excel_manager.py
---------------------------------------------------------
Manejo del archivo Maestro.xlsx
=========================================================
"""

from pathlib import Path
from datetime import datetime
import shutil

import pandas as pd
from openpyxl import load_workbook

import config


class ExcelManager:

    def __init__(self):

        self.archivo = Path(config.MAESTRO)

        self.df = None

        self.columnas = []

        self.columnas_fijas = []

        self.columnas_meses = []

        self.cargar()


    # =====================================================
    # CARGAR EXCEL
    # =====================================================

    def cargar(self):

        if not self.archivo.exists():

            raise FileNotFoundError(
                f"No existe el archivo {self.archivo}"
            )

        self.df = pd.read_excel(self.archivo)

        self.columnas = list(self.df.columns)

        self.detectar_columnas()


    # =====================================================
    # DETECTAR COLUMNAS
    # =====================================================

    def detectar_columnas(self):

        columnas_fijas = []

        columnas_meses = []

        for columna in self.columnas:

            nombre = str(columna).strip()

            if self.es_columna_mes(nombre):

                columnas_meses.append(nombre)

            else:

                columnas_fijas.append(nombre)

        self.columnas_fijas = columnas_fijas

        self.columnas_meses = columnas_meses


    # =====================================================
    # IDENTIFICAR SI ES MES
    # =====================================================

    @staticmethod
    def es_columna_mes(nombre):

        meses = [

            "ene",
            "feb",
            "mar",
            "abr",
            "may",
            "jun",
            "jul",
            "ago",
            "sep",
            "oct",
            "nov",
            "dic"

        ]

        texto = nombre.lower()

        if len(texto) < 5:
            return False

        for mes in meses:

            if texto.startswith(mes):
                return True

        return False


    # =====================================================
    # REGIONES
    # =====================================================

    def obtener_regiones(self):

        regiones = sorted(

            self.df[config.COL_REGION]
            .dropna()
            .unique()
            .tolist()

        )

        return regiones


    # =====================================================
    # EMPRESAS POR REGION
    # =====================================================

    def obtener_empresas(self, region):

        datos = self.df[
            self.df[config.COL_REGION] == region
        ]

        empresas = sorted(

            datos[config.COL_EMPRESA]
            .dropna()
            .tolist()

        )

        return empresas


    # =====================================================
    # Ns POR REGION
    # =====================================================

    def obtener_n(self, region):

        datos = self.df[
            self.df[config.COL_REGION] == region
        ]

        lista = sorted(

            datos[config.COL_N]
            .dropna()
            .tolist()

        )

        return lista


    # =====================================================
    # BUSCAR POR EMPRESA
    # =====================================================

    def buscar_empresa(self, empresa):

        datos = self.df[

            self.df[config.COL_EMPRESA] == empresa

        ]

        if datos.empty:

            return None

        return datos.iloc[0]


    # =====================================================
    # BUSCAR POR N
    # =====================================================

    def buscar_n(self, n):

        datos = self.df[

            self.df[config.COL_N] == n

        ]

        if datos.empty:

            return None

        return datos.iloc[0]


    # =====================================================
    # OBTENER INFORMACION
    # =====================================================

    def obtener_informacion_empresa(self, empresa=None, n=None):

        if empresa is not None:

            fila = self.buscar_empresa(empresa)

        else:

            fila = self.buscar_n(n)

        if fila is None:

            return None

        return {

            "empresa": fila[config.COL_EMPRESA],

            "n": fila[config.COL_N],

            "region": fila[config.COL_REGION],

            "seccion": fila[config.COL_SECCION],

            "division": fila[config.COL_DIVISION],

            "grupo": fila[config.COL_GRUPO]

        }


    # =====================================================
    # EXISTE COLUMNA
    # =====================================================

    def existe_columna(self, nombre):

        return nombre in self.df.columns


    # =====================================================
    # LISTA COLUMNAS MESES
    # =====================================================

    def obtener_meses(self):

        return self.columnas_meses
    # =====================================================
    # FORMATO NOMBRE MES
    # =====================================================

    @staticmethod
    def formatear_mes(mes, anio):

        meses = {
            1: "Ene",
            2: "Feb",
            3: "Mar",
            4: "Abr",
            5: "May",
            6: "Jun",
            7: "Jul",
            8: "Ago",
            9: "Sep",
            10: "Oct",
            11: "Nov",
            12: "Dic"
        }

        if isinstance(mes, str):

            meses_texto = {
                "enero": 1,
                "febrero": 2,
                "marzo": 3,
                "abril": 4,
                "mayo": 5,
                "junio": 6,
                "julio": 7,
                "agosto": 8,
                "septiembre": 9,
                "octubre": 10,
                "noviembre": 11,
                "diciembre": 12
            }

            mes = meses_texto[mes.lower()]

        anio = str(anio)[-2:]

        return f"{meses[mes]}-{anio}"


    # =====================================================
    # CREAR COLUMNA MES
    # =====================================================

    def crear_columna_mes(self, mes, anio):

        nombre = self.formatear_mes(mes, anio)

        if nombre not in self.df.columns:

            self.df[nombre] = pd.NA

            self.columnas.append(nombre)

            self.columnas_meses.append(nombre)

        return nombre


    # =====================================================
    # OBTENER FILA POR N
    # =====================================================

    def obtener_indice_por_n(self, n):

        indice = self.df.index[
            self.df[config.COL_N] == n
        ]

        if len(indice) == 0:

            return None

        return indice[0]


    # =====================================================
    # EXISTE DATO
    # =====================================================

    def existe_dato(self, n, columna):

        fila = self.obtener_indice_por_n(n)

        if fila is None:

            return False

        valor = self.df.loc[fila, columna]

        if pd.isna(valor):

            return False

        return True


    # =====================================================
    # LEER DATO
    # =====================================================

    def leer_valor(self, n, columna):

        fila = self.obtener_indice_por_n(n)

        if fila is None:

            return None

        return self.df.loc[fila, columna]


    # =====================================================
    # ACTUALIZAR VALOR
    # =====================================================

    def actualizar_valor(self, n, mes, anio, valor):

        columna = self.formatear_mes(mes, anio)

        if columna not in self.df.columns:

            self.crear_columna_mes(mes, anio)

        fila = self.obtener_indice_por_n(n)

        if fila is None:

            raise ValueError("No existe el N seleccionado.")

        self.df.loc[fila, columna] = valor

        return columna


    # =====================================================
    # GUARDAR EXCEL
    # =====================================================

    def guardar(self):

        self.df.to_excel(

            self.archivo,

            index=False

        )


    # =====================================================
    # GUARDAR Y RECARGAR
    # =====================================================

    def guardar_y_recargar(self):

        self.guardar()

        self.cargar()


    # =====================================================
    # ACTUALIZAR DESDE FORMULARIO
    # =====================================================

    def guardar_captura(
            self,
            n,
            mes,
            anio,
            valor
    ):

        columna = self.actualizar_valor(
            n=n,
            mes=mes,
            anio=anio,
            valor=valor
        )

        self.guardar()

        return columna

    # =====================================================
    # BUSCAR COLUMNA DE MES
    # =====================================================

    def buscar_columna_mes(self, mes, anio):

        objetivo = self.formatear_mes(mes, anio).lower()

        # Primero busca coincidencia exacta
        for columna in self.df.columns:

            texto = str(columna).strip().lower()

            if texto == objetivo:
                return columna

        # Si la columna es una fecha (Timestamp)
        fecha = datetime(anio, mes if isinstance(mes, int) else {
            "enero":1,
            "febrero":2,
            "marzo":3,
            "abril":4,
            "mayo":5,
            "junio":6,
            "julio":7,
            "agosto":8,
            "septiembre":9,
            "octubre":10,
            "noviembre":11,
            "diciembre":12
        }[mes.lower()], 1)

        for columna in self.df.columns:

            try:

                fecha_col = pd.to_datetime(columna)

                if (
                    fecha_col.month == fecha.month
                    and
                    fecha_col.year == fecha.year
                ):

                    return columna

            except:

                continue

        return None


    # =====================================================
    # BACKUP
    # =====================================================

    def crear_backup(self):

        fecha = datetime.now().strftime("%Y%m%d_%H%M%S")

        nombre = f"Maestro_{fecha}.xlsx"

        destino = config.BACKUP_DIR / nombre

        shutil.copy2(self.archivo, destino)

        return destino


    # =====================================================
    # GUARDAR CON BACKUP
    # =====================================================

    def guardar_seguro(self):

        self.crear_backup()

        self.guardar()


    # =====================================================
    # RECARGAR
    # =====================================================

    def recargar(self):

        self.cargar()


    # =====================================================
    # INFORMACION GENERAL
    # =====================================================

    def resumen(self):

        return {

            "empresas": len(self.df),

            "regiones": len(self.obtener_regiones()),

            "meses": len(self.columnas_meses),

            "columnas": len(self.df.columns)

        }


    # =====================================================
    # ULTIMO MES
    # =====================================================

    def ultimo_mes(self):

        if len(self.columnas_meses) == 0:

            return None

        return self.columnas_meses[-1]


    # =====================================================
    # VERIFICAR ARCHIVOS
    # =====================================================

    def verificar(self):

        return {

            "maestro": self.archivo.exists(),

            "backups": config.BACKUP_DIR.exists()

        }


    # =====================================================
    # CERRAR
    # =====================================================

    def cerrar(self):

        self.df = None
