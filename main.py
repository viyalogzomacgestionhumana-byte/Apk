from kivymd.app import MDApp
from kivy.lang import Builder
from kivymd.uix.list import TwoLineListItem
from kivymd.uix.dialog import MDDialog
from kivymd.uix.button import MDFlatButton
from kivy.clock import Clock
from datetime import datetime, timedelta
from dateutil.relativedelta import relativedelta
import threading
import unicodedata
import requests
import io
import openpyxl

URL_EXCEL = "https://docs.google.com/spreadsheets/d/1VeCzn0x0rdhb3i2NWPHfjZMWTVubJK3G/export?format=xlsx"

KV = '''
MDScreen:
    MDBoxLayout:
        orientation: 'vertical'
        padding: "10dp"
        spacing: "10dp"

        MDTextField:
            id: search_field
            hint_text: "Descargando base de datos..."
            mode: "rectangle"
            disabled: True
            on_text: app.buscar_empleado(self.text)

        MDScrollView:
            MDList:
                id: result_list

    MDSpinner:
        id: spinner
        size_hint: None, None
        size: "46dp", "46dp"
        pos_hint: {'center_x': .5, 'center_y': .5}
        active: True
'''

class BuscadorApp(MDApp):
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.df = []
        self.dialog = None

    def build(self):
        self.theme_cls.primary_palette = "Teal"
        self.theme_cls.theme_style = "Light"
        return Builder.load_string(KV)

    def on_start(self):
        threading.Thread(target=self.cargar_datos, daemon=True).start()

    def normalizar(self, texto):
        texto = str(texto).upper()
        texto = ''.join(c for c in unicodedata.normalize('NFD', texto) if unicodedata.category(c) != 'Mn')
        return texto.replace(' ', '')

    def a_nombre_propio(self, texto):
        if not texto: return ''
        return ' '.join(p.capitalize() for p in str(texto).lower().split())

    def formatear_valor(self, campo, valor):
        campos_fecha = ['FECHA INGRESO', 'FECHA RETIRO', 'FECHA VACIO', 'FECHA NACIMIENTO', 'FECHA VENCIMIENTO', 'LIMITE NOTIFICACION']
        if campo in campos_fecha:
            if isinstance(valor, datetime): return valor.strftime('%d/%m/%Y')
            return str(valor) if valor else ''
        elif campo == 'TIEMPO':
            try: return f"{float(valor):.2f}"
            except: return str(valor) if valor else ''
        elif campo == 'CUENTA BANCARIA':
            v = str(valor)
            if v.endswith('.0'): v = v[:-2]
            return v if v and v != 'None' else ''
        elif campo in ['CELULAR', 'GRADO ESCOLAR']: 
            try: return str(int(float(valor)))
            except: return str(valor) if valor else ''
        return str(valor) if valor is not None else ''

    def calcular_fechas_row(self, row):
        try:
            f_ing = row.get('FECHA INGRESO')
            if not isinstance(f_ing, datetime): return "", None, None
                
            hoy = datetime.today().replace(hour=0, minute=0, second=0, microsecond=0)
            dias = (hoy - f_ing).days
            meses = (hoy.year - f_ing.year) * 12 + (hoy.month - f_ing.month)
            if hoy.day < f_ing.day: meses -= 1
                
            hito, f_ret = "", None
            
            if dias <= 23:
                hito = "Periodo de prueba"
                f_ret = f_ing + timedelta(days=18)
            elif meses < 3:
                hito = "Alerta 3 meses"
                f_ret = f_ing + relativedelta(months=3) - timedelta(days=1)
            elif meses < 6:
                hito = "Alerta 6 meses"
                f_ret = f_ing + relativedelta(months=6) - timedelta(days=1)
            elif meses < 9:
                hito = "Alerta 9 meses"
                f_ret = f_ing + relativedelta(months=9) - timedelta(days=1)
            elif meses < 12:
                hito = "Alerta 12 meses"
                f_ret = f_ing + relativedelta(months=12) - timedelta(days=1)
            elif meses < 24:
                hito = "Renovación anual 2do año"
                f_ret = f_ing + relativedelta(years=2) - timedelta(days=1)
            elif meses < 36:
                hito = "Renovación anual 3er año"
                f_ret = f_ing + relativedelta(years=3) - timedelta(days=1)
            elif meses < 48:
                hito = "Renovación anual 4to año"
                f_ret = f_ing + relativedelta(years=4) - timedelta(days=1)
            else:
                hito = "Pasa a INDEFINIDO"
                
            f_pre = f_ret - timedelta(days=30) if f_ret else None
            return hito, f_ret, f_pre
        except:
            return "Error", None, None

    def generar_diagnostico_legal(self, row):
        try:
            estado = str(row.get('ESTADO', '')).strip().upper()
            tipo_contrato = str(row.get('TIPO DE CONTRATO', '')).strip().upper()
            
            if estado == 'RETIRADO':
                f_retiro = self.formatear_valor('FECHA RETIRO', row.get('FECHA RETIRO'))
                motivo = row.get('MOTIVO', 'No especificado')
                return f"🛑 ESTADO: RETIRADO\n📅 Retiro: {f_retiro}\n📋 Motivo: {motivo}"
                
            elif estado == 'ACTIVO' and 'TERMINO FIJO' not in tipo_contrato:
                return f"ℹ️ NO APLICA: Preaviso no requerido.\n📌 Contrato: {tipo_contrato}."
                
            elif estado == 'ACTIVO' and 'TERMINO FIJO' in tipo_contrato:
                f_ingreso = self.formatear_valor('FECHA INGRESO', row.get('FECHA INGRESO'))
                f_vence_str = self.formatear_valor('FECHA VENCIMIENTO', row.get('FECHA VENCIMIENTO'))
                limite_str = self.formatear_valor('LIMITE NOTIFICACION', row.get('LIMITE NOTIFICACION'))
                hito = row.get('HITO ACTUAL', '')
                
                f_vence_dt = row.get('FECHA VENCIMIENTO')
                dias_restantes = (f_vence_dt - datetime.today().replace(hour=0, minute=0, second=0, microsecond=0)).days if isinstance(f_vence_dt, datetime) else 0
                alerta = "🔴 ACCIÓN URGENTE - Notificar hoy." if dias_restantes <= 30 else "🟢 Estado: En seguimiento."
                
                return (f"📅 Ingreso: {f_ingreso}\n📍 Etapa: {hito}\n📆 Límite preaviso: {limite_str}\n"
                        f"⚠️ Días restantes: {dias_restantes}\n🛑 Retiro: {f_vence_str}\n{alerta}")
            return "⚠️ Faltan datos operativos."
        except:
            return "⚠️ Error al generar diagnóstico."

    def cargar_datos(self):
        try:
            respuesta = requests.get(URL_EXCEL)
            respuesta.raise_for_status()
            datos_memoria = io.BytesIO(respuesta.content)
            
            wb = openpyxl.load_workbook(filename=datos_memoria, data_only=True)
            ws = wb['BD']
            data = []
            headers = []
            
            for i, row in enumerate(ws.iter_rows(values_only=True)):
                if i < 8: continue
                if i == 8:
                    headers = [str(cell).strip().upper() if cell else f"COL_{j}" for j, cell in enumerate(row)]
                    continue
                
                row_dict = dict(zip(headers, row))
                row_dict['CEDULA'] = str(row_dict.get('CEDULA', '')).split('.')[0] if row_dict.get('CEDULA') else ''
                row_dict['NOMBRE'] = str(row_dict.get('NOMBRE', ''))
                
                hito, f_vence, lim_notif = self.calcular_fechas_row(row_dict)
                row_dict['HITO ACTUAL'] = hito
                row_dict['FECHA VENCIMIENTO'] = f_vence
                row_dict['LIMITE NOTIFICACION'] = lim_notif
                row_dict['DIAGNOSTICO LEGAL'] = self.generar_diagnostico_legal(row_dict)
                
                data.append(row_dict)
            
            self.df = data
            Clock.schedule_once(self.activar_interfaz)
        except Exception as e:
            print(f"Error cargando base de datos: {e}")

    def activar_interfaz(self, dt):
        self.root.ids.spinner.active = False
        self.root.ids.search_field.disabled = False
        self.root.ids.search_field.hint_text = "Escriba cédula o nombre..."

    def buscar_empleado(self, texto):
        self.root.ids.result_list.clear_widgets()
        criterio = texto.strip().lower()
        if not criterio or not self.df: return

        resultados = []
        for row in self.df:
            ced = str(row.get('CEDULA', '')).lower()
            nom = str(row.get('NOMBRE', '')).lower()
            if criterio in ced or criterio in nom:
                resultados.append(row)
                if len(resultados) >= 15: break
        
        for row in resultados:
            item = TwoLineListItem(
                text=f"{row['NOMBRE']}",
                secondary_text=f"C.C: {row['CEDULA']} - Cargo: {row.get('CARGO', '')}",
                on_release=lambda x, fila=row: self.mostrar_detalle(fila)
            )
            self.root.ids.result_list.add_widget(item)

    def mostrar_detalle(self, fila):
        nombre = self.a_nombre_propio(fila.get('NOMBRE', ''))
        cedula = fila.get('CEDULA', '')
        estado = fila.get('ESTADO', '')
        diag = fila.get('DIAGNOSTICO LEGAL', '')
        
        contenido = (
            f"[b]Estado:[/b] {estado}\n"
            f"[b]Cédula:[/b] {cedula}\n"
            f"[b]Cargo:[/b] {fila.get('CARGO', '')}\n"
            f"[b]Labor:[/b] {fila.get('LABOR', '')}\n"
            f"[b]Supervisor:[/b] {fila.get('SUPERVISOR', '')}\n"
            f"[b]Celular:[/b] {self.formatear_valor('CELULAR', fila.get('CELULAR', ''))}\n"
            f"{'-'*30}\n"
            f"[b]⚖️ DIAGNÓSTICO LEGAL[/b]\n{diag}"
        )
        
        self.dialog = MDDialog(
            title=f"{nombre}",
            text=contenido,
            radius=[20, 7, 20, 7],
            buttons=[MDFlatButton(text="CERRAR", theme_text_color="Custom", text_color=self.theme_cls.primary_color, on_release=lambda x: self.dialog.dismiss())],
        )
        self.dialog.open()

if __name__ == '__main__':
    BuscadorApp().run()
