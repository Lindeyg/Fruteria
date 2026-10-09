import re
import sys
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from ttkbootstrap.dialogs import Messagebox
from reportlab.pdfgen import canvas
from datetime import datetime
from Rapport import generar_reporte_ventas, generar_reporte_stock, generar_reporte_cuadre_caja                      #AGREGADO 1: Hasta fila 15
from tkinter import filedialog                                                                                      #Importaciones para el diálogo de guardar archivo Excel    
from Rapport import (                                                                                                                                                                                            #AGREGADO 1    
    generar_reporte_ventas,                                                                                             
    generar_reporte_stock,                                                                                          
    generar_reporte_cuadre_caja                                                                                     
)                                                                                                                   #AGREGADO 1: HASTA AQUI    

# ==========================================
# 1. CONFIGURACIÓN Y CONSTANTES DEL SISTEMA
# ==========================================
TIPO_IVA = 0.04
BOLSA_INFO = ("🛍️", "Bolsa", 0.05, "ud")
FRUTAS = {
    "cereza":    ("🍒", "Cereza",    4.50, "kg"),
    "datil":     ("🌴", "Dátil",     6.20, "kg"),
    "fresa":     ("🍓", "Fresa",     3.80, "kg"),
    "kiwi":      ("🥝", "Kiwi",      3.20, "kg"),
    "limon":     ("🍋", "Limón",     1.60, "kg"),
    "mango":     ("🥭", "Mango",     3.50, "kg"),
    "manzana":   ("🍎", "Manzana",   1.95, "kg"),
    "melocoton": ("🍑", "Melocotón", 2.40, "kg"),
    "melon":     ("🍈", "Melón",     1.20, "ud"),
    "naranja":   ("🍊", "Naranja",   1.50, "kg"),
    "pera":      ("🍐", "Pera",      2.15, "kg"),
    "pina":      ("🍍", "Piña",      1.80, "ud"),
    "platano":   ("🍌", "Plátano",   2.10, "kg"),
    "sandia":    ("🍉", "Sandía",    0.95, "ud"),
    "uva":       ("🍇", "Uva",       2.90, "kg")
}
# ==========================================
# 2. FUNCIONES DE LÓGICA Y AUXILIARES
# ==========================================
def formato_precio(numero):
    return f"{numero:.2f} €".replace(".", ",")
def parsear_entrada_rafaga(entrada):
    entrada = entrada.strip()
    if not entrada:
        return "", None
    if entrada.startswith("&"):
        resto = entrada[1:].strip()
        return "bolsa", resto if resto else "1"
    m = re.search(r'(?=[+\-*]|\s+\d)', entrada)
    if m:
        idx = m.start()
        fruta = entrada[:idx].strip()
        cantidad = entrada[idx:].strip()
        return fruta, cantidad
    return entrada, None
def buscar_fruta(fruta):
    clave_entrada = fruta.lower().replace("á", "a").replace("é", "e").replace("í", "i").replace("ó", "o").replace("ú", "u")
    if not clave_entrada:
        return []
    return [k for k in FRUTAS if k.startswith(clave_entrada)]
def calcular_variacion_rafaga(kg_acumulados, cantidad_str, unidad="kg"):
    if not cantidad_str:
        return 1.0
    cantidad_str = cantidad_str.strip().replace(",", ".")
    if cantidad_str.startswith("+"):
        cantidad_str = cantidad_str[1:]
    if cantidad_str.startswith("."):
        cantidad_str = "0" + cantidad_str
    elif cantidad_str.startswith("-."):
        cantidad_str = cantidad_str.replace("-.", "-0.")
    elif cantidad_str.startswith("*."):
        cantidad_str = cantidad_str.replace("*.", "*0.")
    es_fijado = False
    if cantidad_str.startswith("*"):
        es_fijado = True
        cantidad_str = cantidad_str.replace("*", "")
    try:
        num = int(float(cantidad_str)) if unidad == "ud" else float(cantidad_str)
        if es_fijado:
            return num - kg_acumulados
        if num < 0 and (kg_acumulados + num) < 0:
            return None  # Error: restar más de lo acumulado
        return num
    except ValueError:
        return None
def crear_pdf(cesta, id_ticket, entrega, cambio, metodo_pago):
    total = sum(item["total"] for item in cesta.values())
    base = total / (1 + TIPO_IVA)
    iva = total - base
    ANCHO_L = 36
    fecha_actual = datetime.now().strftime("%d/%m/%Y %H:%M")
    lineas = [
        "FRUITERIA L'HORT".center(ANCHO_L),
        "CIF: B-12345678".center(ANCHO_L),
        "Carrer Major 123, Barcelona".center(ANCHO_L),
        f"Fecha: {fecha_actual}".center(ANCHO_L),
        f"Nº Ticket: {id_ticket}".center(ANCHO_L),
        "=" * ANCHO_L,
        f"{'PRODUCTO':<10} {'CANT':>8} {'PVP':>8} {'TOTAL':>7}",
        "-" * ANCHO_L
    ]
    for item in cesta.values():
        nombre = item['nombre'][:11]
        cant_str = f"{int(item['kg'])} Ud" if item.get("unidad") == "ud" else f"{item['kg']:.2f} Kg".replace(".", ",")
        pvp = f"{item['pvp']:.2f} €".replace(".", ",")
        lineas.append(f"{nombre:<11} {cant_str:>7} {pvp:>8} {formato_precio(item['total']):>7}")
    bloque_pago = [
        "-" * ANCHO_L,
        f"{'Base Imponible:':<27} {formato_precio(base):>8}",
        f"{f'IVA ({TIPO_IVA * 100:g}%):':<27} {formato_precio(iva):>8}",
        "=" * ANCHO_L,
        f"{'TOTAL A PAGAR:':<27} {formato_precio(total):>8}",
        "=" * ANCHO_L,
        f"{'Forma de Pago:':<27} {metodo_pago:>8}"
    ]
    if metodo_pago == "Efectivo":
        bloque_pago.extend([
            f"{'Entregado:':<27} {formato_precio(entrega):>8}",
            f"{'Cambio:':<27} {formato_precio(cambio):>8}"
        ])
    bloque_pago.extend([
        "- " * (ANCHO_L // 2),
        "",
        "¡Gracias por su compra!".center(ANCHO_L)
    ])
    lineas.extend(bloque_pago)
    ancho_ticket = 226
    alto_ticket = 40 + len(lineas) * 12
    c = canvas.Canvas(f"ticket_{id_ticket}.pdf", pagesize=(ancho_ticket, alto_ticket))
    text_object = c.beginText(14, alto_ticket - 20)
    text_object.setFont("Courier-Bold", 8.5)
    text_object.setLeading(12)
    for linea in lineas:
        text_object.textLine(linea)
    c.drawText(text_object)
    c.save()
    
# ==========================================
# 3. DIÁLOGO DE PAGO EN GUI
# ==========================================
class VentanaPago(ttk.Toplevel):
    def __init__(self, parent, total):
        super().__init__(parent)
        self.title("Procesar Pago")
        self.geometry("400x350")
        self.resizable(False, False)
       
        self.total = total
        self.resultado = None  # (metodo, entrega, cambio)
        ttk.Label(self, text="Cobro de Pedido", font=("Helvetica", 16, "bold"), bootstyle="primary").pack(pady=10)
        ttk.Label(self, text=f"Total a Pagar: {formato_precio(total)}", font=("Helvetica", 14, "bold")).pack(pady=5)
        frame_efectivo = ttk.Labelframe(self, text=" Pago en Efectivo ", padding=15)
        frame_efectivo.pack(fill=X, padx=20, pady=10)
        ttk.Label(frame_efectivo, text="Importe Entregado (€):").pack(anchor=W)
        self.entry_entrega = ttk.Entry(frame_efectivo)
        self.entry_entrega.pack(fill=X, pady=5)
        self.entry_entrega.focus_set()
        self.entry_entrega.bind("<Return>", lambda e: self.pagar_efectivo())
        btn_efectivo = ttk.Button(frame_efectivo, text="Cobrar Efectivo", bootstyle="success", command=self.pagar_efectivo)
        btn_efectivo.pack(fill=X, pady=5)
        ttk.Separator(self).pack(fill=X, padx=20, pady=5)
        btn_tarjeta = ttk.Button(self, text="💳 Pago con Tarjeta", bootstyle="info", command=self.pagar_tarjeta)
        btn_tarjeta.pack(fill=X, padx=20, pady=5)
    def pagar_efectivo(self):
        txt = self.entry_entrega.get().strip().replace(",", ".")
        try:
            entrega = float(txt)
            if entrega < self.total:
                Messagebox.show_error(f"Importe insuficiente. Faltan {formato_precio(self.total - entrega)}", title="Error de Cobro", parent=self)
                return
            cambio = entrega - self.total
            self.resultado = ("Efectivo", entrega, cambio)
            self.destroy()
        except ValueError:
            Messagebox.show_error("Introduce un número válido para el pago en efectivo.", title="Error", parent=self)
    def pagar_tarjeta(self):
        self.resultado = ("Tarjeta", self.total, 0.0)
        self.destroy()
        
    def descargar_reporte_ttkbootstrap(parent, tipo_reporte):                                                 # AGREGADO 2: Función para descargar reportes Excel, hasta fila 216
        """Abre el diálogo modal nativo para guardar archivo Excel con respuesta TTKBootstrap."""
        fecha_str = datetime.now().strftime("%Y%m%d_%H%M%S")
        nombre_defecto = f"reporte_{tipo_reporte}_{fecha_str}.xlsx"
    
        ruta_guardar = filedialog.asksaveasfilename(
            parent=parent,
            title=f"Guardar Reporte de {tipo_reporte.capitalize()}",
            initialfile=nombre_defecto,
            defaultextension=".xlsx",
            filetypes=[("Libro de Excel", "*.xlsx"), ("Todos los archivos", "*.*")]
        )
    
        if not ruta_guardar:
            return  # El usuario canceló la ventana
        
        try:
            if tipo_reporte == "ventas":
                generar_reporte_ventas(ruta_guardar)
            elif tipo_reporte == "stock":
                generar_reporte_stock(ruta_guardar)
            elif tipo_reporte == "cuadre":
                generar_reporte_cuadre_caja(ruta_guardar)
            
            Messagebox.show_info(
                f"El reporte de {tipo_reporte.capitalize()} se ha guardado correctamente en:\n\n{ruta_guardar}",
                title="Descarga Exitosa",
                parent=parent
            )
        except Exception as e:
            Messagebox.show_error(
                f"No se pudo guardar el archivo Excel:\n{str(e)}",
                title="Error al Guardar",
                parent=parent                                                                                     
            )                                                                                                  # AGREGADO 2  hasta AQUI 
# ==========================================
# 4. APLICACIÓN PRINCIPAL (GUI)
# ==========================================
class AppTPV(ttk.Window):
    def __init__(self):
        super().__init__(themename="flatly")
        self.title("🛒 Frutería - Punto de Venta")
        self.geometry("1100x700")
        self.cesta = {}
        self.crear_layout()
        self.cargar_catalogo()
    def crear_layout(self):
        # Panel superior (Cabecera)
        header = ttk.Frame(self, padding=10, bootstyle="primary")
        header.pack(fill=X)
        ttk.Label(header, text="🛒 FRUTERÍA - PUNTO DE VENTA", font=("Helvetica", 18, "bold"), bootstyle="inverse-primary").pack(side=LEFT)
        # Contenedor Principal (2 Columnas)
        main_container = ttk.Frame(self, padding=10)
        main_container.pack(fill=BOTH, expand=True)
        # COLUMNA IZQUIERDA: Catálogo de Frutas
        left_frame = ttk.Labelframe(main_container, text=" Catálogo de Productos ", padding=10)
        left_frame.pack(side=LEFT, fill=BOTH, expand=True, padx=(0, 5))
        self.tabla_catalogo = ttk.Treeview(left_frame, columns=("icono", "nombre", "pvp"), show="headings", height=15)
        self.tabla_catalogo.heading("icono", text="")
        self.tabla_catalogo.heading("nombre", text="Fruta")
        self.tabla_catalogo.heading("pvp", text="PVP")

        self.tabla_catalogo.column("icono", width=40, anchor=CENTER)
        self.tabla_catalogo.column("nombre", width=150, anchor=W)
        self.tabla_catalogo.column("pvp", width=100, anchor=E)
        self.tabla_catalogo.pack(fill=BOTH, expand=True)
        self.tabla_catalogo.bind("<Double-1>", self.on_click_catalogo)
        # COLUMNA DERECHA: Entrada Ráfaga, Cesta y Acciones
        right_frame = ttk.Frame(main_container)
        right_frame.pack(side=RIGHT, fill=BOTH, expand=True, padx=(5, 0))
        # Entrada Ráfaga
        rafaga_frame = ttk.Labelframe(right_frame, text=" Entrada Ráfaga ⚡ ", padding=10)
        rafaga_frame.pack(fill=X, pady=(0, 10))
        self.entry_rafaga = ttk.Entry(rafaga_frame, font=("Helvetica", 12))
        self.entry_rafaga.pack(side=LEFT, fill=X, expand=True, padx=(0, 5))
        self.entry_rafaga.bind("<Return>", lambda e: self.procesar_rafaga())
        self.entry_rafaga.focus_set()
        btn_add = ttk.Button(rafaga_frame, text="Añadir", bootstyle="success", command=self.procesar_rafaga)
        btn_add.pack(side=RIGHT)
        ttk.Label(rafaga_frame, text="Ejemplos: [uva] (+1kg) | [nar+2] (+2kg) | [&] (1 Bolsa)", font=("Helvetica", 8), bootstyle="secondary").pack(anchor=W, pady=(5, 0))
        # Tabla Cesta de la Compra
        cesta_frame = ttk.Labelframe(right_frame, text=" Cesta Actual ", padding=10)
        cesta_frame.pack(fill=BOTH, expand=True)
        self.tabla_cesta = ttk.Treeview(cesta_frame, columns=("icono", "nombre", "cant", "total"), show="headings")
        self.tabla_cesta.heading("icono", text="")
        self.tabla_cesta.heading("nombre", text="Producto")
        self.tabla_cesta.heading("cant", text="Cantidad")
        self.tabla_cesta.heading("total", text="Total")
        self.tabla_cesta.column("icono", width=40, anchor=CENTER)
        self.tabla_cesta.column("nombre", width=140, anchor=W)
        self.tabla_cesta.column("cant", width=100, anchor=E)
        self.tabla_cesta.column("total", width=100, anchor=E)
        self.tabla_cesta.pack(fill=BOTH, expand=True, pady=(0, 10))
        # Total Label
        self.lbl_total = ttk.Label(cesta_frame, text="TOTAL: 0,00 €", font=("Helvetica", 16, "bold"), bootstyle="danger", anchor=E)
        self.lbl_total.pack(fill=X, pady=5)
        # Panel de Botones Globales                                                                         AGREGAR 3: HASTA FILA 322
        btn_frame = ttk.Frame(right_frame)
        btn_frame.pack(fill=X, pady=(10, 0))
        ttk.Button(btn_frame, text="🛍️ Bolsa (&)", bootstyle="info-outline", command=self.agregar_bolsa).pack(side=LEFT, fill=X, expand=True, padx=2)
        ttk.Button(btn_frame, text="❌ Cancelar (/)", bootstyle="warning-outline", command=self.cancelar_pedido).pack(side=LEFT, fill=X, expand=True, padx=2)
        ttk.Button(btn_frame, text="💳 Ticket / Cobrar (%)", bootstyle="success", command=self.procesar_ticket).pack(side=LEFT, fill=X, expand=True, padx=2)

        # --- AÑADIR A PARTIR DE AQUÍ (Panel de botones de Excel) ---
        reportes_frame = ttk.Labelframe(right_frame, text=" Módulo de Reportes Excel ", padding=10)
        reportes_frame.pack(fill=X, pady=(10, 0))

        ttk.Button(reportes_frame, text="📊 Ventas", bootstyle="secondary-outline",
                command=self.generar_reporte_ventas
                ).pack(side=LEFT, fill=X, expand=True, padx=2)

        ttk.Button(reportes_frame, text="📦 Stock", bootstyle="secondary-outline",
                command=self.generar_reporte_stock
                ).pack(side=LEFT, fill=X, expand=True, padx=2)

        ttk.Button(reportes_frame, text="💵 Cuadre Caja", bootstyle="secondary-outline",
                command=self.generar_reporte_cuadre_caja
                ).pack(side=LEFT, fill=X, expand=True, padx=2)
        # Contenedor del módulo de descargas con ttkbootstrap
        reportes_frame = ttk.Labelframe(right_frame, text=" 📥 Descargar Reportes Excel ", padding=10)       
        reportes_frame.pack(fill=X, pady=(10, 0))

        ttk.Button(
            reportes_frame,
            text="📊 Ventas",
            bootstyle="primary-outline",
            command=lambda: descargar_reporte_ttkbootstrap(self, "ventas")
        ).pack(side=LEFT, fill=X, expand=True, padx=2)

        ttk.Button(
            reportes_frame,
            text="📦 Stock",
            bootstyle="info-outline",
            command=lambda: descargar_reporte_ttkbootstrap(self, "stock")
        ).pack(side=LEFT, fill=X, expand=True, padx=2)

        ttk.Button(
            reportes_frame,
            text="💵 Cuadre Caja",
            bootstyle="success-outline",
            command=lambda: descargar_reporte_ttkbootstrap(self, "cuadre")
        ).pack(side=LEFT, fill=X, expand=True, padx=2)                                                          # AGREGAR 3: HASTA AQUI
    def _importar_reportes(self):
        from Rapport import generar_reporte_ventas, generar_reporte_stock, generar_reporte_cuadre_caja, registrar_venta_db
        return registrar_venta_db, generar_reporte_ventas, generar_reporte_stock, generar_reporte_cuadre_caja
    def generar_reporte_ventas(self):
        _, generar_reporte_ventas, _, _ = self._importar_reportes()
        Messagebox.show_info(f"Reporte generado: {generar_reporte_ventas()}", parent=self)
    def generar_reporte_stock(self):
        _, _, generar_reporte_stock, _ = self._importar_reportes()
        Messagebox.show_info(f"Reporte generado: {generar_reporte_stock()}", parent=self)
    def generar_reporte_cuadre_caja(self):
        _, _, _, generar_reporte_cuadre_caja = self._importar_reportes()
        Messagebox.show_info(f"Reporte generado: {generar_reporte_cuadre_caja()}", parent=self)
    def cargar_catalogo(self):
        for item in self.tabla_catalogo.get_children():
            self.tabla_catalogo.delete(item)
        for key, (icono, nombre, pvp, unidad) in FRUTAS.items():
            pvp_txt = f"{pvp:.2f} €/{unidad}"
            self.tabla_catalogo.insert("", END, iid=key, values=(icono, nombre, pvp_txt))
    def actualizar_cesta_gui(self):
        for item in self.tabla_cesta.get_children():
            self.tabla_cesta.delete(item)
        total_general = 0.0
        for key, item in self.cesta.items():
            cant_txt = f"{int(item['kg'])} Ud" if item.get("unidad") == "ud" else f"{item['kg']:.2f} Kg".replace(".", ",")
            self.tabla_cesta.insert("", END, values=(item['icono'], item['nombre'], cant_txt, formato_precio(item['total'])))
            total_general += item['total']
        self.lbl_total.config(text=f"TOTAL: {formato_precio(total_general)}")
    def procesar_rafaga(self):
        entrada = self.entry_rafaga.get().strip()
        self.entry_rafaga.delete(0, END)
        if not entrada:
            return
        if entrada == "/":
            self.cancelar_pedido()
            return
        elif entrada == "%":
            self.procesar_ticket()
            return
        fruta, cantidad_previo = parsear_entrada_rafaga(entrada)
        if fruta == "bolsa":
            clave = "bolsa"
            icono, nombre_bonito, precio, unidad = BOLSA_INFO
        else:
            coincidencias = buscar_fruta(fruta)
            if len(coincidencias) == 1:
                clave = coincidencias[0]
                icono, nombre_bonito, precio, unidad = FRUTAS[clave]
            elif len(coincidencias) > 1:
                sug = [FRUTAS[k][1] for k in coincidencias]
                Messagebox.show_warning(f"Coincidencias múltiples: {', '.join(sug)}", title="Atención", parent=self)
                return
            else:
                Messagebox.show_error(f"La fruta '{fruta}' no existe en el catálogo.", title="Error", parent=self)
                return
        existente = self.cesta.get(clave)
        kg_acum = existente["kg"] if existente else 0.0
        delta_kg = calcular_variacion_rafaga(kg_acum, cantidad_previo, unidad)
        if delta_kg is None:
            Messagebox.show_error("Sintaxis o cantidad no válida.", title="Error", parent=self)
            return
        self.actualizar_cesta_dict(clave, nombre_bonito, icono, precio, delta_kg, unidad)
        self.actualizar_cesta_gui()
    def actualizar_cesta_dict(self, clave, nombre, icono, precio, kg, unidad):
        if kg == 0:
            return
        if clave in self.cesta:
            self.cesta[clave]["kg"] += kg
            self.cesta[clave]["total"] = self.cesta[clave]["kg"] * precio
            if self.cesta[clave]["kg"] <= 0:
                del self.cesta[clave]
        elif kg > 0:
            self.cesta[clave] = {
                "icono": icono,
                "nombre": nombre,
                "pvp": precio,
                "kg": kg,
                "unidad": unidad,
                "total": precio * kg
            }
    def on_click_catalogo(self, event):
        selected = self.tabla_catalogo.selection()
        if selected:
            clave = selected[0]
            icono, nombre_bonito, precio, unidad = FRUTAS[clave]
            existente = self.cesta.get(clave)
            kg_acum = existente["kg"] if existente else 0.0

            delta_kg = 1.0
            self.actualizar_cesta_dict(clave, nombre_bonito, icono, precio, delta_kg, unidad)
            self.actualizar_cesta_gui()
    def agregar_bolsa(self):
        icono, nombre_bonito, precio, unidad = BOLSA_INFO
        self.actualizar_cesta_dict("bolsa", nombre_bonito, icono, precio, 1.0, unidad)
        self.actualizar_cesta_gui()
    def cancelar_pedido(self):
        if not self.cesta:
            Messagebox.show_info("La cesta ya está vacía.", title="Información", parent=self)
            return
        respuesta = Messagebox.yesno("¿Seguro que deseas cancelar el pedido actual?", title="Confirmar", parent=self)
        if respuesta == "Yes":
            self.cesta.clear()
            self.actualizar_cesta_gui()
    def procesar_ticket(self):
        if not self.cesta:
            Messagebox.show_warning("La cesta está vacía. Añade al menos un producto.", title="Atención", parent=self)
            return
        total_a_pagar = sum(item["total"] for item in self.cesta.values())
        dlg = VentanaPago(self, total_a_pagar)
        self.wait_window(dlg)
        if dlg.resultado:
            metodo_pago, entrega, cambio = dlg.resultado
            id_ticket = datetime.now().strftime("%Y%m%d_%H%M%S")
            registrar_venta_db = self._importar_reportes()[0]
            registrar_venta_db(id_ticket, self.cesta, entrega, cambio, metodo_pago)                         #AGREGADO 4: Llamar función para registrar venta en bd
            crear_pdf(self.cesta, id_ticket, entrega, cambio, metodo_pago)                                  #AGREGADO 4: HASTA AQUI 
            msg = f"Ticket Nº {id_ticket} generado correctamente en PDF.\n"
            if metodo_pago == "Efectivo":
                msg += f"Cambio a devolver: {formato_precio(cambio)}"
            Messagebox.show_info(msg, title="Venta Completada", parent=self)
            self.cesta.clear()
            self.actualizar_cesta_gui()
# ==========================================
# 5. PUNTO DE ENTRADA
# ==========================================
if __name__ == "__main__":
    app = AppTPV()
    app.mainloop()
 
 