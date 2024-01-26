import tkinter as tk
from tkinter import messagebox
import datetime
import pandas as pd
import xlsxwriter

class VentanaPrincipal:
    def __init__(self, root):
        self.root = root
        self.root.title("Registro de Ventas")
        
        self.numero_venta = 1
        self.ventas = []
        
        self.label_nombre = tk.Label(root, text="Nombre:")
        self.label_nombre.grid(row=0, column=0, padx=10, pady=10, sticky=tk.W)
        self.entry_nombre = tk.Entry(root)
        self.entry_nombre.grid(row=0, column=1)
        
        self.label_producto = tk.Label(root, text="Producto:")
        self.label_producto.grid(row=1, column=0, padx=10, pady=10, sticky=tk.W)
        self.entry_producto = tk.Entry(root)
        self.entry_producto.grid(row=1, column=1)
        
        self.label_detalle = tk.Label(root, text="Detalle:")
        self.label_detalle.grid(row=2, column=0, padx=10, pady=10, sticky=tk.W)
        self.entry_detalle = tk.Entry(root)
        self.entry_detalle.grid(row=2, column=1)
        
        self.label_importe = tk.Label(root, text="Importe:")
        self.label_importe.grid(row=3, column=0, padx=10, pady=10, sticky=tk.W)
        self.entry_importe = tk.Entry(root)
        self.entry_importe.grid(row=3, column=1)
        
        self.label_fecha = tk.Label(root, text="Fecha (DD/MM/YYYY):")
        self.label_fecha.grid(row=4, column=0, padx=10, pady=10, sticky=tk.W)
        self.entry_fecha = tk.Entry(root)
        self.entry_fecha.grid(row=4, column=1)
        
        self.btn_registrar = tk.Button(root, text="Registrar Venta", command=self.registrar_venta)
        self.btn_registrar.grid(row=5, column=1, padx=10, pady=10)
        
        self.btn_balance = tk.Button(root, text="Generar Balance", command=self.generar_balance)
        self.btn_balance.grid(row=1, column=2, padx=10, pady=10)
        
        self.btn_corte = tk.Button(root, text="Generar Planilla de Corte", command=self.generar_planilla_corte)
        self.btn_corte.grid(row=2, column=2, padx=10, pady=10, sticky=tk.E)
        
        self.btn_ventas = tk.Button(root, text="Generar Planilla de Ventas", command=self.generar_planilla_ventas)
        self.btn_ventas.grid(row=3, column=2, padx=10, pady=10, sticky=tk.E)
        
    def registrar_venta(self):
        nombre = self.entry_nombre.get()
        producto = self.entry_producto.get()
        detalle = self.entry_detalle.get()
        importe = self.entry_importe.get()
        fecha_str = self.entry_fecha.get()
        
        if nombre and producto and importe and fecha_str:
            try:
                fecha = datetime.datetime.strptime(fecha_str, "%d/%m/%Y").date()
                venta = {
                    "Número de Venta": self.numero_venta,
                    "Nombre": nombre,
                    "Producto": producto,
                    "Detalle": detalle,
                    "Importe": importe,
                    "Fecha": fecha
                }
                self.numero_venta += 1
                self.ventas.append(venta)
                
                messagebox.showinfo("Venta Registrada", "Venta registrada con éxito.")
                self.limpiar_campos()
                # Aquí puedes guardar el registro de venta en una base de datos o hacer lo que necesites con los datos ingresados.
            except ValueError:
                messagebox.showwarning("Error", "Fecha inválida. Utiliza el formato DD/MM/YYYY.")
        else:
            messagebox.showwarning("Error", "Debe completar todos los campos obligatorios.")
    
    def limpiar_campos(self):
        self.entry_nombre.delete(0, tk.END)
        self.entry_producto.delete(0, tk.END)
        self.entry_detalle.delete(0, tk.END)
        self.entry_importe.delete(0, tk.END)
        self.entry_fecha.delete(0, tk.END)
    
    def generar_balance(self):
        if not self.ventas:
            messagebox.showwarning("Error", "No hay ventas registradas.")
            return
        
        # Cálculo de cantidad de ventas totales
        cantidad_ventas = len(self.ventas)
        
        # Cálculo del importe total de ventas
        importe_total = sum(float(venta["Importe"]) for venta in self.ventas)
        
        ventas_por_mes = {}
        for venta in self.ventas:
            mes = venta["Fecha"].strftime("%Y-%m")
            if mes in ventas_por_mes:
                ventas_por_mes[mes] += float(venta["Importe"])
            else:
                ventas_por_mes[mes] = float(venta["Importe"])

        # Agregar importe total de ventas por mes al diccionario ventas_por_mes
        for mes, importe in ventas_por_mes.items():
            ventas_por_mes[mes] = {
                "Importe": importe,
                "Cantidad": 0  # Inicializar la cantidad de ventas por mes en 0
            }
        
        # Cálculo del importe promedio por mes
        meses = set(venta["Fecha"].strftime("%Y-%m") for venta in self.ventas)
        cantidad_meses = len(meses)
        importe_promedio = importe_total / cantidad_meses if cantidad_meses > 0 else 0
        
        # Cálculo del ranking de los productos más vendidos
        productos_vendidos = {}
        for venta in self.ventas:
            producto = venta["Producto"]
            mes = venta["Fecha"].strftime("%Y-%m")
            if producto in productos_vendidos:
                productos_vendidos[producto] += 1
            else:
                productos_vendidos[producto] = 1
            
            # Actualizar la cantidad de ventas por mes
            ventas_por_mes[mes]["Cantidad"] += 1
        
        ranking_productos = sorted(productos_vendidos.items(), key=lambda x: x[1], reverse=True)
        
        # Generar el reporte en un archivo Excel
        workbook = xlsxwriter.Workbook("balance_ventas.xlsx")
        worksheet = workbook.add_worksheet()
        
        worksheet.write("A1", "Cantidad de Ventas Totales")
        worksheet.write("B1", cantidad_ventas)
        
        worksheet.write("A2", "Importe Total de Ventas")
        worksheet.write("B2", importe_total)
        
        worksheet.write("A3", "Importe Promedio por Mes")
        worksheet.write("B3", importe_promedio)
        
        worksheet.write("F1", "Ranking de los Productos Más Vendidos")
        row = 1
        for producto, cantidad in ranking_productos:
            worksheet.write(row, 5, producto)
            worksheet.write(row, 6, cantidad)
            row += 1
            
        # Escribir el importe total de ventas por mes
        row = 5
        worksheet.write("A"+str(row), "Importe Total de Ventas por Mes")
        worksheet.write("B"+str(row), "Mes")
        worksheet.write("C"+str(row), "Importe")
        worksheet.write("D"+str(row), "Cantidad")  # Agregar encabezado para cantidad de ventas por mes
        row += 1
        for mes, data in ventas_por_mes.items():
            worksheet.write("B"+str(row), mes)
            worksheet.write("C"+str(row), data["Importe"])
            worksheet.write("D"+str(row), data["Cantidad"])
            row += 1
        
        workbook.close()
        
        messagebox.showinfo("Balance Generado", "El balance de ventas se ha generado exitosamente.")
    
    def generar_planilla_corte(self):
        # Aquí puedes implementar la generación de la planilla de corte
        pass
    
    def generar_planilla_ventas(self):
        # Aquí puedes implementar la generación de la planilla de ventas
        pass
        
        messagebox.showinfo("Reporte Generado", "El reporte 'Ver Balance' se ha generado correctamente.")
    
    def generar_planilla_corte(self):
        if not self.ventas:
            messagebox.showwarning("Error", "No hay ventas registradas.")
            return
        
        # Generar la planilla de corte en un archivo Excel
        data = {
            "Número de Venta": [venta["Número de Venta"] for venta in self.ventas],
            "Nombre": [venta["Nombre"] for venta in self.ventas],
            "Producto": [venta["Producto"] for venta in self.ventas],
            "Detalle": [venta["Detalle"] for venta in self.ventas],
            "Fecha": [venta["Fecha"].strftime("%d/%m/%Y") for venta in self.ventas]
        }
        
        df = pd.DataFrame(data)
        df.to_excel("planilla_corte.xlsx", index=False)
        
        messagebox.showinfo("Reporte Generado", "El reporte 'Planilla de Corte' se ha generado correctamente.")
    
    def generar_planilla_ventas(self):
        if not self.ventas:
            messagebox.showwarning("Error", "No hay ventas registradas.")
            return
        
        # Generar la planilla de ventas en un archivo Excel
        data = {
            "Número de Venta": [venta["Número de Venta"] for venta in self.ventas],
            "Nombre": [venta["Nombre"] for venta in self.ventas],
            "Producto": [venta["Producto"] for venta in self.ventas],
            "Detalle": [venta["Detalle"] for venta in self.ventas],
            "Importe": [venta["Importe"] for venta in self.ventas],
            "Fecha": [venta["Fecha"].strftime("%d/%m/%Y") for venta in self.ventas]
        }
        
        df = pd.DataFrame(data)
        df.to_excel("planilla_ventas.xlsx", index=False)
        
        messagebox.showinfo("Reporte Generado", "El reporte 'Planilla de Ventas' se ha generado correctamente.")

if __name__ == "__main__":
    root = tk.Tk()
    ventana = VentanaPrincipal(root)
    root.mainloop()