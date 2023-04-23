from fpdf import FPDF
import matplotlib.pyplot as plt
import matplotlib.ticker as ticker
import datetime

footer = 'resources/images/footer.png'
logo = 'resources/images/logo.png'
color_fondo = 'resources/images/color.png'

nombre, apellido = "Rudy", "Velasquez"
genero = ""
telefono_cliente = "123456789"

nombre_asesor = "Carlos Estrada"
telefono_asesor = "987654321"

tipo_de_proyecto = "Industria"

def select_font(font_name = None):
    if font_name == "h1":
        pdf.set_font('RalewayB', "", 14)
    elif font_name == "sub":
        pdf.set_font('Raleway', "", 8)
    elif font_name == "subsub":
        pdf.set_font('Raleway', "", 7)
    elif font_name == "b":
        pdf.set_font('RalewayB', "", 11)
    else:
        pdf.set_font('Raleway', "", 11)
        
# Inicializar
pdf = FPDF(orientation = 'P', unit = 'mm', format = 'letter')
pdf.add_font('Raleway', '', 'resources/fonts/Raleway-Regular.ttf', uni = True)
pdf.add_font('RalewayB', '', 'resources/fonts/Raleway-Bold.ttf', uni = True)
pdf.set_text_color(0, 0, 0)
pdf.l_margin = 29.97
pdf.r_margin = 29.97
pdf.t_margin = 24.89
pdf.d_margin = 24.89

#* Datos de formulario


#* Datos guardados editables
no_cuenta_de_empresa = "123-4567890-9"
banco_de_empresa = "Banco Industrial"

#* Datos del excel

generacion_mensual_promedio = 970

departamento = "Guatemala"
produccion_fotovoltaica = 8800

cantidad_paneles = 99
potencia_paneles = 550
area_total_modulo = 41

# Descripcion
inversion_total = 99528.00

# Cuotas
precio_cuotas1 = 999
precio_cuotas2 = 999
precio_cuotas3 = 999


# 1
def portada():
    pdf.add_page()
    pdf.image('resources/pages/portada.png', 0, 0, 216, 279)

# 2
def first():
    
    MEDIDA = "kWh"

    pdf.add_page()
    pdf.image(footer, -8, 242, 223, 40)
    pdf.image(logo, 29.71, 24.89, 71.62, 22.86)
    pdf.image('resources/images/table.png', 29.97, 51, 156.02, 20.32, type = 'PNG')

    today = datetime.date.today()
    meses = ['enero','febrero','marzo','abril','mayo','junio','julio',
            'agosto','septiembre','octubre','noviembre','diciembre']

    select_font()

    pdf.set_y(41)
    pdf.set_x(104)
    pdf.write(5, f'Guatemala, {today.day} de {meses[today.month-1]} {today.year}')


    # ~ Tabla

    pdf.set_y(51)
    pdf.set_text_color(255, 255, 255)
    select_font("b")
    pdf.cell(25, 6, "Cliente", 0, 0, 'R')
    pdf.set_text_color(0, 0, 0)
    select_font()
    pdf.cell(49, 6, f" {nombre} {apellido}", 0, 0, 'L')
    pdf.cell(22, 6, "Teléfono", 0, 0, 'C')
    pdf.cell(59, 6, f" {telefono_cliente}", 0, 1, 'L')

    pdf.set_text_color(255, 255, 255)
    select_font("b")
    pdf.cell(25, 6, "Asesor", 0, 0, 'R')
    pdf.set_text_color(0, 0, 0)
    select_font()
    pdf.cell(49, 6, f" {nombre_asesor}", 0, 0, 'L')
    pdf.cell(22, 6, "Teléfono", 0, 0, 'C')
    pdf.cell(59, 6, f" {telefono_asesor}", 0, 1, 'L')

    pdf.set_text_color(255, 255, 255)
    select_font("b")
    pdf.cell(25, 6, "Proyecto de", 0, 0, 'R')
    pdf.set_text_color(0, 0, 0)
    select_font()
    pdf.cell(49, 6, f" {tipo_de_proyecto}", 0, 0, 'L')
    pdf.cell(22, 6, "Correo", 0, 0, 'C')
    pdf.cell(59, 6, f" ecolumen@solarpower.com.gt", 0, 1, 'L')

    # ~ Cuerpo

    string = (f"\nSeñor{genero}\n{nombre} {apellido}\nPresente.\n\nEstimad{genero if genero else 'o'} Sr{genero}. {apellido}:\n\n\
    Gracias por su fina atención y tomarnos en cuenta_a_transferencias para la propuesta de energía solar en su compañía con el cual estará reduciendo considerablemente el consumo de energía actual.\n\n\
    A continuación, le detallo una lista de información que nos gustaría que revise detalladamente:\n\n\
        1.    Estudio de producción de energía según coordenada del sitio de instalación.\n\
        2.    Análisis de su factura eléctrica y composición de esta.\n\
        3.    Análisis de ahorro con la implementación de la planta solar.\n\
        4.    Descripción del proyecto por rubros.\n\
        5.    Políticas de instalación, forma de pago y garantías.\n\n\n\
    Con nuestro sistema de Auto generador conectado a la red, usted estará generando mensualmente en promedio ")

    pdf.write(5, string)

    # ~ Dato
    select_font("b")
    pdf.write(5, f"{generacion_mensual_promedio}{MEDIDA}\n\n")

    # ~ Firma
    select_font()
    pdf.write(5, f"Quedamos a la espera de sus comentarios, nos queda pendiente la revisión de campo en el proyecto y una reunión para conversar detalles y alcances de este, queremos expresarle que cuenta_a_transferencias con todo nuestro apoyo y respaldo para lograr la realización del proyecto.\n\
                                                                            Atentamente,\n\n\n\
                                                                            {nombre_asesor}\n\
                                                                            Consultor en Paneles Solares\n\
                                                                            Ecolumen®")

# 3
def ubicacion():
    
    energia_anual = 12289
    
    energia_por_mes = {"Ene": 899,
                    "Feb": 964,
                    "Mar": 1023,
                    "Abr": 1087,
                    "May": 1145,
                    "Jun": 1020,
                    "Jul": 1089,
                    "Ago": 1154,
                    "Sep": 1010,
                    "Oct": 990,
                    "Nov": 910,
                    "Dic": 840}
    
    pdf.add_page()

    fig, ax = plt.subplots(figsize=(10, 6))
    bars = ax.bar(energia_por_mes.keys(), energia_por_mes.values(), color = "#41ab42", zorder=3, width=0.7)
    for bar in bars:
        height = bar.get_height()
        ax.text(bar.get_x() + bar.get_width() / 2, height, int(height), ha='center', va='bottom')
    ax.grid(axis="y", linestyle="-", alpha=0.5, zorder=0)
    plt.savefig("resources/graphs/bargraph.png", dpi = 100, bbox_inches = "tight")
    plt.close()

    depas = {
        "Guatemala": "resources/images/Departamento/guat.png",
        "Alta Verapaz": "resources/images/Departamento/alta.png",
        "Baja Verapaz": "resources/images/Departamento/baja.png",
        "Chimaltenango": "resources/images/Departamento/chim.png",
        "Chiquimula": "resources/images/Departamento/chiq.png",
        "Escuintla": "resources/images/Departamento/esc.png",
        "Huehuetenango": "resources/images/Departamento/hue.png",
        "Izabal": "resources/images/Departamento/iza.png",
        "Jalapa": "resources/images/Departamento/jala.png",
        "Jutiapa": "resources/images/Departamento/juti.png",
        "San Marcos": "resources/images/Departamento/marc.png",
        "Petén": "resources/images/Departamento/peten.png",
        "El Progreso": "resources/images/Departamento/prog.png",
        "Quetzaltenango": "resources/images/Departamento/quetz.png",
        "Quiché": "resources/images/Departamento/quiche.png",
        "Retalhuleu": "resources/images/Departamento/reu.png",
        "Santa Rosa": "resources/images/Departamento/rosa.png",
        "Sacatepéquez": "resources/images/Departamento/saca.png",
        "Sololá": "resources/images/Departamento/solo.png",
        "Suchitepéquez": "resources/images/Departamento/suchi.png",
        "Totonicapán": "resources/images/Departamento/toto.png",
        "Zacapa": "resources/images/Departamento/zaca.png"
    }

    pdf.image(footer, -8, 242, 223, 40)

    select_font('h1')

    pdf.cell(0, 10, "UBICACIÓN GEORGRÁFICA DEL PROYECTO", 0, 1, 'C')
    pdf.image(f'{depas[departamento]}', x = 88,w=40,h=45)
    pdf.cell(0, 8, "", 0, 1, 'C')        
                
    pdf.cell(0, 8, "ESTUDIO DE PRODUCCIÓN ANUAL DE ENERGÍA EN UBICACIÓN", 0, 1, 'C')                                                             

    
    pdf.set_x(64.5)

    pdf.set_text_color(255, 255, 255)
    pdf.set_fill_color(65, 171, 66)
    select_font('b')
    pdf.cell(87.5, 5, "RESUMEN", 0, 2, 'C', True)
    pdf.set_text_color(0, 0, 0)
    select_font()
    pdf.cell(50, 5, "Producción Fotovoltaica", "RB", 0, 'L')
    pdf.cell(37.5, 5, f"{produccion_fotovoltaica} watts", "LB", 1, 'L')
    pdf.set_x(64.5)
    pdf.cell(50, 5, "Módulo", "RB", 0, 'L')
    pdf.cell(37.5, 5, f"{cantidad_paneles/2}", "LB", 2, 'L')
    pdf.set_x(64.5)
    pdf.cell(50, 5, "Área total de los módulos", "RB", 0, 'L')
    pdf.cell(37.5, 5, f"{area_total_modulo} m²", "LB", 2, 'L')
    pdf.set_x(64.5)
    pdf.cell(50, 5, "Energía Anual", "TR", 0, 'L')
    pdf.cell(37.5, 5, f"{energia_anual} Kwh", "TL", 1, 'L')

    pdf.cell(0, 3, "", 0, 1, 'C') 
    select_font('h1')
    pdf.cell(0, 10, "PRODUCCIÓN DE ENERGÍA POR MES*", 0, 1, 'C')
    pdf.cell(0, 3, "", 0, 1, 'C') 
    pdf.image('resources/graphs/bargraph.png',w=156,h=80)

    select_font('sub')

    pdf.ln(2)
    pdf.cell(0, 3, "* Estimación preliminar en base a los datos satelitales de la NASA de los últimos 10 años a 14 grados", 0, 1, 'C')
    pdf.cell(0, 3, "Si su proyecto es comercial o industrial favor solicitar visitar técnica a su nombre_asesor.", 0, 1, 'C')
    pdf.cell(0, 3, " Las pérdidas son calculadas de acuerdo con la visita técnica.", 0, 1, 'C')

# 4
def consumo():
        
    contadores = 1
    k_cons_mes = 415
    costo_k = 1.00
    cargo_fijo = 112.50
    energia = 413.99
    pot_c = 370.47
    pot_max = 105.54
    excedente = 0
    distri = 0
    iva = 197.49
    total = 1199.99

    roi = 33
    ah_men = 1111.85
    ah_ano = 13342.88

    CO = 0.0007
    CAR = 0.0002
    TREE = 0.012

    ahorro = {"Ahorro en 5 años": ah_ano*5, "Ahorro en 10 años": ah_ano*10, "Ahorro en 20 años": ah_ano*20, "Ahorro en 30 años": ah_ano*30}


    pdf.add_page()

    pdf.image(footer, -8, 242, 223, 40)
    pdf.image(color_fondo, 0, 0, 216, 98)
    select_font('h1')
    pdf.cell(0, 10, "DESGLOSE DE LA FACTURA DE CONSUMO ELÉCTRICA ACTUAL", 0, 1, 'C')

    pdf.ln(4)
    pdf.set_x(64.5)
    pdf.set_text_color(255, 255, 255)
    pdf.set_fill_color(87,87,87)
    select_font('b')
    pdf.cell(87.5, 4, "ACTUAL", 0, 2, 'C', True)
    pdf.set_text_color(0, 0, 0)

    select_font()
    pdf.set_font_size(9)
    pdf.cell(60, 4, "Contadores", 0, 0, 'R')
    pdf.cell(27.5, 4, f"{contadores if contadores else '-'}", 0, 1, 'L')
    pdf.set_x(64.5)
    pdf.cell(60, 4, "Kwh consumidos mensualmente", 0, 0, 'R')
    pdf.cell(27.5, 4, f"Q {k_cons_mes if k_cons_mes else '-'}", 0, 2, 'L')
    pdf.set_x(64.5)
    pdf.cell(60, 4, "Costo del kwh", 0, 0, 'R')
    pdf.cell(27.5, 4, f"Q {costo_k if costo_k else '-'}", 0, 2, 'L')
    pdf.set_x(64.5)
    pdf.cell(60, 4, "Cargo fijo", 0, 0, 'R')
    pdf.cell(27.5, 4, f"Q {cargo_fijo if cargo_fijo else '-'}", 0, 2, 'L')
    pdf.set_x(64.5)
    pdf.cell(60, 4, "Energía", 0, 0, 'R')
    pdf.cell(27.5, 4, f"Q {energia if energia else '-'}", 0, 2, 'L')
    pdf.set_x(64.5)
    pdf.cell(60, 4, "Potencia contratada", 0, 0, 'R')
    pdf.cell(27.5, 4, f"Q {pot_c if pot_c else '-'}", 0, 2, 'L')
    pdf.set_x(64.5)
    pdf.cell(60, 4, "Potencia máxima", 0, 0, 'R')
    pdf.cell(27.5, 4, f"Q {pot_max if pot_max else '-'}", 0, 2, 'L')
    pdf.set_x(64.5)
    pdf.cell(60, 4, "Excedente de produccion_fotovoltaica", 0, 0, 'R')
    pdf.cell(27.5, 4, f"Q {excedente if excedente else '-'}", 0, 2, 'L')
    pdf.set_x(64.5)
    pdf.cell(60, 4, "Distribución", 0, 0, 'R')
    pdf.cell(27.5, 4, f"Q {distri if distri else '-'}", 0, 2, 'L')
    pdf.set_x(64.5)
    pdf.cell(60, 4, "IVA y alumbrado público", 0, 0, 'R')
    pdf.cell(27.5, 4, f"Q {iva if iva else '-'}", 0, 2, 'L')
    select_font('b')
    pdf.set_font_size(9)
    pdf.set_x(64.5)
    pdf.set_text_color(255, 255, 255)
    pdf.set_fill_color(65, 171, 66)
    pdf.cell(60, 4, "Total pago promedio", 0, 0, 'R',True)
    pdf.cell(27.5, 4, f"Q {total}", 0, 2, 'L', True)

    pdf.ln(20)

    select_font('h1')
    pdf.set_text_color(0,0,0)
    pdf.cell(0, 10, "AHORRO ESTIMADO CON PLANTA DE PRODUCCIÓN SOLAR", 0, 1, 'C')

    pdf.set_x(64.5)
    pdf.set_text_color(255, 255, 255)
    pdf.set_fill_color(87,87,87)
    select_font('b')
    pdf.cell(87.5, 4, "PROYECTO", 0, 2, 'C', True)
    pdf.set_text_color(0, 0, 0)

    select_font()
    pdf.set_font_size(9)
    pdf.cell(60, 4, "Paneles", 0, 0, 'R')
    pdf.cell(27.5, 4, f"{cantidad_paneles} x {potencia_paneles}", 0, 1, 'R')
    pdf.set_x(64.5)
    pdf.cell(60, 4, "ROI", 0, 0, 'R')
    pdf.cell(27.5, 4, f"{roi} %", 0, 2, 'R')
    pdf.set_x(64.5)
    pdf.cell(60, 4, "Ahorro mensual", 0, 0, 'R')
    pdf.cell(27.5, 4, f"Q {ah_men}", 0, 2, 'R')
    select_font('b')
    pdf.set_font_size(9)
    pdf.set_x(64.5)
    pdf.set_text_color(255, 255, 255)
    pdf.set_fill_color(65, 171, 66)
    pdf.cell(60, 4, "Ahorro Anual", 0, 0, 'R',True)
    pdf.cell(27.5, 4, f"Q {ah_ano}", 0, 2, 'R', True)

    pdf.ln(4)

    fig, ax = plt.subplots(figsize=(10, 6))

    # Create the bar chart with thinner bars
    bars = ax.bar(ahorro.keys(), ahorro.values(), color = "#41ab42", zorder=3, width=0.6)

    # Add horizontal grid lines to the background
    ax.grid(axis="y", linestyle="-", alpha=0.5, zorder=0)

    # Format the tick labels on the y-axis
    def fmt(x, pos):
        if x >= 100000000:
            val = x / 100000000
            return f"Q {val:,.0f}00M"
        else:
            return f"Q {x:,.0f}"

    ax.yaxis.set_major_formatter(ticker.FuncFormatter(fmt))

    # Save the figure
    plt.savefig("resources/graphs/bargraph2.png", dpi = 100, bbox_inches = "tight")
    plt.close()

    pdf.image('resources/graphs/bargraph2.png',w=156,h=70)

    pdf.image('resources/images/ic8.png',x=30 ,y=220,w=15, h=15)

    pdf.image('resources/images/ic7.png',x=88,y=221, w=20, h=13)

    pdf.image('resources/images/ic6.png',x=142,y=220, w=13, h=15)

    select_font()
    pdf.set_font_size(30)
    pdf.set_text_color(0,0,0)

    co = format(produccion_fotovoltaica*CO, '.1f')
    car = format(produccion_fotovoltaica*CAR, '.1f')
    tree = format(produccion_fotovoltaica*TREE, '.1f')

    pdf.text(49, 232, f"{co}")
    pdf.text(113, 232, f"{car}")
    pdf.text(159, 232, f"{tree}")

    select_font('sub')
    pdf.text(49, 237, f"Dióxido de carbono")
    pdf.text(113, 237, f"Carros Particulares")
    pdf.text(159, 237, f"Árboles plantados")

    pdf.set_text_color(100,100,100)
    select_font('subsub')
    pdf.text(49, 240, f"Toneladas métricas no emitidas")
    pdf.text(113, 240, f"Retirados por 1 año")
    pdf.text(159, 240, f"Crecimiento a 10 años")
   
# 5
def descripcion():
    
    pdf.add_page()
    pdf.image(footer, -8, 242, 223, 40) 

    select_font('h1')
    pdf.set_text_color(0,0,0)
    pdf.cell(0, 10, "DESCRIPCIÓN DEL PROYECTO", 0, 1, 'L')

    select_font()
    pdf.write(5,f"Potencia a instalar: {produccion_fotovoltaica} W")
    pdf.ln()
    select_font('b')
    pdf.cell(121.06, 5, "  Tecnología", 1, 0, 'L')
    pdf.cell(35, 5, "Inversión¹", 1, 1, 'C')

    select_font()
    pdf.cell(121.06, 5, "  Descripción de producto", 1, 0, 'L')
    pdf.cell(35, 5, "", "TLR", 1, 'C')

    pdf.cell(121.06, 5, f"  {cantidad_paneles} paneles solares {potencia_paneles}w", 1, 0, 'L')
    pdf.cell(35, 5, "", "LRD", 1, 'C')

    pdf.cell(121.06, 5, "", "TD", 0, 'L')
    pdf.cell(35, 5, "", "TD", 1, 'C')

    pdf.cell(121.06, 5, "  Instalación²", 1, 0, 'L')
    pdf.cell(35, 5, "", "TLR", 1, 'C')

    pdf.cell(121.06, 5, "  Diseño arquitectónico", 1, 0, 'L')
    pdf.cell(35, 5, "", "LR", 1, 'C')

    pdf.cell(121.06, 5, "  Sistema eléctrico³", 1, 0, 'L')
    pdf.cell(35, 5, "", "LR", 1, 'C')

    pdf.cell(121.06, 5, "  Sistema Estructural", 1, 0, 'L')
    pdf.cell(35, 5, "", "LR", 1, 'C')

    pdf.cell(121.06, 5, "  Protección DC y AC", 1, 0, 'L')
    pdf.cell(35, 5, "", "LR", 1, 'C')

    pdf.cell(121.06, 5, "  Tierra Física", 1, 0, 'L')
    pdf.cell(35, 5, "", "LR", 1, 'C')

    pdf.cell(121.06, 5, "  Mano de obra", 1, 0, 'L')
    pdf.cell(35, 5, "", "LR", 1, 'C')

    pdf.cell(121.06, 5, "  Trámites ante Empresa eléctrica para autorización del sistema", 1, 0, 'L')
    pdf.cell(35, 5, "", "LR", 1, 'C')

    pdf.cell(121.06, 5, "  Solicitud de Contador Bidireccional", 1, 0, 'L')
    pdf.cell(35, 5, "", "LR", 1, 'C')

    pdf.cell(121.06, 5, "  Seguimiento ante Empresa eléctrica después de la instalación", 1, 0, 'L')
    pdf.cell(35, 5, "", "DLR", 1, 'C')

    pdf.cell(121.06, 5, "", "TD", 0, 'L')
    pdf.cell(35, 5, "", "TD", 1, 'C')

    select_font('b')
    pdf.cell(121.06, 5, "  Inversión total con llave en mano⁴", 1, 0, 'L')
    pdf.cell(35, 5, f"Q {inversion_total}", 1, 1, 'C')

    select_font()
    pdf.write(5, "¹Incluye IVA\n")
    pdf.write(4, "²No incluye reparaciones a la red eléctrica nacional\n")
    pdf.write(4, "³No incluye subestacion eléctrica en caso de requerirla\n")
    select_font('b')
    pdf.write(4, "⁴Validez de la oferta 15 DIAS")
    pdf.ln()
    pdf.ln()

    pdf.image('resources/images/panel.png', w=156.06)

# 6
def firmas():
    pdf.add_page()
    pdf.image('resources/pages/page_firmas.png', 0, 0, 216, 279)
    pdf.image(footer, -8, 242, 223, 40)  

# 7
def cuotas():
    pdf.add_page()
    pdf.image('resources/pages/page_pago.png', 0, 0, 216, 279,"PNG")
    pdf.image(footer, -8, 242, 223, 40)
    select_font()
    pdf.text(109,102.65,f"{no_cuenta_de_empresa} del {banco_de_empresa}")
    pdf.set_xy(50,117)
    pdf.cell(21, 11, "48", "BR", 0, "C")
    pdf.cell(52, 11, "Visacuotas BI", "LBR", 0, "C")
    pdf.cell(41, 11, f"  Q {precio_cuotas1}", "LB", 1, "L")
    pdf.set_x(50)
    pdf.cell(21, 11, "36", "TRB", 0, "C")
    pdf.cell(52, 11, "Visacuotas BAC", 1, 0, "C")
    pdf.cell(41, 11, f"  Q {precio_cuotas2}", "LTB", 1, "L")
    pdf.set_x(50)
    pdf.cell(21, 11, "36", "TR", 0, "C")
    pdf.cell(52, 11, "Visacuotas BAC", "TLR", 0, "C")
    pdf.cell(41, 11, f"  Q {precio_cuotas3}", "TL", 1, "L")
 
# 8
def guia_potencia_electro():
    pdf.add_page()
    pdf.image('resources/pages/page_electro.png', 0, 0, 216, 279)
    pdf.image(footer, -8, 242, 223, 40)

# 9                                                  
def guia_potencia_vatios():
    pdf.add_page()     
    pdf.image('resources/pages/page_vatios.png', 0, 0, 216, 279)  
    pdf.image(footer, -8, 242, 223, 40)                   

import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
from tkinter import messagebox 
import os

# Function to browse Excel file
def browse_excel():
    global excel_file
    excel_file = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
    if not excel_file:
        return
    excel_entry.config(state=tk.NORMAL)
    excel_entry.delete(0, tk.END)
    excel_entry.insert(tk.END, excel_file)
    excel_entry.config(state=tk.DISABLED)

# Function to check if all fields are filled
def check_fields(*args):
    for widget in [nombre_entry, apellido_entry, gender_combobox, telefono_entry, asesor_entry, telefono_asesor_entry, tipo_proyecto_entry]:
        if isinstance(widget, ttk.Entry) and not widget.get().strip():
            generate_button.config(state=tk.DISABLED)
            return
        elif isinstance(widget, ttk.Combobox) and not widget.get():
            generate_button.config(state=tk.DISABLED)
            return
    generate_button.config(state=tk.NORMAL)

import time
# Function to generate PDF
def generate_pdf():
    global nombre, apellido, genero, telefono_cliente, nombre_asesor, telefono_asesor, tipo_de_proyecto
    # Call your 9 functions here to generate the PDF
    nombre, apellido = nombre_entry.get(), apellido_entry.get()
    genero = "a" if gender_combobox.get() == "Femenino" else ""
    telefono_cliente = telefono_entry.get()

    nombre_asesor = asesor_entry.get()
    telefono_asesor = telefono_asesor_entry.get()

    tipo_de_proyecto = tipo_proyecto_entry.get()
        
    progress_bar["value"] = 0
    step = 100 / 9
    
    for i in [portada, first, ubicacion, consumo, descripcion, firmas, cuotas, guia_potencia_electro, guia_potencia_vatios]:
        i()
        progress_bar["value"] += step
        time.sleep(0.05)
        root.update_idletasks()  
    
    pdf.output('Automated PDF Report.pdf')
    messagebox.showinfo("PDF Generated", "The PDF has been successfully generated!")

root = tk.Tk()
root.title("PDF Generator")

# Create the labels and entry widgets for each field, except Gender
fields = ["Nombre", "Apellido", "a","Telefono", "Asesor", "Telefono_Asesor", "Tipo_Proyecto"]

for idx, field in enumerate(fields, 1):

    ttk.Label(root, text=field).grid(row=idx, column=0, sticky=tk.W)
    entry = ttk.Entry(root)


    entry.grid(row=idx, column=1, sticky=tk.W)
    entry.bind("<KeyRelease>", check_fields)
    globals()[f"{field.lower().replace(' ', '_')}_entry"] = entry

# Create the dropdown menu for Gender
ttk.Label(root, text="Genero").grid(row=3, column=0, sticky=tk.W)
gender_combobox = ttk.Combobox(root, values=["Masculino", "Femenino", "Other"], state="readonly")
gender_combobox.grid(row=3, column=1, sticky=tk.W)
gender_combobox.bind("<<ComboboxSelected>>", check_fields)

# Create the Excel file entry and 'Open' button
ttk.Label(root, text="Excel File").grid(row=8, column=0, sticky=tk.W)
excel_entry = ttk.Entry(root)
excel_entry.insert(tk.END, "Cotizador.xlsx")
excel_entry.config(state=tk.DISABLED)
excel_entry.grid(row=8, column=1, sticky=tk.W)
excel_button = ttk.Button(root, text="Open", command=browse_excel)
excel_button.grid(row=8, column=2, sticky=tk.W)

# Create the 'Generate PDF' button
generate_button = ttk.Button(root, text="Generate PDF", command=generate_pdf, state=tk.DISABLED)
generate_button.grid(row=9, columnspan=2)

# Create the progress bar
progress_bar = ttk.Progressbar(root, orient=tk.HORIZONTAL, length=200, mode="determinate")
progress_bar.grid(row=10, columnspan=2, pady=10)

root.mainloop()

