#Autor Ing. Eligio Cazo 14/10/2022
#correo: ezcazo@gmail.com
#telefono: 595981500340

#python -m pip install pywin32

import math
import win32com.client
import pythoncom
import time

def Plotear_A4():
    acad = win32com.client.Dispatch("AutoCAD.Application")  
    doc = acad.ActiveDocument
    print(doc.Name)
    nombre=doc.Name

    nom=nombre[0:len(nombre) - 4]
    print(nom)

    inicio = 1
    ruta = 'C:\\acadprg\\MOPC44\\' #poner la ruta donde se van a imprimir los planos
    
    try:
        doc.SelectionSets.Item("SS1").Delete()
    except:
        a="Delete selection failed"
    
    ssget1 = doc.SelectionSets.Add("SS1")
    ssget1.SelectOnScreen()
    k=0
    z = inicio - 1
    for entity in ssget1: # Recorre todas las entidades
        time.sleep(0.05) # se incuye para que no de error , ya que python supera en velocidad a autocad en las consultas
        name = entity.EntityName
        print(name)
        #if name == "AcDb2dPolyline": para 2d polylinea 
        if name == 'AcDbPolyline': # para LWPolyline
            k=k+1
            plotfile = ruta+ nom + "-"+ str(k+z).strip() + ".pdf"
            enti_pl1 = entity
            m=entity.Coordinates
            #print(m)
            n=bounding(m)
            P1=SPOINT(n[0],n[1])
            P2=SPOINT(n[2],n[3])

            print(P1)
            print(P2)
            doc.ActiveLayout.ConfigName= "DWG To PDF.pc3"  #se puede cambiar a cualquier pc3 configurado
            doc.ActiveLayout.CanonicalMediaName = "ISO_expand_A4_(210.00_x_297.00_MM)" #debe coincidir exctamente el nombre 
            doc.ActiveLayout.SetWindowToPlot (P1,P2)
            doc.ActiveLayout.PaperUnits = 1
            doc.ActiveLayout.CenterPlot = True
            doc.Plot.QuietErrorMode = False
            doc.ActiveLayout.UseStandardScale = False
            doc.ActiveLayout.SetCustomScale(1, 1) #escala del dibujo
            doc.SetVariable('BACKGROUNDPLOT', 0)
            doc.Regen(1)
            doc.ActiveLayout.CenterPlot = True
            doc.ActiveLayout.PlotRotation = 0
            doc.ActiveLayout.StyleSheet = "monochrome.ctb" #plantilla de plumillas
            doc.ActiveLayout.PlotType = 4 #acWindow
            doc.ActiveLayout.StandardScale = 0 #acScaleToFit
            doc.Plot.PlotToFile(plotfile) #nombre del fichero donde se va imprimir
            time.sleep(5)

"""
############################## DWG to PDF canonicalMediaName ###############################3
ISO_full_bleed_B5_(250.00_x_176.00_MM)
ISO_full_bleed_B5_(176.00_x_250.00_MM)
ISO_full_bleed_B4_(353.00_x_250.00_MM)
ISO_full_bleed_B4_(250.00_x_353.00_MM)
ISO_full_bleed_B3_(500.00_x_353.00_MM)
ISO_full_bleed_B3_(353.00_x_500.00_MM)
ISO_full_bleed_B2_(707.00_x_500.00_MM)
ISO_full_bleed_B2_(500.00_x_707.00_MM)
ISO_full_bleed_B1_(1000.00_x_707.00_MM)
ISO_full_bleed_B1_(707.00_x_1000.00_MM)
ISO_full_bleed_B0_(1414.00_x_1000.00_MM)
ISO_full_bleed_B0_(1000.00_x_1414.00_MM)
ISO_full_bleed_A5_(210.00_x_148.00_MM)
ISO_full_bleed_A5_(148.00_x_210.00_MM)
ISO_full_bleed_2A0_(1189.00_x_1682.00_MM)
ISO_full_bleed_4A0_(1682.00_x_2378.00_MM)
ISO_full_bleed_A4_(297.00_x_210.00_MM)
ISO_full_bleed_A4_(210.00_x_297.00_MM)
ISO_full_bleed_A3_(420.00_x_297.00_MM)
ISO_full_bleed_A3_(297.00_x_420.00_MM)
ISO_full_bleed_A2_(594.00_x_420.00_MM)
ISO_full_bleed_A2_(420.00_x_594.00_MM)
ISO_full_bleed_A1_(841.00_x_594.00_MM)
ISO_full_bleed_A1_(594.00_x_841.00_MM)
ISO_full_bleed_A0_(841.00_x_1189.00_MM)
ARCH_full_bleed_E1_(30.00_x_42.00_Inches)
ARCH_full_bleed_E_(36.00_x_48.00_Inches)
ARCH_full_bleed_D_(36.00_x_24.00_Inches)
ARCH_full_bleed_D_(24.00_x_36.00_Inches)
ARCH_full_bleed_C_(24.00_x_18.00_Inches)
ARCH_full_bleed_C_(18.00_x_24.00_Inches)
ARCH_full_bleed_B_(18.00_x_12.00_Inches)
ARCH_full_bleed_B_(12.00_x_18.00_Inches)
ARCH_full_bleed_A_(12.00_x_9.00_Inches)
ARCH_full_bleed_A_(9.00_x_12.00_Inches)
ANSI_full_bleed_F_(28.00_x_40.00_Inches)
ANSI_full_bleed_E_(34.00_x_44.00_Inches)
ANSI_full_bleed_D_(34.00_x_22.00_Inches)
ANSI_full_bleed_D_(22.00_x_34.00_Inches)
ANSI_full_bleed_C_(22.00_x_17.00_Inches)
ANSI_full_bleed_C_(17.00_x_22.00_Inches)
ANSI_full_bleed_B_(17.00_x_11.00_Inches)
ANSI_full_bleed_B_(11.00_x_17.00_Inches)
ANSI_full_bleed_A_(11.00_x_8.50_Inches)
ANSI_full_bleed_A_(8.50_x_11.00_Inches)
ISO_expand_A0_(841.00_x_1189.00_MM)
ISO_A0_(841.00_x_1189.00_MM)
ISO_expand_A1_(841.00_x_594.00_MM)
ISO_expand_A1_(594.00_x_841.00_MM)
ISO_A1_(841.00_x_594.00_MM)
ISO_A1_(594.00_x_841.00_MM)
ISO_expand_A2_(594.00_x_420.00_MM)
ISO_expand_A2_(420.00_x_594.00_MM)
ISO_A2_(594.00_x_420.00_MM)
ISO_A2_(420.00_x_594.00_MM)
ISO_expand_A3_(420.00_x_297.00_MM)
ISO_expand_A3_(297.00_x_420.00_MM)
ISO_A3_(420.00_x_297.00_MM)
ISO_A3_(297.00_x_420.00_MM)
ISO_expand_A4_(297.00_x_210.00_MM)
ISO_expand_A4_(210.00_x_297.00_MM)
ISO_A4_(297.00_x_210.00_MM)
ISO_A4_(210.00_x_297.00_MM)
ARCH_expand_E1_(30.00_x_42.00_Inches)
ARCH_E1_(30.00_x_42.00_Inches)
ARCH_expand_E_(36.00_x_48.00_Inches)
ARCH_E_(36.00_x_48.00_Inches)
ARCH_expand_D_(36.00_x_24.00_Inches)
ARCH_expand_D_(24.00_x_36.00_Inches)
ARCH_D_(36.00_x_24.00_Inches)
ARCH_D_(24.00_x_36.00_Inches)
ARCH_expand_C_(24.00_x_18.00_Inches)
ARCH_expand_C_(18.00_x_24.00_Inches)
ARCH_C_(24.00_x_18.00_Inches)
ARCH_C_(18.00_x_24.00_Inches)
ANSI_expand_E_(34.00_x_44.00_Inches)
ANSI_E_(34.00_x_44.00_Inches)
ANSI_expand_D_(34.00_x_22.00_Inches)
ANSI_expand_D_(22.00_x_34.00_Inches)
ANSI_D_(34.00_x_22.00_Inches)
ANSI_D_(22.00_x_34.00_Inches)
ANSI_expand_C_(22.00_x_17.00_Inches)
ANSI_expand_C_(17.00_x_22.00_Inches)
ANSI_C_(22.00_x_17.00_Inches)
ANSI_C_(17.00_x_22.00_Inches)
ANSI_expand_B_(17.00_x_11.00_Inches)
ANSI_expand_B_(11.00_x_17.00_Inches)
ANSI_B_(17.00_x_11.00_Inches)
ANSI_B_(11.00_x_17.00_Inches)
ANSI_expand_A_(11.00_x_8.50_Inches)
ANSI_expand_A_(8.50_x_11.00_Inches)
ANSI_A_(11.00_x_8.50_Inches)
ANSI_A_(8.50_x_11.00_Inches)
############################## DWG to PDF canonicalMediaName ###############################3
"""
def centroide(Lista):
    x=[]
    y=[]
    centro=[]

    for i in range(0, len(Lista), 2) :
        x.append(Lista[i])
        y.append(Lista[i+1])

    xm=(max(x)+min(x))/2
    ym=(max(y)+min(y))/2
    centroide=[xm,ym]
    return centro

def bounding(Lista):
    x=[]
    y=[]
    
    for i in range(0, len(Lista), 2) :
        x.append(Lista[i])
        y.append(Lista[i+1])

    xmin=min(x)
    ymin=min(y)
    xmax=max(x)
    ymax=max(y)
    borde=[xmin,ymin,xmax, ymax]

    return borde

# pywin 32 functions///////////////////////////////////////////////////////////////////////////////////
def POINT(x, y, z):
    """Coordinate points are converted to floating point numbers""" 
    return win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, (x, y, z))

def SPOINT(x, y):
    """Coordinate points are converted to floating point numbers""" 
    return win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, (x, y))

def vtobj(obj):
    """ is converted to an object array """ 
    return win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_DISPATCH, obj)

def vtFloat(list):
    """The list is converted to a floating point number""" 
    return win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, list)
    
def vtInt(list):
    """list is converted to integer """ 
    return win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_I2, list)

def aDispatch(vObject):
    return win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_DISPATCH,(vObject))

def nullObj(vObject):
    win32com.client.VARIANT(pythoncom.VT_VARIANT | pythoncom.VT_NULL,(vObject))

# pywin 32 functions///////////////////////////////////////////////////////////////////////////////////

def main():
    Plotear_A4() 

main()