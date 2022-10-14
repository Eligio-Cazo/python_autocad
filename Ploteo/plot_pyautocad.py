#en proceso 13/10/2022
#python -m pip install pywin32

import math
import win32com.client
import pythoncom
import time
import array
from pyautocad import Autocad, APoint

def Plotear_planilla_pyautocad_A4():
    acad = Autocad(create_if_not_exists=True)
    print (acad.doc.Name)
    doc = acad.ActiveDocument
    print(doc.Name)
    nombre=doc.Name


    #nomb = "000" & (Left(ActiveDocument.Name, Len(ActiveDocument.Name) - 4)) & "R0"
    #MsgBox (Left(ActiveDocument.Name, Len(ThisDrawing.Name) - 4))
    nom=nombre[0:len(nombre) - 4]
    print(nom)

    empiezaen = 1
    z = empiezaen - 1
        
    pointp1=[]
    pointp2=[]
    frame=[]
    point1=[]
    point2=[]
         
    try:
        doc.SelectionSets.Item("SS1").Delete()
    except:
        a="Delete selection failed"
        #print (a)
    
    ssget1 = doc.SelectionSets.Add("SS1")
    ssget1.SelectOnScreen()
    #print(ssget1)

    for entity in ssget1: # buscamos bloques con atributos
        time.sleep(0.05) # se incuye para que no de error , ya que python supera a autocad en las consultas
        name = entity.EntityName
        print(name)
        #if name == "AcDb2dPolyline":
        if name == 'AcDbPolyline':
            enti_pl1 = entity
            enti_pl1.GetBoundingBox (point1,point2)
            
            m=entity.Coordinates
            #print(m)
            n=bounding(m)
            
            pointp1=(n[0],n[1])
            pointp2=(n[2],n[3])
            
            print (point1)
            print (point2)
            print (pointp1)
            print (pointp2)

            #P1=APoint(point1) genera error
            #P2=APoint(point2) genera error
            
            doc.ActiveLayout.ConfigName= "DWG To PDF.pc3"
            doc.ActiveLayout.CanonicalMediaName = "ISO_expand_A4_(210.00_x_297.00_MM)"
            doc.ActiveLayout.PaperUnits = 1
            doc.ActiveLayout.CenterPlot = True
            doc.Plot.QuietErrorMode = False
            doc.ActiveLayout.UseStandardScale = False
            doc.ActiveLayout.SetCustomScale(1, 1)
            doc.SetVariable('BACKGROUNDPLOT', 0)
            doc.Regen(1)
            doc.ActiveLayout.CenterPlot = True
            doc.ActiveLayout.PlotRotation = 0
            doc.ActiveLayout.StyleSheet = "monochrome.ctb"
            #doc.ActiveLayout.SetWindowToPlot (P1,P2) genera error
            #doc.ActiveLayout.SetWindowToPlot (array.array('d', point1), array.array('d',  point2)) genera error
            #doc.ActiveLayout.SetWindowToPlot (array.array('d', pointp1), array.array('d',  pointp2)) genera error

def centroide(Lista):
    x=[]
    y=[]
    centro=[]

    for i in range(0, len(Lista), 2) :
        x.append(Lista[i])
        y.append(Lista[i+1])

    xm=(max(x)+min(x))/2
    ym=(max(y)+min(y))/2
    centro=[xm,ym]
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




def main():
    Plotear_planilla_pyautocad_A4()

main()