#python -m pip install pywin32

import win32com.client
import pythoncom
# import pyautocad

# ////////////////////////////////////////////////////////////////////////////////////
def POINT(x, y, z):
    """Coordinate points are converted to floating point numbers""" 
    return win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, (x, y, z))

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

# ///////////////////////////////////////////////////////////////////////////////////////


def Addextrudedsolid_slice():
    # Crea un solido a partir de un recngulo en Autocad
    acad = win32com.client.Dispatch("AutoCAD.Application")  
    doc = acad.ActiveDocument  
    
    p1=[0,0,0]  #Punto de insercion de la region
    b1=0.30     #base del rectangulo
    h1=0.50     #altura del rectangulo

    p01 = POINT(p1[0] + b1 / 2,p1[1] - h1 / 2,0) #Funcion POINT transforma vector a variant double precision 
    p02 = POINT(p1[0] + b1 / 2,p1[1] + h1 / 2,0)
    p03 = POINT(p1[0] - b1 / 2,p1[1] + h1 / 2,0)
    p04 = POINT(p1[0] - b1 / 2,p1[1] - h1 / 2,0)
    
    lin1 = doc.ModelSpace.AddLine(p01, p02)
    lin2 = doc.ModelSpace.AddLine(p02, p03)
    lin3 = doc.ModelSpace.AddLine(p03, p04)
    lin4 = doc.ModelSpace.AddLine(p04, p01)

    
    curves =[lin1,lin2,lin3,lin4] #debe ser un area cerrada

    curves=aDispatch(curves) #cambia a objeto win32com.client

    region=doc.ModelSpace.AddRegion(curves) #crea la region

        
    # ****************** crear el solido via sendcommand *******************************
    #comsend='(command "extrude" "l" ""' +'"'+str(lon)+'" "") '
    #doc.SendCommand(comsend)
    # ****************** crear el solido via sendcommand *******************************
    
    try:
        doc.SelectionSets.Item("SS2").Delete() # Borra si ya existe conjunto de seleccion
    except:
        print("Seleccion SS2 aun no creada")
    
    ssget1 = doc.SelectionSets.Add("SS2") #crea conjunto de seleccion si no existe")
    ssget1.Select(4) #selecciona ultima entidad dibujada
    
    obj = ssget1[0] #toma el objeto en particular
    print(obj.layer) #imprime el layer solo es informativo

    for entity in ssget1: # buscamos la region recien creada
        name = entity.EntityName
    if name == 'AcDbRegion':
        #print(entity)
        # usar AddExtrudedSolid(entidad region variable acadentity, altura de extrusión variable float , angulo float)
        altura=10.0
        solidObj=doc.ModelSpace.AddExtrudedSolid(entity, altura, 0.0) #crea el solido con 10 m de longitud
        Pslice1= POINT(-3.914, 3.43,0.50) #definimos punto1 del plano para cortar el solido
        Pslice2= POINT(4.812,2.61,1.50) #definimos punto2 del plano para cortar el solido
        Pslice3= POINT(0.392,-2.79,3.0) #definimos punto3 del plano para cortar el solido
        borrar=bool(True) # si se pone False borra uno de los solidos

      
        sliceObj1 = solidObj.SliceSolid(Pslice1,Pslice2, Pslice3, borrar) #cortar el solido por un plano que pasa por los 3 puntos


def main():
    Addextrudedsolid_slice() 
       
main()