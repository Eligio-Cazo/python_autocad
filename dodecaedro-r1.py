#python -m pip install pywin32

import math
import win32com.client
import pythoncom

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
def punto_x(X,a,radio):
    angulo_1=X
    angulo_2=[a,a,a,a,a]
    p=list(map(lambda x,y : (round(radio*math.cos(x*math.pi/180)*math.cos(y*math.pi/180),12)), angulo_1,angulo_2) )
    return p

def punto_y(X,a,radio):
    angulo_1=X
    angulo_2=[a,a,a,a,a]
    p=list(map(lambda x,y : (round(radio*math.sin(x*math.pi/180)*math.cos(y*math.pi/180),12)), angulo_1,angulo_2) )
    return p

def punto_z(a,radio):
    angulo_2=[a,a,a,a,a]
    p=list(map(lambda x : (round(radio*math.sin(x*math.pi/180),12)), angulo_2) )
    return p


def Add_dodecaedro():
    # Crea un dodecaedros solido a partir de una esfera en Autocad.
    acad = win32com.client.Dispatch("AutoCAD.Application")  
    doc = acad.ActiveDocument
    
    p1=POINT(0,0,0)  #Punto de insercion del dodecaedro
    r=1         #radio del circulo de la esfera

    #definimos las listas que contendran los puntos vertices del dodecaedro
    xs=[]
    ys=[]
    zs=[]
    xi=[]
    yi=[]
    zi=[]
    xms=[]
    yms=[]
    zms=[]
    xmi=[]
    ymi=[]
    zmi=[]

#Definimos los angulos de nuestro poliedro
#para 270 la punta del dodecaedro superior apunta hacia -Y, hay 72º de diferencia entra cada uno =360/5=72
    angulo_sup = [270,342, 54,126,198]  #Angulos del pentagono superior
    angulo_inf = [234,306,18,90,162]    #Angulos del pentagono inferior hay un defasje de 36º =72/2
    angulo_msup = [270,342,54,126,198]  # Esta alineado con angulo_superior
    angulo_minf = [306,18,90,162,234]   # Esta alineado con angulo_inferior pero se suma 72º para los indices 234+72=306 

    angulo_zeta1 = 52.62263187  #angulo superior que forma el circulo inscripto en el poliedro
    angulo_zeta2 = -52.62263187 #angulo inferior que forma el circulo inscripto en el poliedro
    angulo_zeta3 = 10.81231696  #angulo que forma el circulo inscripto en vertices del poliedro en los puntos medios superiores
    angulo_zeta4 = -10.81231696 #angulo que forma el circulo inscripto en vertices del poliedro en los puntos medios inferiores
    
    # Puntos superiores del pentagono
    xs=punto_x(angulo_sup,angulo_zeta1,r)
    ys=punto_y(angulo_sup,angulo_zeta1,r)
    zs=punto_z(angulo_zeta1,r)

    #convertir a puntos de 3 coordenadas para Usar en Autocad
    Pslice_s1=POINT(xs[0],ys[0],zs[0]) 
    Pslice_s2=POINT(xs[1],ys[1],zs[1])
    Pslice_s3=POINT(xs[2],ys[2],zs[2])
    Pslice_s4=POINT(xs[3],ys[3],zs[3])
    Pslice_s5=POINT(xs[4],ys[4],zs[4])

    # Puntos inferiores del pentagono
    xi=punto_x(angulo_inf,angulo_zeta2,r)
    yi=punto_y(angulo_inf,angulo_zeta2,r)
    zi=punto_z(angulo_zeta2,r)

    #convertir a puntos de 3 coordenadas para Usar en Autocad
    Pslice_i1=POINT(xi[0],yi[0],zi[0])
    Pslice_i2=POINT(xi[1],yi[1],zi[1])
    Pslice_i3=POINT(xi[2],yi[2],zi[2])
    Pslice_i4=POINT(xi[3],yi[3],zi[3])
    Pslice_i5=POINT(xi[4],yi[4],zi[4])
    
    # Puntos medios superiores
    xms=punto_x(angulo_msup,angulo_zeta3,r)
    yms=punto_y(angulo_msup,angulo_zeta3,r)
    zms=punto_z(angulo_zeta3,r)

    #convertir a puntos de 3 coordenadas para Usar en Autocad
    Pslice_ms1=POINT(xms[0],yms[0],zms[0])
    Pslice_ms2=POINT(xms[1],yms[1],zms[1])
    Pslice_ms3=POINT(xms[2],yms[2],zms[2])
    Pslice_ms4=POINT(xms[3],yms[3],zms[3])
    Pslice_ms5=POINT(xms[4],yms[4],zms[4])

    # Puntos medios inferiores
    xmi=punto_x(angulo_minf,angulo_zeta4,r)
    ymi=punto_y(angulo_minf,angulo_zeta4,r)
    zmi=punto_z(angulo_zeta4,r)

    #convertir a puntos de 3 coordenadas para Usar en Autocad
    Pslice_mi1=POINT(xmi[0],ymi[0],zmi[0]) 
    Pslice_mi2=POINT(xmi[1],ymi[1],zmi[1])
    Pslice_mi3=POINT(xmi[2],ymi[2],zmi[2])
    Pslice_mi4=POINT(xmi[3],ymi[3],zmi[3])
    Pslice_mi5=POINT(xmi[4],ymi[4],zmi[4])

    #dibuja puntos las 20 coordenadas del dodecaedro
    point1=doc.ModelSpace.AddPoint(Pslice_s1) 
    point2=doc.ModelSpace.AddPoint(Pslice_s2) 
    point3=doc.ModelSpace.AddPoint(Pslice_s3) 
    point4=doc.ModelSpace.AddPoint(Pslice_s4) 
    point5=doc.ModelSpace.AddPoint(Pslice_s5) 
    point6=doc.ModelSpace.AddPoint(Pslice_i1) 
    point7=doc.ModelSpace.AddPoint(Pslice_i2) 
    point8=doc.ModelSpace.AddPoint(Pslice_i3) 
    point9=doc.ModelSpace.AddPoint(Pslice_i4) 
    point10=doc.ModelSpace.AddPoint(Pslice_i5)
    point11=doc.ModelSpace.AddPoint(Pslice_ms1) 
    point12=doc.ModelSpace.AddPoint(Pslice_ms2) 
    point13=doc.ModelSpace.AddPoint(Pslice_ms3) 
    point14=doc.ModelSpace.AddPoint(Pslice_ms4) 
    point15=doc.ModelSpace.AddPoint(Pslice_ms5) 
    point16=doc.ModelSpace.AddPoint(Pslice_mi1) 
    point17=doc.ModelSpace.AddPoint(Pslice_mi2) 
    point18=doc.ModelSpace.AddPoint(Pslice_mi3) 
    point19=doc.ModelSpace.AddPoint(Pslice_mi4) 
    point20=doc.ModelSpace.AddPoint(Pslice_mi5)

    #dibujar la esfera
    esfera=doc.ModelSpace.Addsphere(p1,r) #crea la esfera
    
    borrar=bool(False) # si se pone False borra uno de los solidos

    #hay 12 planos de cortes
    #corte pentagonos laterales superiores
    sliceObj1 = esfera.SliceSolid(Pslice_mi1,Pslice_s2, Pslice_s1, borrar) #cortar el solido por un plano 3 puntos
    sliceObj2 = esfera.SliceSolid(Pslice_mi2,Pslice_s3, Pslice_s2, borrar) #cortar el solido por un plano 3 puntos
    sliceObj3 = esfera.SliceSolid(Pslice_mi3,Pslice_s4, Pslice_s3, borrar) #cortar el solido por un plano 3 puntos
    sliceObj4 = esfera.SliceSolid(Pslice_mi4,Pslice_s5, Pslice_s4, borrar) #cortar el solido por un plano 3 puntos
    sliceObj5 = esfera.SliceSolid(Pslice_mi5,Pslice_s1, Pslice_s5, borrar) #cortar el solido por un plano 3 puntos
    #corte pentagono superior
    sliceObj6 = esfera.SliceSolid(Pslice_s2,Pslice_s3, Pslice_s1, borrar) #cortar el solido por un plano 3 puntos

    #corte pentagonos laterales inferiores
    sliceObj7 = esfera.SliceSolid(Pslice_ms1,Pslice_i1, Pslice_i2, borrar) #cortar el solido por un plano 3 puntos
    sliceObj8 = esfera.SliceSolid(Pslice_ms2,Pslice_i2, Pslice_i3, borrar) #cortar el solido por un plano 3 puntos
    sliceObj9 = esfera.SliceSolid(Pslice_ms3,Pslice_i3, Pslice_i4, borrar) #cortar el solido por un plano 3 puntos
    sliceObj10 = esfera.SliceSolid(Pslice_ms4,Pslice_i4, Pslice_i5, borrar) #cortar el solido por un plano 3 puntos
    sliceObj11 = esfera.SliceSolid(Pslice_ms5,Pslice_i5, Pslice_i1, borrar) #cortar el solido por un plano 3 puntos
    #corte pentagono superior
    sliceObj12 = esfera.SliceSolid(Pslice_i1,Pslice_i3, Pslice_i2, borrar) #cortar el solido por un plano 3 puntos

def main():
    Add_dodecaedro() 

main()