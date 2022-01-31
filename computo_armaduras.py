#python -m pip install pywin32

from dataclasses import replace
import win32com.client
acad = win32com.client.Dispatch("AutoCAD.Application")
import pythoncom
import time

doc = acad.ActiveDocument # Document object

#seleccionar en pantalla()
def reconoce(texto,P1,h):
    ht=h
    bb=""
    P=P1
    for i in texto:
        k=i
        if ord(k)==32:
            continue
        elif k=="C":
            bb=bb+"c"
        else:
            bb=bb+k
    #next
    comp=0
    try:
        h1=bb.split("%%")
        s1=h1[0]
        s2=h1[1]
        h2=s2.split("-")
        can=s1
        lg=h2[1]
        h3=h2[0].split("c")
        #print("h2",h2[0])
        #print("h3",h3)
        dia=h3[1]
        #print (can)
        #print("cant = " + can + " diam = " + dia + "  long= " + lg)
        comp=float(can)*float(lg)
    except:
        comp=0

    if comp!=0 and dia!="":
        instxt(dia,str(comp),P,ht)


def reconoce_previo(texto):
    bb=""
    txt=""
    
    for i in texto:
        k=i
        if ord(k)==32:
            continue
        elif k=="C":
            bb=bb+"c"
        else:
            bb=bb+k

    texto=bb
    #print(texto)
    fi0= texto.find('Ø')
    if fi0!=-1:
        texto=texto.replace("Ø","%%c")
        print(texto)
    
    zipe= texto.find('Pos')
    if zipe!=-1:
        txt=texto.split('_')
        tx2=txt[1]
        tx3=tx2.replace("(","-")
        texto=tx3.replace(")","")

    normal= texto.find('%%c')

    fi1= texto.find('fide')
    if fi1>0:
        texto=texto.replace("fide","%%c")

    fi2= texto.find('fi')
    if fi2>0:
        texto=texto.replace("fi","%%c")

    #print (zipe, normal, fi1,fi2)
    
    if zipe!=-1 or normal!=-1 or fi1!=-1 or fi2!=-1:
        txt=texto
    else: 
        txt=""

    #print (txt)

    return txt
    

def computa_texto():

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
        if name == 'AcDbText':
            #print(entity.textString)
            #print(entity.insertionPoint)
            entity.Height
            txt1=reconoce_previo(entity.textString)
            if txt1 !="":
                #reconoce(entity.textString,entity.insertionPoint)
                reconoce(txt1,entity.insertionPoint,entity.Height)

def POINT(x,y,z):
    return win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, (x,y,z))  

def instxt(diam,cantidad,P1,h):
    acad = win32com.client.Dispatch("AutoCAD.Application")  
    doc = acad.ActiveDocument  
    ms = doc.ModelSpace  
    files = diam
    ht=h
    x=P1[0]
    y=P1[1]
    z=P1[2]
    pt1= POINT(x,y+ht,z)
    #InsertBlock(punto de insercion vector 3 doble , bloque string, escalax doble,escalay doble,escalaz doble, angulo )
    ms.InsertBlock(pt1, files, 0.01,0.01,0.01, 0)
    #print(reconoce(" 22 %%C 16 c/25-455"))
    
    try:
        doc.SelectionSets.Item("SS2").Delete()
    except:
        print("Delete selection failed")
    
    ssget1 = doc.SelectionSets.Add("SS2")
    ssget1.Select(4)
    #print(ssget1)
    #slt=doc.Select(0) # acSelectionSetAll = 5
    obj = ssget1[0]
    print(obj.layer)

    for entity in ssget1: # buscamos bloques con atributos
        name = entity.EntityName
        if name == 'AcDbBlockReference':
            HasAttributes = entity.HasAttributes
        if HasAttributes:
            for attrib in entity.GetAttributes():
                #item=attrib.TagString
                #cantidad=attrib.TextString
                attrib.TextString = cantidad #'455'
                attrib.Update()



def halla_valores():
    #seleccionar en pantalla()
    #Computa los bloques con atributos sumando cada bloque con atributo
    sumVHO=0
    sum6=0
    sum8=0
    sum10=0
    sum12=0
    sum16=0
    sum20=0
    sum25=0
    sum32=0
    doc = acad.ActiveDocument # Document object

    try:
        doc.SelectionSets.Item("SS1").Delete()
    except:
        a="Delete selection failed"
        #print(a)
    
    ssget1 = doc.SelectionSets.Add("SS1")
    ssget1.SelectOnScreen()
    #print(ssget1)

    for entity in ssget1: # buscamos bloques con atributos
        name = entity.EntityName
        if name == 'AcDbBlockReference':
            HasAttributes = entity.HasAttributes
            if HasAttributes:
                for attrib in entity.GetAttributes():
                    item=attrib.TagString
                    cantidad=attrib.TextString
                    if item=="VHO":
                        sumVHO=sumVHO+float(cantidad)
                    elif item=="6":
                        sum6=sum6+float(cantidad)
                    elif item=="8":
                        sum8=sum8+float(cantidad)
                    elif item=="10":
                        sum10=sum10+float(cantidad)
                    elif item=="12":
                        sum12=sum12+float(cantidad)
                    elif item=="16":
                        sum16=sum16+float(cantidad)
                    elif item=="20":
                        sum20=sum20+float(cantidad)
                    elif item=="25":
                        sum25=sum25+float(cantidad)
                    elif item=="32":
                        sum32=sum32+float(cantidad)

    if sumVHO>0:
        print("Volumen Hormigon=", sumVHO/1000000)

    if sum6>0:
        print("Varillas de 6 mm=", sum6/100, "mts.")

    if sum8>0:
        print("Varillas de 8 mm=", sum8/100, "mts.")

    if sum10>0:
        print("Varillas de 10 mm=", sum10/100, "mts.")

    if sum12>0:
        print("Varillas de 12 mm=", sum12/100, "mts.")

    if sum16>0:
        print("Varillas de 16 mm=", sum16/100, "mts.")

    if sum20>0:
        print("Varillas de 20 mm=", sum20/100, "mts.")

    if sum25>0:
        print("Varillas de 25 mm=", sum25/100, "mts.")

    if sum32>0:
        print("Varillas de 32 mm=", sum32/100, "mts.")


def crea_atributos():
    #esto solo debe ejecutarse una vez, esto crea automaticamente los bloques con atributos
    acad = win32com.client.Dispatch("AutoCAD.Application")  
    doc = acad.ActiveDocument  
            
    insertionPnt=POINT(0,0,0)
    height = 10
    mode = 0
    value = "0"
    
    # Block de 4.2 mm    
    try:
        oldblk = doc.Blocks.Item ("4.2") #'
        print ('Block 4.2 existe')
    except:
        prompt = "4.2"
        tag = "4.2"
        color=1
        obj= doc.Blocks.Add(insertionPnt, prompt )
        obj.AddAttribute(height, mode, prompt, insertionPnt, tag, value)   
        obj.Item(0).Color = color
        # Block de 4.2 mm    
    
    # Block de 6 mm    
    try:
        oldblk = doc.Blocks.Item ("6") #
        print ('Block 6 existe')
    except:
        prompt = "6"
        tag = "6"
        color=2
        obj= doc.Blocks.Add(insertionPnt, prompt )
        obj.AddAttribute(height, mode, prompt, insertionPnt, tag, value)   
        obj.Item(0).Color = color
    # Block de 6 mm    

    # Block de 8 mm    
    try:
        oldblk = doc.Blocks.Item ("8") #
        print ('Block 8 existe')
    except:
        prompt = "8"
        tag = "8"
        color=3
        obj= doc.Blocks.Add(insertionPnt, prompt )
        obj.AddAttribute(height, mode, prompt, insertionPnt, tag, value)   
        obj.Item(0).Color = color
    # Block de 8 mm    

    # Block de 10 mm    
    try:
        oldblk = doc.Blocks.Item ("10") #'
        print ('Block 10 existe')
    except:
        prompt = "10"
        tag = "10"
        color=4
        obj= doc.Blocks.Add(insertionPnt, prompt )
        obj.AddAttribute(height, mode, prompt, insertionPnt, tag, value)   
        obj.Item(0).Color = color
    # Block de 10 mm    

    # Block de 12 mm    
    try:
        oldblk = doc.Blocks.Item ("12") #
        print ('Block 12 existe')
    except:
        prompt = "12"
        tag = "12"
        color=5
        obj= doc.Blocks.Add(insertionPnt, prompt )
        obj.AddAttribute(height, mode, prompt, insertionPnt, tag, value)   
        obj.Item(0).Color = color
    # Block de 12 mm    

    # Block de 16 mm    
    try:
        oldblk = doc.Blocks.Item ("16") #'
        print ('Block 16 existe')
    except:
        prompt = "16"
        tag = "16"
        color=6
        obj= doc.Blocks.Add(insertionPnt, prompt )
        obj.AddAttribute(height, mode, prompt, insertionPnt, tag, value)   
        obj.Item(0).Color = color
    # Block de 16 mm    

    # Block de 20 mm    
    try:
        oldblk = doc.Blocks.Item ("20") #'
        print ('Block 20 existe')
    except:
        prompt = "20"
        tag = "20"
        color=7
        obj= doc.Blocks.Add(insertionPnt, prompt )
        obj.AddAttribute(height, mode, prompt, insertionPnt, tag, value)   
        obj.Item(0).Color = color
    # Block de 20 mm    

    # Block de 25 mm    
    try:
        oldblk = doc.Blocks.Item ("25") #
        print ('Block 25 existe')
    except:
        prompt = "25"
        tag = "25"
        color=8
        obj= doc.Blocks.Add(insertionPnt, prompt )
        obj.AddAttribute(height, mode, prompt, insertionPnt, tag, value)   
        obj.Item(0).Color = color
    # Block de 25 mm    

    # Block de 32 mm 
    try:
        oldblk = doc.Blocks.Item ("32") #
        print ('Block 32 existe')
    except:   
        prompt = "32"
        tag = "32"
        color=9
        obj= doc.Blocks.Add(insertionPnt, prompt )
        obj.AddAttribute(height, mode, prompt, insertionPnt, tag, value)   
        obj.Item(0).Color = color
    # Block de 32 mm    

    # Block de VHO    
    try:
        oldblk = doc.Blocks.Item ("VHO") #
        print ('Block VHO existe')
    except:
        prompt = "VHO"
        tag = "VHO"
        color=32
        obj= doc.Blocks.Add(insertionPnt, prompt )
        obj.AddAttribute(height, mode, prompt, insertionPnt, tag, value)   
        obj.Item(0).Color = color
    # Block de VHO

def main():
    # crea_atributos()
    # computa_texto() #Busca textos y computa cada texto de acuerdo al formato de texto
    halla_valores() #imprime el computo de las varillas para tipo de armaduras
       
    
main()