#Import needed modules
import os
import win32com.client
import pythoncom
import random
import math

#from pythoncom import  VT_DISPATCH, VT_ARRAY, VT_UI2,VT_VARIANT,VT_BSTR
def Point(x, y, z):
    """Coordinate points are converted to floating point numbers""" 
    return win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, (x, y, z))

def vtInt(list):
    """list is converted to integer """ 
    return win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_I2, list)

def aArray(vObject):
    return win32com.client.VARIANT(pythoncom.VT_VARIANT | pythoncom.VT_ARRAY,(vObject))


def main():
    #Get running AutoCAD instance
    try:
        acad = win32com.client.Dispatch("AutoCAD.Application")  

    except:
        print ("AutoCAD isn't running!")

    doc = acad.ActiveDocument #Document Object

    try:
        doc.SelectionSets.Item("SS1").Delete()
    except:
        a="Delete selection failed"
    
    ssget1 = doc.SelectionSets.Add("SS1")

    
    SELECT_ALL = int(5) # 5=variable de autocad para selecionar todo

    #selecciona entidad circulo, Codigo dxf=0, Filterdata="Circle"
    #selecciona radio del circulo, Codigo dxf=40, Filerdata=10, Todos los cirulos con radio =10
    #selecciona layer, Codigo dxf=8, Filterdata="uno", y que estan en el layer uno
    
    valor=Point(0,0,0) #Necesario para parte del filtro
    FilterType=vtInt((0,40,8))
    FilterData=aArray(('Circle',10,'uno'))

    #selecionar de acuerdo al criterio de filtro
    ssget1.Select(SELECT_ALL,valor,valor,FilterType,FilterData)

    #Contar entidades
    icount = ssget1.Count

    print(icount)

    #Moverl al layer movido
    for entity in ssget1:
        entity.Layer="movido"

def dimension():
    #Get running AutoCAD instance
    try:
        acad = win32com.client.Dispatch("AutoCAD.Application")  

    except:
        print ("AutoCAD isn't running!")

    doc = acad.ActiveDocument #Document Object

    try:
        doc.SelectionSets.Item("SS1").Delete()
    except:
        a="Delete selection failed"
    
    ssget1 = doc.SelectionSets.Add("SS1")

    
    SELECT_ALL = int(5) # 5=variable de autocad para selecionar todo

    #selecciona entidad circulo, Codigo dxf=0, Filterdata="Circle"
    #selecciona radio del circulo, Codigo dxf=40, Filerdata=10, Todos los cirulos con radio =10
    #selecciona layer, Codigo dxf=8, Filterdata="uno", y que estan en el layer uno
    
    valor=Point(0,0,0) #Necesario para parte del filtro
    FilterType=vtInt((0,3))
    FilterData=aArray(('Dimension','Standard'))

    #selecionar de acuerdo al criterio de filtro
    ssget1.Select(SELECT_ALL,valor,valor,FilterType,FilterData)

    #Contar entidades
    icount = ssget1.Count

    print(icount)

    #Moverl al layer movido
    for entity in ssget1:
        entity.StyleName="ISO-25"



def dim():
    #Dibujar circulos en forma aleatoria
    z=0
    try:
        acad = win32com.client.Dispatch("AutoCAD.Application")  
    except:
        print ("AutoCAD isn't running!")
    
    doc = acad.ActiveDocument #Document Object
    
    estilos=['Standard','ISO-25','cota50']
    layer=['uno','dos','tres','cuatro']

    for i in range (0, 100):
        x1=round(random.random()*100,2)
        y1=round(random.random()*100,2)
        x2=round(random.random()*100,2)
        y2=round(random.random()*100,2)
        
        lay=random.choice(layer)
        sty=random.choice(estilos)
     
        pdmx = (x1 + x2) / 2+0.5
        pdmy = (y1 + y2) / 2 + 0.5
        po1 = Point(x1,y1,0)
        po2 = Point(x2,y2,0)
        pod = Point(pdmx,pdmy,0)

        dm1 = doc.ModelSpace.AddDimAligned(po1, po2, pod)
        #dm1 = doc.ModelSpace.AddDimRotated(po1, po2, pod, 0)
        dm1.Layer=lay
        dm1.StyleName=sty





def circulos():
    #Dibujar circulos en forma aleatoria
    z=0
    try:
        acad = win32com.client.Dispatch("AutoCAD.Application")  
    except:
        print ("AutoCAD isn't running!")
    
    doc = acad.ActiveDocument #Document Object
    
    Layer=['uno','dos','tres','cuatro']

    for i in range (0, 1000):
        x=round(random.random()*100,2)
        y=round(random.random()*100,2)
        r=round(random.random()*20+1,0)
        lay=random.choice(Layer)
        P1=Point(x,y,z)
        circle=doc.ModelSpace.AddCircle(P1,r)
        circle.Layer=Layer[lay]












"""
    if icount != 0:
        #avoid the combination between items' incrementation and deletion
        #|
        #|
        AllElem = [ssBlocks.Item(i) for i in range(0,icount)]
        #|
        #|
        #_____
        ssBlocks.Clear

        for i in range(0,icount):
            AllElem[i].Delete
        AllElem = []
"""

#Execution Part
if __name__ == '__main__':
    #main()
    #circulos()
    #dimension()
    dim()


""" Codigos Para el Filtro:
0,Text string indicating the entity type (fixed)
1,Primary text value for an entity
2,Name (attribute tag, block name, and so on)
3-4,Other text or name values
5,Entity handle; text string of up to 16 hexadecimal digits (fixed)
6,Linetype name (fixed)
7,Text style name (fixed)
8,Layer name (fixed)
9,DXF: variable name identifier (used only in HEADER section of the DXF file)

10,Primary point; this is the start point of a line or text entity, center of a circle, and so on DXF: X value of the primary point (followed by Y and Z value codes 20 and 30). APP: 3D point (list of three reals)

11-18,Other points. DXF: X value of other points (followed by Y value codes 21-28 and Z value codes 31-38) APP: 3D point list of three reals)

20, 30,DXF: Y and Z values of the primary point

21-28,
31-37, DXF: Y and Z values of other points

38,DXF: entity's elevation if nonzero
39, Entity's thickness if nonzero (fixed)

40-48,Double-precision floating-point values (text height, scale factors, and so on)

48, Linetype scale; double precision floating point scalar value; default value is defined for all entity types

49,Repeated double-precision floating-point value. Multiple 49 groups may appear in one entity for variable-length tables (such as the dash lengths in the LTYPE table). A 7x group always appears before the first 49 group to specify the table length

50-58, Angles (output in degrees to DXF files and radians through AutoLISP and ObjectARX applications)
60, Entity visibility; integer value; absence or 0 indicates visibility; 1indicates invisibility
62, Color number (fixed)
66, “Entities follow” flag (fixed)
67, Space—that is, model or paper space (fixed)
68, APP: identifies whether viewport is on but fully off screen; is not active or is off
69, APP: viewport identification number
70-78, Integer values, such as repeat counts, flag bits, or modes
90-99, 32-bit integer values
100, Subclass data marker (with derived class name as a string). Required for all objects and entity classes that are derived from another concrete class. The subclass data marker segregates data defined by different classes in the inheritance chain for the same object.
This is in addition to the requirement for DXF names for each distinct concrete class derived from ObjectARX (see Subclass Markers)

102, Control string, followed by “{<arbitrary name>” or “}”. Similar to the xdata 1002 group code, except that when the string begins with “{“, it can be followed by an arbitrary string whose interpretation is up to the application. The only other control string allowed is “}” as a group terminator. AutoCAD does not interpret these strings except during drawing audit operations. They are for application use

105, Object handle for DIMVAR symbol table entry
110, UCS origin (appears only if code 72 is set to 1) DXF: X value; APP: 3D point
111, UCS X-axis (appears only if code 72 is set to 1) DXF: X value; APP: 3D vector
112, UCS Y-axis (appears only if code 72 is set to 1) DXF: X value; APP: 3D vector
120-122, DXF: Y value of UCS origin, UCS X-axis, and UCS Y-axis
130-132, DXF: Z value of UCS origin, UCS X-axis, and UCS Y-axis
140-149, Double-precision floating-point values (points, elevation, and DIMSTYLE settings, for example)
170-179, 16-bit integer values, such as flag bits representing DIMSTYLE settings
210, Extrusion direction (fixed) DXF: X value of extrusion direction APP: 3D extrusion direction vector
220, 230, DXF: Y and Z values of the extrusion direction
270-279, 16-bit integer values
280-289, 16-bit integer value
290-299, Boolean flag value
300-309, Arbitrary text strings
310-319, Arbitrary binary chunks with same representation and limits as 1004 group codes: hexadecimal strings of up to 254 characters represent data chunks of up to 127 bytes

320-329, Arbitrary object handles; handle values that are taken “as is”. They are not translated during INSERT and XREF operations

330-339, Soft-pointer handle; arbitrary soft pointers to other objects within same DXF file or drawing. Translated during INSERT and XREF operations

340-349, Hard-pointer handle; arbitrary hard pointers to other objects within same DXF file or drawing. Translated during INSERT and XREF operations

350-359, Soft-owner handle; arbitrary soft ownership links to other objects within same DXF file or drawing. Translated during INSERT and XREF operations

360-369, Hard-owner handle; arbitrary hard ownership links to other objects within same DXF file or drawing. Translated during INSERT and XREF operations

370-379, Lineweight enum value (AcDb::LineWeight). Stored and moved around as a 16-bit integer. Custom non-entity objects may use the full range, but entity classes only use 371-379 DXF group codes in their representation, because AutoCAD and AutoLISP both always assume a 370 group code is the entity's lineweight. This allows 370 to behave like other “common” entity fields

380-389, PlotStyleName type enum (AcDb::PlotStyleNameType). Stored and moved around as a 16-bit integer. Custom non-entity objects may use the full range, but entity classes only use 381-389 DXF group codes in their representation, for the same reason as the Lineweight range above

390-399, String representing handle value of the PlotStyleName object, basically a hard pointer, but has a different range to make backward compatibility easier to deal with. Stored and moved around as an object ID (a handle in DXF files) and a special type in AutoLISP. Custom non-entity objects may use the full range, but entity classes only use 391-399 DXF group codes in their representation, for the same reason as the lineweight range above

400-409, 16-bit integers
410-419, String

420-427,32-bit integer value. When used with True Color; a 32-bit integer representing a 24-bit color value. The high-order byte (8 bits) is 0, the low-order byte an unsigned char holding the Blue value (0-255), then the Green value, and the next-to-high order byte is the Red Value. Convering this integer value to hexadecimal yields the following bit mask: 0x00RRGGBB. For example, a true color with Red==200, Green==100 and Blue==50 is 0x00C86432, and in DXF, in decimal, 13132850

430-437, String; when used for True Color, a string representing the name of the color
440-447, 32-bit integer value. When used for True Color, the transparency value
450-459,Long
460-469, Double-precision floating-point value
470-479, String

480-481, Hard-pointer handle; arbitrary hard pointers to other objects within same DXF file or drawing. Translated during INSERT and XREF operations

999, DXF: The 999 group code indicates that the line following it is a comment string. SAVEAS does not include such groups in a DXF output file, but OPEN honors them and ignores the comments. You can use the 999 group to include comments in a DXF file that you've edited

1000, ASCII string (up to 255 bytes long) in extended data
1001, Registered application name (ASCII string up to 31 bytes long) for extended data
1002, Extended data control string (“{” or “}”)
1003, Extended data layer name
1004, Chunk of bytes (up to 127 bytes long) in extended data
1005, Entity handle in extended data; text string of up to 16 hexadecimal digits
1010, A point in extended data.DXF: X value (followed by 1020 and 1030 groups).APP: 3D point
1020, 1030,DXF: Y and Z values of a point
1011,A 3D world space position in extended data. DXF: X value (followed by 1021 and 1031 groups).APP: 3D point
1021, 1031,DXF: Y and Z values of a world space position
1012, A 3D world space displacement in extended data.DXF: X value (followed by 1022 and 1032 groups).APP: 3D vector
1022, 1032, DXF: Y and Z values of a world space displacement
1013, A 3D world space direction in extended data. DXF: X value (followed by 1022 and 1032 groups).APP: 3D vector
1023, 1033, DXF: Y and Z values of a world space direction
1040, Extended data double-precision floating-point value
1041, Extended data distance value
1042, Extended data scale factor
1070, Extended data 16-bit signed integer
1071, Extended data 32-bit signed long 

"""