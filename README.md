# python_autocad
Proyectos para automatizacion de cad en Python
requiere pywin32
python -m pip install pywin32
*****************************************************************************************************
computo_armaduras.py
*****************************************************************************************************
Este programa crea bloques con atributos para computar varillas
se crean bloques llamados:

"4.2" para varillas de 4.2 mm

"6" para varillas de 6 mm

"8" para varillas de 8 mm

"10" para varillas de 10 mm

"12" para varillas de 12 mm

"16" para varillas de 16 mm

"20" para varillas de 20 mm

"25" para varillas de 25 mm

"32" para varillas de 32 mm

El programa reconoce los textos en formato:

"2 %%C 20 - 600"

"3 Ø 10 - 1018"

"22 fi de 16 c/24-455"

"Pos44_2%%c16c/24(455)"

Computa las armaduras e inserta un bloque cpn atributos para los diferentes tipos de armaduras.
Finalmente puede computar la suma de todas la armaduras sumando cada tipo de bloque, leyendo los atributos de cada uno

*****************************************************************************************************
AddExtrudedSolid
*****************************************************************************************************
Ejemplo  como añadir una region y crear un sólido a partir de esa region
