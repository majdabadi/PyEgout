
import win32com.client                                                  # For Application connection
import pythoncom 



global maxbasex , maxbasey , minbasex , minbasey , acad


def vtpt(x, y, z=0):
     return win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, (x, y, z))

def vtobj(obj):
    return win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_DISPATCH, obj)

def vtFloat(lis):
    """ list converted to floating points"""
    return win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, lis)

def vtint(val):
    return win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_I2, val)

def vtvariant(var):
    return win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_VARIANT, var)


    
def diaxdata(dia, slope, obj):
            dataType = (1001, 1040, 1040, 1040, 1040, 1040, 1040, 1040, 1040, 1040, 1040)                 # Define Xdata
            data = ('Diameter',dia * 10 , slope ,3 , 0.7 , .013 , 0 , 0 ,  0 ,  2 ,  1)
            dataType = vtint(dataType)                          # Converse dataType format
            data = vtvariant(data)
            obj.SetXData(dataType, data)    

try:
            acad = win32com.client.Dispatch("AutoCAD.Application")  
            doc = acad.ActiveDocument #Document Object
            model = doc.ModelSpace
except:
            print ("AutoCAD isn't running!")


def delss(setss,doc):
        try:
            doc.SelectionSets.Item("SS1").Delete()
        except:
            a="Delete selection failed"

   
    
acad.application.Visible = True
AppActivate = acad.application.Caption
print(AppActivate)

doc.Layers.Add("Pview")
layerObj = doc.Layers.Add("EGOUT")
doc.Layers.Add("EDIT")
layerObj = doc.Layers.Add("sewer")
layerObj.Color = 1
layerObj = doc.Layers.Add("area")
layerObj.Color = 6
layerObj = doc.Layers.Add("area-pline")
lyerObj = doc.Layers.Add("Den-pline")
layerObj.Color = 5
layerObj = doc.Layers.Add("area-density")
layerObj.Color = 3
layerObj = doc.Layers.Add("density")
layerObj = doc.Layers.Add("Latral")
layerObj.Color = 5
layerObj = doc.Layers.Add("SHEET-index")
layerObj = doc.Layers.Add("Grid")
layerObj =doc.Layers.Add("POP-TXT")
layerObj.Color = 4
layerObj = doc.Layers.Add("Elevation")
layerObj.Color = 7
layerObj =doc.Layers.Add("GELV")
layerObj.Color = 2
layerObj =doc.Layers.Add("Pview")
doc.ActiveLayer =doc.Layers("sewer")
print('Current Layer=' , doc.ActiveLayer.Name)   

delss('ss1',doc)
ssget1 = doc.SelectionSets.Add("SS1")
   

SELECT_ALL = int(5) 

    
valor=vtpt(0,0,0) #Necesario para parte del filtro
FilterType=vtint((0,8))
FilterData=vtvariant(('Line','Sewer'))

    #selecionar de acuerdo al criterio de filtro

ssget1.Select(SELECT_ALL,valor,valor,FilterType,FilterData)
x1=0
y1=0
x2=0
y2=0
    #Contar entidades
icount = ssget1.Count 
if icount>0:
    

    x2=ssget1[0].startpoint[0]
    y2=ssget1[0].startpoint[1]
    x1=x2
    y1=y2

for elem1 in ssget1:
    point2 = elem1.startpoint
    
    if point2[0] > x1:
         x1 = point2[0]
    if point2[1] > y1:
        y1 = point2[1]
             #print ('x1=',x1)
    if point2[0] < x2:
        x2 = point2[0]
    if point2[1] < y2:
        y2 = point2[1]
        
    point2 = elem1.endpoint
    if point2[1] > y1:
       y1 = point2[1]
    if point2[0] > x1:
        x1 = point2[0]
    if point2[0] < x2:
        x2 = point2[0]
    if point2[1] < y2 :
        y2 = point2[1]
       #print ('base',(minbasex),minbasey)
    acad.ZoomWindow (vtpt(x2,y2,0),vtpt(x1,y1,0))
    acad.ZoomScaled (0.9, 1)
    print(ssget1.Count ,'line')


delss('ss1',doc)
ssget1 = doc.SelectionSets.Add("SS1")
pt1=vtpt(0,0,0) #Necesario para parte del filtro
FilterType=vtint((0,8,0))
FilterData=vtvariant(('Circle','Egout'))


    #selecionar de acuerdo al criterio de filtro
ssget1.Select(SELECT_ALL,pt1,pt1,FilterType,FilterData)
FilterData=vtvariant(('Text','Egout'))
ssget1.Select(SELECT_ALL,pt1,pt1,FilterType,FilterData)


    #Contar entidades
icount = ssget1.Count 

print(icount)

if icount == 0 :
        #minbasex=0
        minbasey=0
        print(x2,y2)
        txtobj = model.AddText( 'EGOUT' , vtpt(x1 + 100  , y1 + 100 ,0  ), 100 )
        txtobj.layer ='Egout'
        txtobj.Alignment = 4
        txtobj.TextAlignmentPoint = vtpt(x1 + 100  ,y1 + 100 ,0  )
        for i in range( 20 , 40 ,5):
            cirobj = model.AddCircle(vtpt(x1 + 100  , y1 + 100,0 ), i//2 )
            cirobj.layer = "EGOUT"
           
            diaxdata(i , 0.004 , cirobj)
        for i in range(50, 110 , 10):
            cirobj = model.AddCircle(vtpt(x1 + 100,y1 + 100,0), i//2 )
            cirobj.layer = "EGOUT"
            diaxdata(i, 0.001,cirobj)
           
            dataType = (1001, 1040, 1040, 1040, 1040, 1040, 1040, 1040, 1040, 1040, 1040, 1000, 1040)                   # Define Xdata
            data = ("Criteria",0 , 0 ,100, 100 , 0 , 0 , 0 ,  0 ,  1 ,  0, "N", 5)
            dataType = vtint(dataType)                          # Converse dataType format
            data = vtvariant(data)
            txtobj.SetXData(dataType, data) 
            dataType = (1001, 1040, 1040, 1040, 1040, 1040, 1040, 1040, 1040, 1040, 1040, 1040, 1040)                   # Define Xdata
            data = ( "Runoff",0 , 0 , 0, 0 , 0 , 180 , 0 ,  0 ,  1 , .5, 4, 1)
            dataType = vtint(dataType)                          # Converse dataType format
            data = vtvariant(data)
            txtobj.SetXData(dataType, data)            
            dataType = (1001, 1040, 1040, 1040, 1040, 1040, 1040, 1040, 1040, 1040, 1040, 1040)  # Define Xdata
            data = ( "Infiltration",0 , 0 , 0, 0 , 0 , 0, 0 , 0, 0 , 0 , 0)
            dataType = vtint(dataType)                          # Converse dataType format
            data = vtvariant(data)
            txtobj.SetXData(dataType, data)     
            dataType = (1001, 1040, 1040, 1040, 1040, 1040, 1040,1040, 1040, 1040, 1040)           # Define Xdata
            data = ( "Population",0 , 0 , 0, 0 , 0 , 0, 0, 0 , 0 , 0)
            dataType = vtint(dataType)                          # Converse dataType format
            data = vtvariant(data)
            
            txtobj.SetXData(dataType, data) 
            dataType = (1001, 1000, 1000, 1000, 1000, 1000, 1000, 1000,1040, 1040, 1040, 1040)           # Define Xdata
            data = ( "Title", ' 1' , '2 ' , '3 ', '4 ' , ' 5' , ' 6' , '7', 0 , 0 , 0 , 0)
            dataType = vtint(dataType)                          # Converse dataType format
            data = vtvariant(data)
            txtobj.SetXData(dataType, data)   
             
            dataType = (1001, 1040, 1040, 1040, 1040, 1040, 1040, 1040)           # Define Xdata
            data = ( "Sheetindex",0 , 0 , 0, 0 , 0 , 0, 0)
            dataType = vtint(dataType)                          # Converse dataType format
            data = vtvariant(data)
            txtobj.SetXData(dataType, data)   
else:
   for el1 in ssget1:
        
        if el1.entityname=="AcDbText":
            print(el1.entityname)
            el1.TextAlignmentPoint = vtpt(x1 + 100  ,y1 + 100 ,0  )
        if el1.entityname=="AcDbCircle":
            el1.center=vtpt(x1+100,y1+100,0)
            
acad.ZoomExtents            
              
print ('Acad load')

