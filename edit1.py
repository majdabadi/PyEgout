
import win32com.client                                                  # For Application connection
import pythoncom 


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


try:
            acad = win32com.client.Dispatch("AutoCAD.Application")  
            doc = acad.ActiveDocument #Document Object
            model = doc.ModelSpace
except:
            print ("AutoCAD isn't running!")

def delss(setss,doc):
        try:
            doc.SelectionSets.Item(setss).Delete()
        except:
            a="Delete selection failed"

   
    #selecionar de acuerdo al criterio de filtro

acad.ZoomExtents
delss('ss1',doc)
ssget1 = doc.SelectionSets.Add("ss1")

SELECT_ALL = int(5) # 5=variable de autocad para selecionar todo

    #selecciona entidad circulo, Codigo dxf=0, Filterdata="Circle"
#selecciona radio del circulo, Codigo dxf=40, Filerdata=10, Todos los cirulos con radio =10
    #selecciona layer, Codigo dxf=8, Filterdata="uno", y que estan en el layer uno
    
pt1=vtpt(0,0,0)
pt2=vtpt(0,0,0)#Necesario para parte del filtro
FilterType=vtint((0,8))
FilterData=vtvariant(('Text','Sewer'))

    #selecionar de acuerdo al criterio de filtro

ssget1.Select(SELECT_ALL,pt1,pt2,FilterType,FilterData)

    #Contar entidades
icount = ssget1.Count 
print(icount)
for iobj in ssget1:
    iobj.Erase()    

           
            
delss('ss1',doc)
mainsel1 = doc.SelectionSets.Add("ss1")
alor=vtpt(0,0,0) #Necesario para parte del filtro
FilterType=vtint((0,8))
FilterData=vtvariant(('Line','Sewer'))
mainsel1.Select(int(5),pt1,pt2,FilterType,FilterData)
sscount = mainsel1.Count 
print(sscount)

if sscount>0:
    elem= mainsel1[0]
    ssobj = [elem for c in range(mainsel1.Count+1)]

#for elem in mainsel1:
     #startpoint=elem.startpoint
    #endpoint=elem.endpoint
    #lineobj=model.AddLine(vtpt(startpoint[0],startpoint[1],0),vtpt(endpoint[0],endpoint[1],0))
  
    #elem.erase()




dataType = (1001, 1000, 1000, 1040, 1040, 1040, 1040, 1040, 1000, 1040, 1040, 1040, 1040, 1040,1040, 1000)
data = ('Sewer', 'name1' , ' ',3 , 0.7 , .013 , 0 , 0 ,  '0' ,  2 ,  1,0,0,0, 0,' ')
dataTypemaa = (1001, 1040, 1040, 1040, 1040, 1040, 1040, 1040, 1040,1040,  1000)                 # Define Xdata
datamaa = ('MAA', 0 , 0, 0 , 0 , 0 , 0 , 0 ,  0  , 0,  ' ')

i=0    #  print (Xdata)
for elem1 in mainsel1:
    xdata =elem1.GetXData("sewer", dataType, data)
    data=xdata[1]
    i += 1
    ssobj[i]=elem1
    if i>0:      
        if data== None:    
            data = ('Sewer',  '*' + str(i) , ' ',3 , 0.7 , .013 , 0 , 0 ,  '0' ,  2 ,  1,0,0,0,0, ' ')
            dataType1 = vtint(dataType)                          # Converse dataType format
            data1 = vtvariant(data)
            elem1.SetXData(dataType1, data1)    
            datamaa = ('MAA', i , 0, 0 , 0 , 0 , 0 , 0 ,  0  , 0,  ' ')
            dataTypemaa1 = vtint(dataTypemaa)                          # Converse dataType format
            datamaa1 = vtvariant(datamaa)
            elem1.SetXData( dataTypemaa1, datamaa1)
        else: 
            dataType1=None
            data1=None
            xdata =elem1.GetXData("sewer", dataType, data)
            data=xdata[1]
            sdata=data[1]
            if sdata[0]=='*' and sdata[1] != '*':
                sdata='*' + sdata
                data=data[0],sdata, *data[2:]
                print(sdata,data)
                dataType1 = vtint(dataType)                          # Converse dataType format
                data1 = vtvariant(data)
                elem1.SetXData(dataType1, data1)  
                
            dataTypemaa1=None
            datamaa1=None
            xdatamaa =elem1.GetXData("MAA", dataTypemaa, datamaa)
            datamaa=xdatamaa[1]
            datamaa =  'MAA', i,*datamaa[2:8],0,0,' '
            print (datamaa)
            dataTypemaa1 = vtint(dataTypemaa)                          # Converse dataType format
            datamaa1 = vtvariant(datamaa)
            elem1.SetXData( dataTypemaa1, datamaa1)
      
print('ss=',sscount)


    
print ("Check Xdata")
for elem1 in mainsel1:
    if elem1.Length <= 2:
         elem1.Erase
    elif elem1.Length < 4:
        print("Error : Length <4")



ymax = mainsel1[0].startpoint

x2=ymax[0]
y2=ymax[1]
x1=x2
y1=y2

for  elem1 in mainsel1:
    #xdata= elem1.GetXData ("SEWER", dataType, data)

    #xmin=vtpt(0,0,0)
    #ymax=vtpt(0,100000,0)
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
xmin=vtpt(x2,y2,0)
ymax=vtpt(x1,y1,0)
print (xmin,ymax)
acad.ZoomWindow (xmin, ymax)
acad.ZoomScaled (0.9, 1)
print(mainsel1.Count ,'line')
if mainsel1.Count < 1:
    print( "Nothing Selected")



#pointsArray=[0.00 for c in range (6)] 
pointsArray=(0,0,0,0,0,0)
#
intPoints=(0,0,0)
for  elem in mainsel1:
    point1=elem.startpoint
    point2=elem.endpoint
    pointsArray =(point1[0]+.5,point1[1]+.5,point1[2],point2[0]-.5,point2[1]-.5,point2[2])
   
    pointsArray= win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, pointsArray )
    #print(pointsArray)
    
    delss('ss2',doc)
    mainsel2 = doc.SelectionSets.Add('ss2')
    FilterType=vtint((0,8))
    FilterData=vtvariant(('Line','sewer'))
    
    #mainsel2.Select(5, pt1, pt2, FilterType, FilterData)
    mainsel2.SelectByPolygon (2, pointsArray , FilterType, FilterData)
    
    icount = mainsel2.Count
    #print('ii=',icount)
    
    
    for  elem1 in mainsel2:
        intPoints =None
        intPoints = elem.IntersectWith(elem1, 1)
        if intPoints == point1 or intPoints == point2 :
            intPoints= None
        else:
            if len(intPoints)>0 :
                elem.color=2
        
delss('ss2',doc)
mainsel2 = doc.SelectionSets.Add('ss2')
FilterType=vtint((0,2))
FilterData=vtvariant(('Insert','Endblock'))
mainsel2.Select(int(5),pt1,pt2,FilterType,FilterData)    
icount = mainsel2.Count
print('endblock=',icount)  
if icount>0:
    elem= mainsel2[0]
    endblock = [elem for c in range(mainsel2.Count+1)]
    i=0
    for elem in mainsel2:
        i +=1
        endblock[i]=elem



sslist=[]
for elem in mainsel2:
   
    insertPoint = elem.insertionPoint
   # acad.ZoomWindow (vtpt(insertPoint[0]+50,insertPoint[1]+50,0), vtpt(insertPoint[0]-50,insertPoint[1]-50 ,0) )    
    print('mainsel11=',mainsel1.count)
   # acad.ZoomScaled (0.9, 1)
        
    print( "Endblock X= " , insertPoint[0], " Y=", insertPoint[1])
    delss('ss3',doc)
    mainsel1=None
    mainsel1 = doc.SelectionSets.Add('ss3')
    FilterType=vtint((0,8))
    FilterData=vtvariant(('Line','sewer'))
    
    pointsArray =(insertPoint[0]+.5,insertPoint[1]+.5,insertPoint[2],insertPoint[0]-.5,insertPoint[1]-.5,insertPoint[2])
   
    pointsArray= win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, pointsArray )
    
    mainsel1.Select (1, vtpt(insertPoint[0]+0.5,insertPoint[1]+0.5,0), vtpt(insertPoint[0]-0.5,insertPoint[1]-0.5 ,0) ,FilterType,FilterData)    
    print('mainsel1=',mainsel1.count)
    for varAttributes in elem.GetAttributes():
         if varAttributes.TagString == 'MhName':
            end1 = varAttributes.TextString 
            print (end1)

    for elem1 in mainsel1:
        
        data1=None
        dataType1= None
        datamaa1=None
        ataTypemaa1=None
        xdata= elem1.GetXData ("SEWER", dataType, data)
        
        data =xdata[1]
        dataType=xdata[0]
        data = data[0],data[1], end1 , *data[3:]
        print(data)
        dataType1 = vtint(dataType)                          # Converse dataType format
        data1= vtvariant(data)
        elem1.SetXData( dataType1, data1)
        xdatamaa=elem1.GetXData ('MAA', dataTypemaa, datamaa)
        dataTypemaa=xdatamaa[0]
        datamaa=xdatamaa[1]
        print('datamaaend=',datamaa)
        datamaa = *datamaa[0:9], elem1.length, datamaa[10]
        print( datamaa)
        sslist.append(datamaa[1])
        print(sslist,datamaa)
        
        dataTypemaa1 = vtint(dataTypemaa)                          # Converse dataType format
        datamaa1= vtvariant(datamaa)
        elem1.SetXData( dataTypemaa1, datamaa1)
        elem1.color=2
        if abs(elem1.endpoint[0]- insertPoint[0])<0.1 and abs(elem1.endpoint[1]- insertPoint[1])<0.1:
            print(elem1.endpoint)
        else:
            endpoint=elem1.endpoint
            print(vtpt(endpoint[0],endpoint[0],0))
            elem1.startpoint=vtpt(endpoint[0],endpoint[1],0)
            elem1.endpoint=vtpt( insertPoint[0], insertPoint[1],0)
           # print(elem1.endpoint , insertPoint)
print ('sslist=',sslist)
acad.ZoomExtents()
#acad.ZoomWindow (xmin, ymax)
#acad.ZoomScaled (0.9, 1)

print('Net Def')
print(sslist)
while len(sslist)>0:
        print('color ', ssobj [int(sslist[0])].color)
        print ('ss0',int(sslist[0]))
        point1=ssobj [int(sslist[0])].startpoint
        xdatamaa=ssobj [int(sslist[0])].GetXData('MAA', dataTypemaa, datamaa)
        datamaa=xdatamaa[1]
        totlen=datamaa[9]
        #print('tlen',totlen)
        #acad.ZoomWindow (vtpt(point1[0]+50,point1[1]+50,0), vtpt(point1[0]-50,point1[1]-50 ,0) )    
        #acad.ZoomScaled (0.5, 1)
        sslist.remove(sslist[0])
        delss('ss1', doc)
        mainsel1 = doc.SelectionSets.Add('ss1')
        FilterType=vtint((0,8))
        FilterData=vtvariant(('Line','sewer'))
        mainsel1.Select (1, vtpt(point1[0]+0.1,point1[1]+0.1,0), vtpt(point1[0]-0.1,point1[1]-0.1 ,0) ,FilterType,FilterData)    
        print('mainsel1=',mainsel1.count)
        for elem1 in mainsel1:
            datamaa1=None
            dataTypemaa1=None
            xdatamaa=elem1.GetXData ('MAA', dataTypemaa, datamaa)
            datamaa=xdatamaa[1]
            print('d9',datamaa[9])
            if datamaa[9]==0:
                #print('datamaa=',datamaa)
                datamaa = *datamaa[0:9], totlen + elem1.length, datamaa[10]
                sslist.append(datamaa[1])
                #print(datamaa)
                dataTypemaa1 = vtint(dataTypemaa)                          # Converse dataType format
                datamaa1= vtvariant(datamaa)
                elem1.SetXData( dataTypemaa1, datamaa1)
                elem1.color=6
                if abs(elem1.endpoint[0]- point1[0])<0.1 and abs(elem1.endpoint[1]- point1[1])<0.1 :
                    elem1.endpoint
                    elem1.color=3
                else:
                    endpoint=elem1.endpoint
                    #print(vtpt(endpoint[0],endpoint[0],0))
                    elem1.startpoint=vtpt(endpoint[0],endpoint[1],0)
                    elem1.endpoint=vtpt( point1[0], point1[1],0)
                    elem1.color=4
            print('list',sslist)