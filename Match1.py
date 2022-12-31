
import win32com.client                                                  # For Application connection
import pythoncom 
from pyautocad import Autocad,APoint,aDouble,aInt,aShort



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

 
dataType = (1001, 1000, 1000, 1040, 1040, 1040, 1040, 1040, 1000, 1040, 1040, 1040, 1040, 1040, 1040,1000)                 # Define Xdata
data = ('Sewer', 'name1' , ' ',3 , 0.7 , .013 , 0 , 0 ,  '0' ,  2 ,  1,0,0,0, 0,' ')
dataTypemaa = (1001, 1040, 1040, 1040, 1040, 1040, 1040, 1040, 1040,1040, 1000)                 # Define Xdata
datamaa = ('MAA', 0 , 0, 0 , 0 , 0 , 0 , 0 ,  0 , 0,  ' ')
          
            
delss('ss1',doc)
mainsel1 = doc.SelectionSets.Add("ss1")
alor=vtpt(0,0,0) #Necesario para parte del filtro
FilterType=vtint((0,8))
FilterData=vtvariant(('Line','Sewer'))
mainsel1.Select(int(5),pt1,pt2,FilterType,FilterData)
sscount = mainsel1.Count +1


if sscount>0:
    elem= mainsel1[0]
    ssobj = [elem for _ in range(mainsel1.Count+1)]
    point=elem.startpoint
    xdatamaa =mainsel1[0].GetXData("MAA", dataTypemaa, datamaa)
        
    maa = [xdatamaa[1] for _ in range(mainsel1.Count+1)]
    ma=[['MAA'] for _ in range (mainsel1.Count+1)]
    point1=[point in range(sscount+1)]
    point2=[point in range(sscount+1)]
    print (datamaa)
#for elem in mainsel1:
     #startpoint=elem.startpoint
    #endpoint=elem.endpoint
    #lineobj=model.AddLine(vtpt(startpoint[0],startpoint[1],0),vtpt(endpoint[0],endpoint[1],0))
  
    #elem.erase()

i=0    #  print (Xdata)
for elem1 in mainsel1 :
    datamaa1=None
    dataTypemaa1 =None
    xdatamaa =elem1.GetXData("MAA", dataTypemaa, datamaa)
    dataTypemaa=xdatamaa[0]
    datamaa=xdatamaa[1]
    i +=1
    datamaa= 'MAA' , i ,0,0,0,0,0 ,*datamaa[7:]
    print(datamaa)
   
    dataTypemaa1 = vtint(dataTypemaa)                          # Converse dataType format
    datamaa1= vtvariant(datamaa)
    elem1.SetXData( dataTypemaa1, datamaa1)
   
   
    #print (datamaa)
   
    ssobj[i]=elem1
    point1=ssobj[i].startpoint
    point2=ssobj[i].endpoint
    
    maa[i]=datamaa
    ma[i].append (i)
    ma[i].append (0)
    
    print ('ma',ma[i])
print('ss=',sscount)


for i in range(1,sscount):
    point1=ssobj[i].startpoint
    point2=ssobj[i].endpoint
    delss('ss1',doc)
    mainsel1 = doc.SelectionSets.Add('ss1')
    FilterType=vtint((0,8))
    FilterData=vtvariant(('Line','sewer'))
    mainsel1.Select (int(1), vtpt(point1[0]+0.5,point1[1]+0.5,0), vtpt(point1[0]-0.5,point1[1]-0.5 ,0) ,FilterType,FilterData)    
    print('mainsel1=',mainsel1.count)
    for elem1 in mainsel1: 
        
        j=0
        xdatamaa =elem1.GetXData("MAA", dataTypemaa, datamaa)
        datamaa=xdatamaa[1]
        j=int(datamaa[1])
     
      
        if j != i:    
            if abs(ssobj[j].endpoint[0]- ssobj[i].startpoint[0])<0.1 and abs(ssobj[j].endpoint[1]- ssobj[i].startpoint[1])<0.1  :
                
                   
                    ma[i]=*ma[i],j
                    print(ma[i],i)
                    ssobj[i].color=4
                    
                    ma[j]='MAA',ma[j][1],i,*ma[j][3:]
                    
        #print(ma[i])
        
for i in range(1,sscount) :
    datamaa1=None
    xdatamaa =ssobj[i].GetXData("MAA", dataTypemaa, datamaa)
    datamaa=xdatamaa[1]
    datamaa= *ma[i][0:], *datamaa[len(ma[i]):]
    
    if (datamaa[7])>0:
        ssobj[i].color=(datamaa[7])
    else:
            ssobj[i].color=256
    dataTypemaa1 = vtint(dataTypemaa)                          # Converse dataType format
    datamaa1= vtvariant(datamaa)
    ssobj[i].SetXData( dataTypemaa1, datamaa1)
          
        
print("Match...")
           