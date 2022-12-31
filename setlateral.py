import win32com.client                                                  # For Application connection
import pythoncom 
from pyautocad import Autocad,APoint,aDouble,aInt,aShort

import random as rand
import time
import os
import sys
import glob


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
ssget1 = doc.SelectionSets.Add('ss1')

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
dataTypemaa = (1001, 1040, 1040, 1040, 1040, 1040, 1040, 1040, 1040, 1040, 1000)                 # Define Xdata
datamaa = ('MAA', 0 , 0, 0 , 0 , 0 , 0 , 0 ,  0 ,  0 ,   ' ')
          
            
delss('ss1',doc)
mainsel1 = doc.SelectionSets.Add("ss1")
alor=vtpt(0,0,0) #Necesario para parte del filtro
FilterType=vtint((0,8))
FilterData=vtvariant(('Line','Sewer'))
mainsel1.Select(int(5),pt1,pt2,FilterType,FilterData)
sscount = mainsel1.Count +1
print('ss=',sscount)

if sscount>1:
    elem= mainsel1[0]
    ssobj = [elem for c in range(mainsel1.Count+1)]
    xdatamaa =mainsel1[0].GetXData("MAA", dataTypemaa, datamaa)
    maa = [xdatamaa[1] for c in range(mainsel1.Count+1)]

#for elem in mainsel1:
     #startpoint=elem.startpoint
    #endpoint=elem.endpoint
    #lineobj=model.AddLine(vtpt(startpoint[0],startpoint[1],0),vtpt(endpoint[0],endpoint[1],0))
  
    #elem.erase()
totlenmax=0 
jmax=0
ma5=[]
ma4=[]  
i=0    #  print (Xdata)
for elem1 in mainsel1:
    datamaa1=None
    dataTypemaa1=None
    xdatamaa =elem1.GetXData("MAA", dataTypemaa, datamaa)
    dataTypemaa=xdatamaa[0]
    datamaa=xdatamaa[1]
    ssobj[int(datamaa[1])]=elem1
    maa[int(datamaa[1])]='MAA',datamaa[1],abs(datamaa[2]),* datamaa[3:8],0,datamaa[9] ,' '
    
    if totlenmax<datamaa[9]:
         totlenmax=datamaa[9]
         jmax=int(datamaa[1])
         
    j1=0
    j2=0
   
    j1=int(datamaa[3])
    j2=int(datamaa[4])
    
    if j1==0:
             
        ma5.append(int(datamaa[1]))
    if j2>0:
        
        ma4.append(int(datamaa[1]))
    
    
    maa[0]='MAA',0,0,0,0,0,0,0,0,0,' '
    
    #datamaa = *datamaa[0:9], 0, ' '
    #datamaa='MAA',datamaa[1],abs(datamaa[2]),* datamaa[3:]
    #dataTypemaa1 = vtint(dataTypemaa)                          # Converse dataType format
    #datamaa1= vtvariant(datamaa)
    #elem1.SetXData( dataTypemaa1, datamaa1)
    #elem1.color=6
    




for i in range(1, sscount):
   
    xdatamaa =ssobj[i].GetXData("MAA", dataTypemaa, datamaa)
    datamaa=xdatamaa[1]
    
    if totlenmax<datamaa[9]:
         totlenmax=datamaa[9]
         jmax=int(datamaa[1])
    



 


 # Total Length

for  jj in ma5:
    #print(jj)
    totlen=0
    while jj>0   :        
        totlen=totlen + ssobj[jj].length
        totlen8=maa[jj][8]
        #print('totlen',totlen,'ma8' ,totlen8)
        if totlen >= totlen8:
            maa[jj]=*maa[jj][0:8] , totlen , maa[jj][9] , ' '
            #print('totlen8',maa[jj])
        jj=abs(int(maa[jj][2]))
        if  maa[jj][8]>totlen :
            print(totlen,maa[jj][8])
            jj=0
 
 
for jj in ma4:
    j1=0
    j2=0
    j3=0
    print('j3',maa[jj][2],maa[jj][4],maa[jj][5])
    j1=int(maa[jj][3])
    j2=int(maa[jj][4])
    j3=int(maa[jj][5]) 
    
    if j2>0 and maa[j2][8]> maa[j1][8] and maa[j2][8]> maa[j3][8]:
         
         maa[jj]='MAA',jj,maa[jj][2],j2,j1,j3,*maa[jj][6:]
        
    if j3>0 and maa[j3][8]> maa[j1][8] and maa[j3][8]> maa[j2][8]:
        maa[jj]='MAA',jj,maa[jj][2],j3,j1,j2,*maa[jj][6:]
                       
for i in range(1,sscount) :
        j1=0
        j2=0
        j3=0
        j1=int(maa[i][3])
        j2=int(maa[i][4])
        j3=int(maa[i][5]) 
        #if j1>0:
             # maa[j1]='MAA',j1,i,*maa[j1][3:]
             #ssobj[j1].color=7
        if j2>0:
          maa[j2]='MAA',j2,-i,*maa[j2][3:]
          
        if j3>0:
            maa[j3]='MAA',j3,-i,*maa[j3][3:] 
            

#ssobj[jmax].color=2  

ii=0
for  jj in ma5:
    ii +=1
    if ii>7:
        ii=1
   
   
    while jj>0  :
        
        print(jj)
        ssobj[int(jj)].color=ii
        jj=maa[int(jj)][2]
        
        
ma=[]
jj=jmax
print('jmax',jmax)
ssobj[jmax].color=7
point1=ssobj[jmax].startpoint
model.addtext('AA'+ maa[jmax][10],vtpt(point1[0],point1[1],0),50)
print ('jmax',jmax,maa[jmax][9])


ii=0   
ll=0 
while jj >0 and ii<sscount :
    ii +=1
    
    maa[jj]= *maa[jj][0:10] ,'A'
    print(maa[jj])
    ssobj[jj].color=7
    j1=int(maa[jj][3])
    print(maa[j1])
    if j1>0:
        maa[j1]= *maa[j1][0:10] ,'A'
        #ssobj[j1].color=51
        
    j2=int(maa[jj][4])
    if j2>0:
        #ssobj[j2].color=21
        ll=ll+1
        maa[j2]= *maa[j2][0:10] ,'A'+  str(ll)
        ma.append(j2)
    j3=int(maa[jj][5])
    if j3>3:
        #ssobj[j3].color=43
        ll=ll+1
        maa[j3]= *maa[j3][0:10] ,'A'+ str(ll)
        ma.append(j3)
    j4=int(maa[jj][6])
    jj=abs(int(maa[jj][2]))
      
print (ma)
while len(ma)>0 :
    jj=ma[0] 
    ma.remove(ma[0]) 
    latname=maa[jj][10]
    print('lat',latname)
    ll=0
    while jj > 0 :
        maa[jj]=* maa[jj][0:10],latname
        #ssobj[jj].color=121    
        j2=int(maa[jj][4])
        print (maa[j2][10])
        if maa[j2][10]==' ':
            #ssobj[j2].color=171
            ll=ll+1
            maa[j2]= *maa[j2][0:10] ,latname + "-" + str(ll)
            ma.append(j2)
        j3=int(maa[jj][5])
        if maa[j3][10]==' ':
            #ssobj[jj].color=151
            ll=ll+1
            maa[j3]= *maa[j3][0:10] ,latname + "-" + str(ll)
            ma.append(j2)
        j4=int(maa[jj][6])
        jj=int(maa[jj][3])
for i in range(1,sscount) :
    point1=ssobj[i].startpoint
    model.addtext(maa[i][10],vtpt(point1[0],point1[1],0),5)
    print(maa[i])
#--------------------------------------------------------------------
dataTypemaa = (1001, 1040, 1040, 1040, 1040, 1040, 1040, 1040, 1040, 1040, 1000)

for i in range(1,sscount) :
       
    
    datamaa1=None
    dataTypemaa1==None
    xdatamaa =ssobj[i].GetXData("MAA", dataTypemaa, datamaa)
    #dataTypemaa=xdatamaa[0]
    datamaa=xdatamaa[1]
    print('ma=',maa[i])
    
    datamaa=maa[i]
    dataTypemaa1 = vtint(dataTypemaa)                          # Converse dataType format
    datamaa1= vtvariant(datamaa)
    ssobj[i].SetXData( dataTypemaa1, datamaa1)