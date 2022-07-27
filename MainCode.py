import os
import struct
import numpy
import pandas as pd
from sklearn.tree import DecisionTreeClassifier
from sklearn.neighbors import KNeighborsClassifier
from sklearn.tree import export_graphviz
import seaborn as sns
import time
from datetime import *
from read_comtrade import ReadAllComtrade 
from scipy.ndimage import gaussian_filter1d
from Functions import *

"""
from Auto_TestBed_Hypersim import *

print( 'Test time') 
print( Test_Time )

#time.sleep(120)                                                                                # Tiempo para descargar los Comtrades la carpeta especificada
input('Donwload the Comtrade Files into the specific path and Press Enter to continue')         # Tiempo para descargar los Comtrades la carpeta especificada y presionar 'Enter' para continuar la ejecución del código
    
print('Next...')
"""



   

"""
Main code start here
This code is for the relay testing protection 

There are two algorithms for the classification 

Decision Tree and k-Nearest Neighbors

"""


Test_Time= {'Caso_01': datetime(2022, 6, 14, 10, 17, 52, 183476), 'Caso_02': datetime(2022, 6, 14, 10, 18, 38, 641417), 
            'Caso_03': datetime(2022, 6, 14, 10, 19, 25, 424889), 'Caso_06': datetime(2022, 6, 14, 10, 20, 11, 829064), 
            'Caso_07': datetime(2022, 6, 14, 10, 20, 58, 175197), 'Caso_08': datetime(2022, 6, 14, 10, 21, 44, 815129), 
            'Caso_09': datetime(2022, 6, 14, 10, 22, 31, 586886), 'Caso_10': datetime(2022, 6, 14, 10, 23, 18, 238828), 
            'Caso_11': datetime(2022, 6, 14, 10, 24, 5, 122402), 'Caso_12': datetime(2022, 6, 14, 10, 24, 51, 680568), 
            'Caso_13': datetime(2022, 6, 14, 10, 25, 38, 67545), 'Caso_14': datetime(2022, 6, 14, 10, 26, 24, 496119), 
            'Caso_15': datetime(2022, 6, 14, 10, 27, 11, 27462), 'Caso_16': datetime(2022, 6, 14, 10, 27, 57, 563659), 
            'Caso_04': datetime(2022, 6, 14, 10, 28, 44, 50639), 'Caso_05': datetime(2022, 6, 14, 10, 29, 30, 251137)}





Carpeta= 'Ensayo Full5'
pathh = 'D:\JoseM\Relé Siemens - Cama de Pruebas\Pruebas Fallas CID\Ensayos Automatismo Python' 
path2= pathh + "\\" + Carpeta
path = os.path.realpath(__file__)
dir_path = os.path.dirname( os.path.realpath(__file__) )

allfiles = []
failFiles = []
allnames=[]
N_dict= {}

for root, dirs, files in os.walk( path2, topdown= False ):
#for root, dirs, files in os.walk( dir_path, topdown= False ):
   for name in files:
       if name.find('.cfg') > -1 or name.find('.CFG') > -1:
           allnames.append(name[:-4])                                           # Nombre del archivo Comtrade sin la extensión '.cfg'
           allfiles.append( os.path.join(root, name) )
           
           pos1= name.find('_')                                                 # Estos nombres por defecto tienen dos guiones, acá se halla la posición del primero
           date_com= name[ (pos1 + 1) : name.find('_', (pos1 + 1) ) ]           # Se saca la fecha y hora que está entre los dos guiones
           
           N_dict[ os.path.join(root, name) ]= name[:-4]                        # Se asocia el path de cada archivo con el nombre extraído
           



DO=[]
NN= {}

for file in list( N_dict.keys() ):
    ComtradeObjec= ReadAllComtrade( file )                          # Firts step, create instance for the comtrade class
    ComtradeObjec.ReadDataFile()                                    # Next, read the data file
    Time= ComtradeObjec.getTimeChannel()/1000000                     # Time in us
    
    A1, A2, A3, A4 = DefCaso( N_dict[ file ] )

    D1, T1= Digital( ComtradeObjec.getDigitalASCCI(1), Time )          # Trip a
    D2, T2= Digital( ComtradeObjec.getDigitalASCCI(2), Time )          # Trip b
    D3, T3= Digital( ComtradeObjec.getDigitalASCCI(3), Time )          # Trip c
    D4, T4= Digital( ComtradeObjec.getDigitalASCCI(4), Time )          # 21 Dir Forward
    D5, T5= Digital( ComtradeObjec.getDigitalASCCI(5), Time )          # 21 Dir Backward
    D6, T6= Digital( ComtradeObjec.getDigitalASCCI(6), Time )          # Trip Z4 reversal
    D7, T7= Digital( ComtradeObjec.getDigitalASCCI(7), Time )          # Trip Z2
    D8, T8= Digital( ComtradeObjec.getDigitalASCCI(8), Time )          # Trip Z1
    D9, T9= Digital( ComtradeObjec.getDigitalASCCI(9), Time )          # Trip 67N
    D10, T10= Digital( ComtradeObjec.getDigitalASCCI(10), Time )       # Trip O 85 - 67N  POTT
    D11, T11= Digital( ComtradeObjec.getDigitalASCCI(11), Time )      # Send O 85 - 67N  POTT
    D12, T12= Digital( ComtradeObjec.getDigitalASCCI(12), Time )      # Trip O 85 - 21  POTT
    D13, T13= Digital( ComtradeObjec.getDigitalASCCI(13), Time )      # Send O 85 - 21  POTT

    NN[ file ]= ComtradeObjec.start
    
    Digital_Out=[ D1, D2, D3, D4, D5, D6, D7, D8, D9, D10, D11, D12, D13, A1, A2, A3, A4]        # Construye el vector de 
#    Digital_Out=[ A1, A2, A3, A4, D1, D2, D3, D4, D5, D6, D7, D8, D9, D10, D11, D12, D13]        # Construye el vector de 
    DO.append(Digital_Out)
 
 


EvaluateObject= Evaluate()                                                              # Create an instance for Evaluation
Data= EvaluateObject.training()
Estimated= [ EvaluateObject.Result( DO[ii] ) for ii in range( len( DO ) ) ]             # Evaluate each DO 
cols_Re= list( Data.columns[17:-3] )                                                    # Catch the columns with the information for the results 
Re= [ Data[ cols_Re ].loc[ Data['Target']== ii ].values[0] for ii in Estimated ]        # Create a matrix with the ordered information results 
#Ree= [ numpy.append( numpy.array( DO[i] ), r ) for i, r in enumerate( Re ) ]

print(Estimated)


# Export Tree
#EvaluateObject.ExportTree( 'ElArbolito' )


Ress= Re[:] 
for ii,k in enumerate( Estimated ):                                                     # Se compara el vector de digitales que se captura de los comtrades con el vector de digitales que el DecisionTree asocia para el caso, con el fin de ver si el caso seleccionado sí es el correcto
    
    var= list( Data[ list( Data.columns[:17] ) ].loc[ Data['Target']== k ].values[0])
    
    if DO[ii] != var:        
        Ress[ii]= [ ' - ' for i in range( len( Re[ii] ) - 1 ) ] + [ 'No Asociado' ]
        
               

#Export excel file result
df = pd.DataFrame( Ress, columns= cols_Re )
folder= os.getcwd()
#File_exp= os.getcwd() + '\\' + 'Resultados.xlsx'
File_exp= path2 + '\\' + 'Resultados__' + Carpeta + '.xlsx'
export_excel = df.to_excel ( File_exp, index= None, header= True)


print( df[['Caso','Calificación']] )



"""

#plot matrix
Cal= df['Calificación'].unique()
#Res= numpy.concatenate( ( Res, ['OK'] ), axis= 0 )
Res= df['Calificación']
Caso= df['Caso']
matrix= numpy.zeros( ( len(Caso), len(Cal) ) ) - 1



pos= []
for i, res in enumerate( Res ):    
    for j, cal in enumerate(Cal):
        
        if res == cal:
            matrix[ i, j ] = 1            
            #pos.append( numpy.where( Res == aa )[0][0] )
        else:
            matrix[ i, j ] = -1


for i,cal in enumerate( Cal ):    
    matrix[ numpy.where( Res == cal )[0] , i ]= 1
   
  
    
  

from sklearn.datasets import load_iris
from sklearn import tree
iris = load_iris()
X, Y = iris.data, iris.target
clf = tree.DecisionTreeClassifier()
clf = clf.fit(X, Y)

import graphviz 
dot_data = tree.export_graphviz(clf, out_file=None) 
graph = graphviz.Source(dot_data) 
graph.render("iris")                        
"""








"""


#plot bar plot

Count= df['Caso'].value_counts()
plt.rcParams.update({'font.size': 12})
sns.barplot( Count.index, Count.values, alpha=0.8,)
plt.title('Resultados',fontsize=20)
plt.ylabel('Numero de ocurrencia', fontsize=20)
plt.xlabel('Diagnostico', fontsize=20)
plt.show()



#plot matrix
Res=df['Resultado'].unique()
Res=numpy.concatenate((Res,['OK']),axis=0)
Res1=df['Resultado']
Esp= df['Esperado']
matrix=numpy.zeros((len(Esp),len(Res)))

pos=[]
for i,z in enumerate(Esp):
    aa=Res1[i]
    if z==aa:
        matrix[i,numpy.where(Res=='OK')[0][0]]=1
        pos.append(numpy.where(Res==aa)[0][0])
    else:
        matrix[i,numpy.where(Res==aa)[0][0]] =-1

matrix=numpy.delete(matrix,pos,1)
Res=numpy.delete(Res,pos,0)

fig, ax = plt.subplots()
startcolor ='#FF0000'
midcolor = '#FFFFFF'
endcolor = '#008000'
othercolor = '#0080f0'



own_cmap1 = matplotlib.colors.LinearSegmentedColormap.from_list( 'own2', [startcolor, midcolor, endcolor, othercolor] )
mat = ax.imshow(matrix, cmap=own_cmap1, interpolation='nearest',vmin=-1, vmax=1)
plt.yticks(range(matrix.shape[0]), Caso )
plt.xticks(range(matrix.shape[1]), Res )
plt.xticks(rotation='vertical')
plt.tight_layout()
plt.show()
plt.savefig("output.png", bbox_inches="tight")
"""


"""
Eo= [ '0' for i in range(16)]
for i,jj in enumerate(DO):
    for k in range( Data.shape[0] ):

        if jj == list( Data[ list( Data.columns[:17] ) ].loc[ Data['Target']== k ].values[0]):
            Eo[i]= Data['Caso'].loc[ Data.Target== k ].values[0]
            
"""            

