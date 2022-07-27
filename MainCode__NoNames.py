import os
import struct
import numpy
import pandas as pd
from sklearn.tree import DecisionTreeClassifier
from sklearn.neighbors import KNeighborsClassifier
from sklearn.tree import export_graphviz
import seaborn as sns
import time
from datetime import datetime
from datetime import timedelta
from read_comtrade import ReadAllComtrade 
from scipy.ndimage import gaussian_filter1d
from Functions import *

"""
from Auto_TestBed_Hypersim import *

print( 'Test Time' ) 
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

"""
# Full_5NN
Test_Time= {'Caso_01': datetime(2022, 6, 14, 10, 17, 52, 183476), 'Caso_02': datetime(2022, 6, 14, 10, 18, 38, 641417), 
            'Caso_03': datetime(2022, 6, 14, 10, 19, 25, 424889), 'Caso_06': datetime(2022, 6, 14, 10, 20, 11, 829064), 
            'Caso_07': datetime(2022, 6, 14, 10, 20, 58, 175197), 'Caso_08': datetime(2022, 6, 14, 10, 21, 44, 815129), 
            'Caso_09': datetime(2022, 6, 14, 10, 22, 31, 586886), 'Caso_10': datetime(2022, 6, 14, 10, 23, 18, 238828), 
            'Caso_11': datetime(2022, 6, 14, 10, 24, 5, 122402), 'Caso_12': datetime(2022, 6, 14, 10, 24, 51, 680568), 
            'Caso_13': datetime(2022, 6, 14, 10, 25, 38, 67545), 'Caso_14': datetime(2022, 6, 14, 10, 26, 24, 496119), 
            'Caso_15': datetime(2022, 6, 14, 10, 27, 11, 27462), 'Caso_16': datetime(2022, 6, 14, 10, 27, 57, 563659), 
            'Caso_04': datetime(2022, 6, 14, 10, 28, 44, 50639), 'Caso_05': datetime(2022, 6, 14, 10, 29, 30, 251137)}

Carpeta= 'Ensayo Full5_NN'

"""


#"""
# Rel_
Test_Time= {'Caso_01': datetime(2022, 7, 12, 9, 57, 3, 684680), 'Caso_02': datetime(2022, 7, 12, 9, 57, 57, 25082), 
            'Caso_03': datetime(2022, 7, 12, 9, 58, 51, 870928), 'Caso_06': datetime(2022, 7, 12, 9, 59, 45, 235558), 
            'Caso_07': datetime(2022, 7, 12, 10, 0, 41, 495335), 'Caso_08': datetime(2022, 7, 12, 10, 1, 33, 523582), 
            'Caso_09': datetime(2022, 7, 12, 10, 2, 26, 526751), 'Caso_10': datetime(2022, 7, 12, 10, 3, 19, 779389), 
            'Caso_11': datetime(2022, 7, 12, 10, 4, 14, 269101), 'Caso_12': datetime(2022, 7, 12, 10, 5, 6, 795384), 
            'Caso_13': datetime(2022, 7, 12, 10, 5, 59, 183010), 'Caso_14': datetime(2022, 7, 12, 10, 6, 53, 114061), 
            'Caso_15': datetime(2022, 7, 12, 10, 7, 47, 99710), 'Caso_16': datetime(2022, 7, 12, 10, 8, 39, 669543), 
            'Caso_04': datetime(2022, 7, 12, 10, 9, 34, 155823), 'Caso_05': datetime(2022, 7, 12, 10, 10, 30, 313631)}

Carpeta= 'Ensayo Rel_Remoto_NN'
#"""



"""
# Rel_3
Test_Time= {'Caso_01': datetime(2022, 7, 15, 16, 13, 47, 869309), 'Caso_02': datetime(2022, 7, 15, 16, 14, 40, 190562), 
            'Caso_03': datetime(2022, 7, 15, 16, 15, 32, 838934), 'Caso_06': datetime(2022, 7, 15, 16, 16, 24, 925250), 
            'Caso_07': datetime(2022, 7, 15, 16, 17, 17, 54736), 'Caso_08': datetime(2022, 7, 15, 16, 18, 9, 262611), 
            'Caso_09': datetime(2022, 7, 15, 16, 19, 1, 758508), 'Caso_10': datetime(2022, 7, 15, 16, 19, 53, 919056), 
            'Caso_11': datetime(2022, 7, 15, 16, 20, 46, 612459), 'Caso_12': datetime(2022, 7, 15, 16, 21, 38, 756361), 
            'Caso_13': datetime(2022, 7, 15, 16, 22, 30, 925556), 'Caso_14': datetime(2022, 7, 15, 16, 23, 23, 520367), 
            'Caso_15': datetime(2022, 7, 15, 16, 24, 17, 87749), 'Caso_16': datetime(2022, 7, 15, 16, 25, 9, 571532), 
            'Caso_04': datetime(2022, 7, 15, 16, 26, 2, 175506), 'Caso_05': datetime(2022, 7, 15, 16, 26, 54, 200793)}

Carpeta= 'Ensayo Rel_Remoto_3_NN'

"""

"""
# Rel_21
Test_Time= {'Caso_01': datetime(2022, 7, 19, 10, 20, 40, 401567), 'Caso_02': datetime(2022, 7, 19, 10, 21, 27, 405756), 
            'Caso_03': datetime(2022, 7, 19, 10, 22, 14, 291150), 'Caso_06': datetime(2022, 7, 19, 10, 23, 1, 353444), 
            'Caso_07': datetime(2022, 7, 19, 10, 23, 48, 146320), 'Caso_08': datetime(2022, 7, 19, 10, 24, 35, 28677), 
            'Caso_09': datetime(2022, 7, 19, 10, 25, 21, 756587), 'Caso_10': datetime(2022, 7, 19, 10, 26, 8, 456634), 
            'Caso_11': datetime(2022, 7, 19, 10, 26, 55, 408401), 'Caso_12': datetime(2022, 7, 19, 10, 27, 42, 124108), 
            'Caso_13': datetime(2022, 7, 19, 10, 28, 29, 95953), 'Caso_14': datetime(2022, 7, 19, 10, 29, 15, 749278), 
            'Caso_15': datetime(2022, 7, 19, 10, 30, 2, 396403), 'Caso_16': datetime(2022, 7, 19, 10, 30, 49, 58218), 
            'Caso_04': datetime(2022, 7, 19, 10, 31, 35, 899325), 'Caso_05': datetime(2022, 7, 19, 10, 32, 22, 827402)}

Carpeta= 'Ensayo Rel_Remoto_21_Simple'
"""


"""
# Rel_4
Test_Time= {'Caso_01': datetime(2022, 7, 21, 11, 4, 9, 360284), 'Caso_02': datetime(2022, 7, 21, 11, 5, 1, 405772), 
            'Caso_03': datetime(2022, 7, 21, 11, 5, 53, 224191), 'Caso_06': datetime(2022, 7, 21, 11, 6, 45, 331547), 
            'Caso_07': datetime(2022, 7, 21, 11, 7, 37, 454908), 'Caso_08': datetime(2022, 7, 21, 11, 8, 29, 355079), 
            'Caso_09': datetime(2022, 7, 21, 11, 9, 21, 271117), 'Caso_10': datetime(2022, 7, 21, 11, 10, 13, 365899), 
            'Caso_11': datetime(2022, 7, 21, 11, 11, 5, 288405), 'Caso_12': datetime(2022, 7, 21, 11, 11, 57, 94975), 
            'Caso_13': datetime(2022, 7, 21, 11, 12, 48, 962572), 'Caso_14': datetime(2022, 7, 21, 11, 13, 40, 970949), 
            'Caso_15': datetime(2022, 7, 21, 11, 14, 32, 813240), 'Caso_16': datetime(2022, 7, 21, 11, 15, 24, 743590), 
            'Caso_04': datetime(2022, 7, 21, 11, 16, 16, 712487), 'Caso_05': datetime(2022, 7, 21, 11, 17, 9, 614627)}

Carpeta= 'Ensayo Rel_Remoto_4_NN'
"""


#Carpeta= 'Ensayo Full5_NN'                                                                                   # En esta carpeta se van a descargar los archivos Comtrade que se generan en el relé
pathh = 'D:\JoseM\Relé Siemens - Cama de Pruebas\Pruebas Fallas CID\Ensayos Automatismo Python'                     # Path de la carpeta anterior
path2= pathh + "\\" + Carpeta

if not os.path.exists( path2 ):                                                                                     # Verifica que la carpeta anterior exista, si no, la crea
    os.makedirs( path2 )


allF2 = []
allnames2=[]

for root, dirs, files in os.walk( path2, topdown= False):
   for name in files:
       if name.find('.cfg') > -1 or name.find('.CFG') > -1:
           allnames2.append(name)
           allF2.append(os.path.join(root, name))           
           
Comt_Time= {}
DO2= []

for file in allF2:
    ComtradeObjec2= ReadAllComtrade(file)                                                   # Firts step, create instance for the comtrade class
    ComtradeObjec2.ReadDataFile()                                                           # Next, read the data file
    Time= ComtradeObjec2.getTimeChannel()/1000000                                           # Time in us
    
    D1, T1= Digital( ComtradeObjec2.getDigitalASCCI(1), Time )                              # Trip A
    D2, T2= Digital( ComtradeObjec2.getDigitalASCCI(2), Time )                              # Trip B
    D3, T3= Digital( ComtradeObjec2.getDigitalASCCI(3), Time )                              # Trip C
    D4, T4= Digital( ComtradeObjec2.getDigitalASCCI(4), Time )                              # 21 Dir Forward
    D5, T5= Digital( ComtradeObjec2.getDigitalASCCI(5), Time )                              # 21 Dir Backward
    D6, T6= Digital( ComtradeObjec2.getDigitalASCCI(6), Time )                              # Trip Z4 reversal
    D7, T7= Digital( ComtradeObjec2.getDigitalASCCI(7), Time )                              # Trip Z2
    D8, T8= Digital( ComtradeObjec2.getDigitalASCCI(8), Time )                              # Trip Z1
    D9, T9= Digital( ComtradeObjec2.getDigitalASCCI(9), Time )                              # Trip 67N
    D10, T10= Digital( ComtradeObjec2.getDigitalASCCI(10), Time )                           # Trip O 85 - 67N  POTT
    D11, T11= Digital( ComtradeObjec2.getDigitalASCCI(11), Time )                           # Send O 85 - 67N  POTT
    D12, T12= Digital( ComtradeObjec2.getDigitalASCCI(12), Time )                           # Trip O 85 - 21  POTT
    D13, T13= Digital( ComtradeObjec2.getDigitalASCCI(13), Time )                           # Send O 85 - 21  POTT
   
    Digital_Out= [D1, D2, D3, D4, D5, D6, D7, D8, D9, D10, D11, D12, D13]                   # Construye el vector de 
    DO2.append( Digital_Out )
    
    Comt_Time[ file ]= { 'Comtrade Time': ComtradeObjec2.start,
                         'DO': Digital_Out }                                                # Crea un dict con los archivos como "keys" y los tiempos de captura de los comtrade como "values"
    


Comt_Time= TimeFormat3( Comt_Time )


DD= {}
for caso in Test_Time.keys():
    
    for file in Comt_Time.keys():               
        
        delta= Comt_Time[ file ]['Comtrade Time'] - Test_Time[ caso ]
                
        if delta <= timedelta( seconds= 5 ) and delta > timedelta( seconds= 0 ):
            DD[ caso  ]= { 'Execution Time': Test_Time[ caso ], 
                           'Comtrade Time': Comt_Time[ file ]['Comtrade Time'], 
                           'Comtrade File': file,
                           'Delta': delta,
                           'DO': Comt_Time[ file ]['DO'] }
            


DD= DefCaso_NN( DD )

EvaluateObject2= Evaluate()                                                                 # Create instance for Evaluation
Data2= EvaluateObject2.training()

Estimated2= []
for j,caso in enumerate(DD):                                                                
    DD[ caso  ]['Estimated']= EvaluateObject2.Result( DD[ caso ]['DO'] )                    # Evaluate each DO
    Estimated2.append( DD[ caso  ]['Estimated'] )

cols_Re= list( Data2.columns[18:-4] )                                                       # Catch the columns with the information for the results 
Re2= [ Data2[ cols_Re ].loc[ Data2['Target']== ii ].values[0] for ii in Estimated2 ]        # Create a matrix with the ordered information results 

                                          

Ress2= Re2[:] 
for ii,k in enumerate( DD ):                                                                # Se compara el vector de digitales que se captura de los comtrades con el vector de digitales que el DecisionTree asocia para el caso, con el fin de ver si el caso seleccionado sí es el correcto
    
    var2= list( Data2[ list( Data2.columns[:17] ) ].loc[ Data2['Target']== Estimated2[ii] ].values[0])
    
    if DD[k]['DO'] != var2:        
        Ress2[ii]= [k] + [ 'No Asociado' ] + [ ' - ' for i in range( len( Re2[ii] ) - 2 ) ]

"""
for i in DD:
    print( '[' + i + '] ' )
    print( DD[i] )
"""    
    

# Export Tree
#EvaluateObject.ExportTree( 'ElArbolito' )


# Export excel file result
df2 = pd.DataFrame( Ress2, columns= cols_Re )
folder= os.getcwd()
#File_exp= os.getcwd() + '\\' + 'Resultados.xlsx'
File_exp= path2 + '\\' + 'Resultados__' + Carpeta + '.xlsx'
#export_excel = df2.to_excel ( File_exp, index= None, header= True)


print( df2[['Caso','Calificación']] )

print( "Fin Pruebas")

                            

