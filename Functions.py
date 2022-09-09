# -*- coding: utf-8 -*-

import os
import struct
import numpy
import pandas as pd
import matplotlib.pyplot as plt
from sklearn.tree import DecisionTreeClassifier
from sklearn.neighbors import KNeighborsClassifier
from sklearn.tree import export_graphviz
import seaborn as sns
import time
from datetime import datetime
from read_comtrade import ReadAllComtrade 
from scipy.ndimage import gaussian_filter1d
from math import ceil
from sklearn import svm
from sklearn.preprocessing import MinMaxScaler




class Evaluate:
    # This algorithm evaluate and classifie comtrades acording to the values of the digitals values

    def __init__(self):
        #print('Starting evaluation decition tree!')
        return


    def training(self):
        
        # Data for trainig the algorithm
        self.Data= pd.read_excel( 'DataSet_TestBed__v10_1.xlsx', index_col= None )
                
        cols= self.Data.columns[:17]
        X  = self.Data[ cols ]
        Y= self.Data['Target']
        decision_tree = DecisionTreeClassifier( random_state= 0, max_depth= None )
        self.decision_tree = decision_tree.fit( X, Y )
        #print('Index of the Leave= ' )
        #print( decision_tree.apply( X ) )
        #print('Decision path= ' )
        #print( decision_tree.decision_path( X ) )
        #print('N° of leaves= ' )
        #print( decision_tree.get_n_leaves() )
        #print('Params of the DecisionTree= ' )
        #print( decision_tree.get_params() )

        #dot_data = export_graphviz(self.decision_tree, out_file=None) 
        #graph = graphviz.Source(dot_data) 
        #graph.render("ElArbolito")
        
        return self.Data


    def Result( self,value ):
        B= int( self.decision_tree.predict( [ value ] ) )

        return B

    def ExportTree(self,name):
        name= name+'.dot'
        export_graphviz(
            self.decision_tree,
            out_file= name,
            filled= True,
            rounded= True)
    

"""
class Evaluate2:
  

        X = self.Data['data']
        y = self.Data['target']
        self.Knn=KNeighborsClassifier()
        self.Knn.fit(X,y)

    def Result(self,value):
        B = int(self.Knn.predict([value]))
        return self.Data['target_name'][B]

"""



def Digital(Signal,Time):
    # This function return 1 if the binary input change from 0 to 1 in any position
    #  returns the activation time
    if numpy.sum(Signal) > 0:
        i, j = numpy.where(Signal == 1)
        TripTime= Time[i[0]]
        Trip= 1
    else:
        TripTime= 999
        Trip= 0

    return Trip,TripTime

  

def DefCaso( name ) -> int:
    # This function assing a 4 bits code depend on the input case     
       
    if name== 'Caso_01':
        A1, A2, A3, A4 = 0, 0, 0, 0
    
    elif name == 'Caso_02':
        A1, A2, A3, A4 = 0, 0, 0, 1
    
    elif name == 'Caso_03':
        A1, A2, A3, A4 = 0, 0, 1, 0
        
    elif name == 'Caso_04':
        A1, A2, A3, A4 = 0, 0, 1, 1
        
    elif name == 'Caso_05':
        A1, A2, A3, A4 = 0, 1, 0, 0
         
    elif name == 'Caso_06':
        A1, A2, A3, A4 = 0, 1, 0, 1
        
    elif name == 'Caso_07':
        A1, A2, A3, A4 = 0, 1, 1, 0
        
    elif name == 'Caso_08':
        A1, A2, A3, A4 = 0, 1, 1, 1
        
    elif name == 'Caso_09':
        A1, A2, A3, A4 = 1, 0, 0, 0
        
    elif name == 'Caso_10':
        A1, A2, A3, A4 = 1, 0, 0, 1
        
    elif name == 'Caso_11':
        A1, A2, A3, A4 = 1, 0, 1, 0
        
    elif name == 'Caso_12':
        A1, A2, A3, A4 = 1, 0, 1, 1
        
    elif name == 'Caso_13':
        A1, A2, A3, A4 = 1, 1, 0, 0
        
    elif name == 'Caso_14':
        A1, A2, A3, A4 = 1, 1, 0, 1
        
    elif name == 'Caso_15':
        A1, A2, A3, A4 = 1, 1, 1, 0
        
    elif name == 'Caso_16':
        A1, A2, A3, A4 = 1, 1, 1, 1 
        
        
    return [A1, A2, A3, A4]
 




def DefCaso2( dict_time, DOs ) -> dict:
    # This function assing a 4 bits code depend on the input case     
    
    N= list( dict_time.values() )                                   # Toma los tiempos del diccionario de casos donde están los archivos como 'keys' y los tiempos de los comtrades como 'values'        
    N.sort()                                                        # Organiza los tiempos de orden ascendente (menor a mayor) 

    M= {}
    count= 0
    for j in N:                                                     # Cada tiempo de la lista organizada (para asignarlas en orden)
        for k,filee in enumerate( dict_time.keys() ):               # Saca los 'file'       
            
            if j == dict_time[ filee ]:                             # La idea es en este loop buscar el 'value' igualito al 'j' para asignarlo a un nuevo diccionario 'M' con los datos ordenados
                
                count += 1
                M[ filee ]= { 'Tiempo': j,                          # Diccionario con los archivos organizados cronológicamente, donde cada archivo contiene un diccionario que guarda el 'Tiempo' del evento y define el 'Caso' SEGÚN ORDEN ASCENDENTE
                              'Caso': 'Caso_' + str( count ) }                            
        
                if M[ filee ]['Caso'] == 'Caso_01':
                    A1, A2, A3, A4 = 0, 0, 0, 0
                                    
                elif M[ filee ]['Caso'] == 'Caso_02':
                    A1, A2, A3, A4 = 0, 0, 0, 1
                                    
                elif M[ filee ]['Caso'] == 'Caso_03':
                    A1, A2, A3, A4 = 0, 0, 1, 0
                                        
                elif M[ filee ]['Caso'] == 'Caso_04':
                    A1, A2, A3, A4 = 0, 0, 1, 1 
                    
                elif M[ filee ]['Caso'] == 'Caso_05':
                    A1, A2, A3, A4 = 0, 1, 0, 0
                                         
                elif M[ filee ]['Caso'] == 'Caso_06':
                    A1, A2, A3, A4 = 0, 1, 0, 1
                                        
                elif M[ filee ]['Caso'] == 'Caso_07':
                    A1, A2, A3, A4 = 0, 1, 1, 0
                                        
                elif M[ filee ]['Caso'] == 'Caso_08':
                    A1, A2, A3, A4 = 0, 1, 1, 1
                                        
                elif M[ filee ]['Caso'] == 'Caso_09':
                    A1, A2, A3, A4 = 1, 0, 0, 0
                                        
                elif M[ filee ]['Caso'] == 'Caso_10':
                    A1, A2, A3, A4 = 1, 0, 0, 1
                                        
                elif M[ filee ]['Caso'] == 'Caso_11':
                    A1, A2, A3, A4 = 1, 0, 1, 0
                                       
                elif M[ filee ]['Caso'] == 'Caso_12':
                    A1, A2, A3, A4 = 1, 0, 1, 1
                                       
                elif M[ filee ]['Caso'] == 'Caso_13':
                    A1, A2, A3, A4 = 1, 1, 0, 0
                                        
                elif M[ filee ]['Caso'] == 'Caso_14':
                    A1, A2, A3, A4 = 1, 1, 0, 1
                                        
                elif M[ filee ]['Caso'] == 'Caso_15':
                    A1, A2, A3, A4 = 1, 1, 1, 0
                    
                elif M[ filee ]['Caso'] == 'Caso_16':
                    A1, A2, A3, A4 = 1, 1, 1, 1                     
        
               # M[ filee ]['DO']= [A1, A2, A3, A4] + DOs[ k ]
                M[ filee ]['DO']= DOs[ k ] + [A1, A2, A3, A4]
        
        
    return M
  


def DefCaso_NN( M ) -> dict:
                                                        
    for k,caso in enumerate( M.keys() ):                        # Saca los 'file'       
        
            if caso == 'Caso 1':
                A1, A2, A3, A4 = 0, 0, 0, 0
                                
            elif caso == 'Caso 2':
                A1, A2, A3, A4 = 0, 0, 0, 1
                
            elif caso == 'Caso 3':
                A1, A2, A3, A4 = 0, 0, 1, 0
                                    
            elif caso == 'Caso 4':
                A1, A2, A3, A4 = 0, 0, 1, 1 
                
            elif caso == 'Caso 5':
                A1, A2, A3, A4 = 0, 1, 0, 0
                                     
            elif caso == 'Caso 6':
                A1, A2, A3, A4 = 0, 1, 0, 1
                                    
            elif caso == 'Caso 7':
                A1, A2, A3, A4 = 0, 1, 1, 0
                                    
            elif caso == 'Caso 8':
                A1, A2, A3, A4 = 0, 1, 1, 1
                                    
            elif caso == 'Caso 9':
                A1, A2, A3, A4 = 1, 0, 0, 0
                                    
            elif caso == 'Caso 10':
                A1, A2, A3, A4 = 1, 0, 0, 1
                                    
            elif caso == 'Caso 11':
                A1, A2, A3, A4 = 1, 0, 1, 0
                                   
            elif caso == 'Caso 12':
                A1, A2, A3, A4 = 1, 0, 1, 1
                                   
            elif caso == 'Caso 13':
                A1, A2, A3, A4 = 1, 1, 0, 0
                                    
            elif caso == 'Caso 14':
                A1, A2, A3, A4 = 1, 1, 0, 1
                                    
            elif caso == 'Caso 15':
                A1, A2, A3, A4 = 1, 1, 1, 0
                
            elif caso == 'Caso 16':
                A1, A2, A3, A4 = 1, 1, 1, 1 
                            
            M[ caso ]['DO']= M[ caso ]['DO'] + [A1, A2, A3, A4]
    
    
    M= { caso:M[ caso ] for caso in sorted( M.keys() ) }                                # Crea un nuevo diccionario con la misma información, pero ordenada respecto a los casos (Casos 4 y 5 donde deben ir, no al final)

    return M



def TimeFormat( Time ) -> str:
    
    y= int( Time.split('-')[0][-2:] )
    m= int( Time.split('-')[1] )
    d= int( Time.split()[0].split('-')[2] )
    h= int( Time.split()[1].split(':')[0] )
    mm= int( Time.split()[1].split(':')[1] )
    s= int( float( Time.split()[1].split(':')[2] ) )
    
    Timee= [y, m, d, h, mm, s]
   
    return Timee



def TimeFormat2( Time ) -> str:
    
    for i in Time.keys():
        
        y= int( Time[i].split('-')[0][-2:] )
        m= int( Time[i].split('-')[1] )
        d= int( Time[i].split()[1].split(':')[0] )
        h= int( Time[i].split()[1].split(':')[0] )
        mm= int( Time[i].split()[1].split(':')[1] )
        s= int( float( Time[i].split()[1].split(':')[2] ) )
    
        Time[i]= [y, m, d, h, mm, s]
    
    return Time


def TimeFormat3( TT ):
    
    #DateTime= {}
    
    for ii in TT:
        
        y= TT[ ii ]['Comtrade Time'][0]
        m= TT[ ii ]['Comtrade Time'][1]
        d= TT[ ii ]['Comtrade Time'][2]
        h= TT[ ii ]['Comtrade Time'][3]
        mm= TT[ ii ]['Comtrade Time'][4]
        s= TT[ ii ]['Comtrade Time'][5]
        
        TT[ ii ]['Comtrade Time']= datetime.strptime( str(y) + '/' + str(m) + '/' + str(d) + ',' + str(h) + ':' + str(mm) + ':' + str(s), '%y/%m/%d,%H:%M:%S'  )
        
    
    #DateTime= dict( sorted( DateTime.items(), key= lambda x: x[1] ) )                                           # Este organiza el diccionario de menor a mayor basado en los "values" (key: VALUE). La función genera una lista de tuplas organizadas, y luego vuelve y crea el diccionario
        
        
    return TT
    
    
    
    

"""

NN= {}
DO2= []

for file in allfiles:
    ComtradeObjec2= ReadAllComtrade(file)                               # Firts step, create instance for the comtrade class
    ComtradeObjec2.ReadDataFile()                                       # Next, read the data file
    Time=ComtradeObjec2.getTimeChannel()/1000000                        # Time in us
    
    D1, T1= Digital( ComtradeObjec2.getDigitalASCCI(1), Time )          # Trip a
    D2, T2= Digital( ComtradeObjec2.getDigitalASCCI(2), Time )          # Trip b
    D3, T3= Digital( ComtradeObjec2.getDigitalASCCI(3), Time )          # Trip c
    D4, T4= Digital( ComtradeObjec2.getDigitalASCCI(4), Time )          # 21 Dir Forward
    D5, T5= Digital( ComtradeObjec2.getDigitalASCCI(5), Time )          # 21 Dir Backward
    D6, T6= Digital( ComtradeObjec2.getDigitalASCCI(6), Time )          # Trip Z4 reversal
    D7, T7= Digital( ComtradeObjec2.getDigitalASCCI(7), Time )          # Trip Z2
    D8, T8= Digital( ComtradeObjec2.getDigitalASCCI(8), Time )          # Trip Z1
    D9, T9= Digital( ComtradeObjec2.getDigitalASCCI(9), Time )          # Trip 67N
    D10, T10= Digital( ComtradeObjec2.getDigitalASCCI(10), Time )       # Trip O 85 - 67N  POTT
    D11, T11= Digital( ComtradeObjec2.getDigitalASCCI(11), Time )       # Send O 85 - 67N  POTT
    D12, T12= Digital( ComtradeObjec2.getDigitalASCCI(12), Time )       # Trip O 85 - 21  POTT
    D13, T13= Digital( ComtradeObjec2.getDigitalASCCI(13), Time )       # Send O 85 - 21  POTT
   
    Digital_Out= [D1, D2, D3, D4, D5, D6, D7, D8, D9, D10, D11, D12, D13]        # Construye el vector de 
    DO2.append( Digital_Out )
    
    NN[ file ]= ComtradeObjec2.start
    


MM= DefCaso2( NN, DO2 )
   
"""
    
    
  










