# -*- coding: latin-1 -*-
import os
import sys
sys.path.append(r'C:\OPAL-RT\HYPERSIM\hypersim_2021.1.2.o150\Windows\HyApi\C\py')
sys.path.append(r'C:\OPAL-RT\HYPERSIM\hypersim_2021.1.2.o150\Windows\ScopeView\lib')
import HyWorksApi
import ScopeViewApi
import time
from datetime import datetime
from datetime import timedelta
import xlwings as xlw
import csv

import pandas as pd
#from datetime import *
from read_comtrade import ReadAllComtrade 
from Functions import *


time.sleep(3)
# Colocar acá el path de los archivos necesarios para la prueba


#"""
designPath = r'D:\JoseM\Cama_de_Pruebas\BancoDePruebasV38_2___V2021__Target2__Rel_Remoto\BancoDePruebasV38_2___V2021__Target2__Rel_Remoto.ecf'
hy_sensors = r'D:\JoseM\Cama_de_Pruebas\BancoDePruebasV38_2___V2021__Target2__Rel_Remoto\SensoresV38_2___V2021__Target2__Cierre_En_Falla.sig'
sv_template = r'D:\JoseM\Relé Siemens - Cama de Pruebas\Template_ScopeView__TestBed__Rel_Rem__CierreEnFalla.svt'
#"""


# Carpeta y path para Comtrades
Carpeta= 'Ensayo CierreEnFalla_1'                                                                              # En esta carpeta se van a descargar los archivos Comtrade que se generan en el relé
pathh = 'D:\JoseM\Cama_de_Pruebas\Pruebas Fallas CID\Ensayos Automatismo Python'                     # Path de la carpeta anterior
path2= pathh + "\\" + Carpeta

if not os.path.exists( path2 ):                                                                                     # Verifica que la carpeta anterior exista, si no, la crea
    os.makedirs( path2 )




def Conf_CB_Ant( PosA, PosB, PosC ):
    
    HyWorksApi.setComponentParameter( 'CB_Recierre_Ant', 'EtatIniA', str( PosA ) ) 
    HyWorksApi.setComponentParameter( 'CB_Recierre_Ant', 'EtatIniB', str( PosB ) ) 
    HyWorksApi.setComponentParameter( 'CB_Recierre_Ant', 'EtatIniC', str( PosC ) ) 
    

def Conf_CB_Cerro( PosA, PosB, PosC ):

    HyWorksApi.setComponentParameter( 'CB_Recierre_Cerro', 'EtatIniA', str( PosA ) ) 
    HyWorksApi.setComponentParameter( 'CB_Recierre_Cerro', 'EtatIniB', str( PosB ) ) 
    HyWorksApi.setComponentParameter( 'CB_Recierre_Cerro', 'EtatIniC', str( PosC ) )




# Inicio del Automatismo que ejecuta las pruebas en Hypersim

# Arrancar Hs (Hypersim), abrir el caso base, carga archivo de sensores, abre SV (Scope View),
#  carga template de mediciones, analiza, map task, compila caso de HS, corre caso base y realiza una adquisición para el caso base

HyWorksApi.startAndConnectHypersim()
#HyWorksApi.startHyperWorks(stdout=None, stderr=None)
#HyWorksApi.connectToHyWorks(host='localhost',timeout=180000)


HyWorksApi.openDesign(designPath)
HyWorksApi.clearCodeDir()
HyWorksApi.loadSensors (hy_sensors)
HyWorksApi.analyze()
HyWorksApi.mapTask()
HyWorksApi.genCode()


"""
HyWorksApi.setComponentParameter( 'ClearLEDs', 'K', '1')
time.sleep(1)
HyWorksApi.setComponentParameter( 'ClearLEDs', 'K', '0')
"""

HyWorksApi.setComponentParameter( 'CB_Recierre_Ant', 'EnaGen', '1')                                             
HyWorksApi.setComponentParameter( 'CB_Recierre_Ant', 'CmdBlockSelect', 'Internal')                       
HyWorksApi.setComponentParameter( 'CB_Recierre_Cerro', 'EnaGen', '1')
HyWorksApi.setComponentParameter( 'CB_Recierre_Cerro', 'CmdBlockSelect', 'Internal')

HyWorksApi.setComponentParameter( 'CP1', 'EnaGen', '1')
HyWorksApi.setComponentParameter( 'CP3', 'EnaGen', '0')

HyWorksApi.startSim()
print('Simulacion en ejecucion... Caso base')
HyWorksApi.takeSnapshot()
print('Snapshot de Caso base tomado')
time.sleep(5)

ScopeViewApi.openScopeView()
ScopeViewApi.loadTemplate(sv_template)
ScopeViewApi.setSync(True)
ScopeViewApi.setTrig(False)
ScopeViewApi.startAcquisition()
print('Adquisicion... Caso base')
HyWorksApi.setComponentParameter( 'CP1', 'EnaGen', '1')
HyWorksApi.setComponentParameter( 'CP3', 'EnaGen', '0')
time.sleep(3)

Test_Time= []



# Caso 1A__BUS: Línea abierta con falla "pegada". Se simula una prefalla y luego de 1 segundo, se cierran el interruptor BUS 
#  para ver si la falla continúa. Se espera durante 1 segundo de posfalla y finaliza el caso. Esta joda se repite pero con el TIE

print('Caso 1A: BUS')

# Configurar la falla en la línea: Falla trifásica de 1 Ohm al 50% de la línea

HyWorksApi.setComponentParameter( 'CP1','fault_loc', '56' )
HyWorksApi.setComponentParameter( 'CP1','RDef', '1' )
HyWorksApi.setComponentParameter( 'CP1','EnaT1', '1' )
HyWorksApi.setComponentParameter( 'CP1','EnaT2', '1' )
HyWorksApi.setComponentParameter( 'CP1','T1', '0.5' )
HyWorksApi.setComponentParameter( 'CP1','T2', '2' )
HyWorksApi.setComponentParameter( 'CP1','T1Pa', '1' )
HyWorksApi.setComponentParameter( 'CP1','T1Pb', '1' )
HyWorksApi.setComponentParameter( 'CP1','T1Pc', '1' )
HyWorksApi.setComponentParameter( 'CP1','T1Pg', '0' )
HyWorksApi.setComponentParameter( 'CP1','T2Pa', '1' )
HyWorksApi.setComponentParameter( 'CP1','T2Pb', '1' )
HyWorksApi.setComponentParameter( 'CP1','T2Pc', '1' )
HyWorksApi.setComponentParameter( 'CP1','T2Pg', '0' )


# Apertura de CBs

HyWorksApi.setComponentParameter( 'CB_Recierre_Ant','EnaT1', '1' )
HyWorksApi.setComponentParameter( 'CB_Recierre_Ant','EnaT2', '1' )
HyWorksApi.setComponentParameter( 'CB_Recierre_Ant','T1', '0' )
HyWorksApi.setComponentParameter( 'CB_Recierre_Ant','T2', '1' )
HyWorksApi.setComponentParameter( 'CB_Recierre_Ant','T1Pa', '1' )
HyWorksApi.setComponentParameter( 'CB_Recierre_Ant','T1Pb', '1' )
HyWorksApi.setComponentParameter( 'CB_Recierre_Ant','T1Pc', '1' )
HyWorksApi.setComponentParameter( 'CB_Recierre_Ant','T2Pa', '1' )
HyWorksApi.setComponentParameter( 'CB_Recierre_Ant','T2Pb', '1' )
HyWorksApi.setComponentParameter( 'CB_Recierre_Ant','T2Pc', '1' )

HyWorksApi.setComponentParameter( 'CB_Recierre_Cerro','EnaT1', '1' )
HyWorksApi.setComponentParameter( 'CB_Recierre_Cerro','T1', '0' )
HyWorksApi.setComponentParameter( 'CB_Recierre_Cerro','T1Pa', '1' )
HyWorksApi.setComponentParameter( 'CB_Recierre_Cerro','T1Pb', '1' )
HyWorksApi.setComponentParameter( 'CB_Recierre_Cerro','T1Pc', '1' )


ScopeViewApi.setTrig(True)
Test_Time.append( datetime.now() )

ScopeViewApi.startAcquisition()
print( 'Tiempo de inyeccion = '+ str( Test_Time[-1] ) )

HyWorksApi.loadSnapshot()
time.sleep(25)
ScopeViewApi.setTrig(False)
ScopeViewApi.startAcquisition()



print('Siguiente caso...')

time.sleep(5)




# Caso 1B__TIE: Línea abierta con falla "pegada". Se simula una prefalla y luego de 1 segundo, se cierran el interruptor TIE
#  para ver si la falla continúa. Se espera durante 1 segundo de posfalla y finaliza el caso.


print('Caso 1B: TIE')

# Configurar la falla en la línea: Falla trifásica de 1 Ohm al 50% de la línea

HyWorksApi.setComponentParameter( 'CP1','fault_loc', '56' )
HyWorksApi.setComponentParameter( 'CP1','RDef', '1' )
HyWorksApi.setComponentParameter( 'CP1','EnaT1', '1' )
HyWorksApi.setComponentParameter( 'CP1','EnaT2', '1' )
HyWorksApi.setComponentParameter( 'CP1','T1', '0.5' )
HyWorksApi.setComponentParameter( 'CP1','T2', '2' )
HyWorksApi.setComponentParameter( 'CP1','T1Pa', '1' )
HyWorksApi.setComponentParameter( 'CP1','T1Pb', '1' )
HyWorksApi.setComponentParameter( 'CP1','T1Pc', '1' )
HyWorksApi.setComponentParameter( 'CP1','T1Pg', '0' )
HyWorksApi.setComponentParameter( 'CP1','T2Pa', '1' )
HyWorksApi.setComponentParameter( 'CP1','T2Pb', '1' )
HyWorksApi.setComponentParameter( 'CP1','T2Pc', '1' )
HyWorksApi.setComponentParameter( 'CP1','T2Pg', '0' )


# Apertura de CBs

HyWorksApi.setComponentParameter( 'CB_Recierre_Cerro','EnaT1', '1' )
HyWorksApi.setComponentParameter( 'CB_Recierre_Cerro','EnaT2', '1' )
HyWorksApi.setComponentParameter( 'CB_Recierre_Cerro','T1', '0' )
HyWorksApi.setComponentParameter( 'CB_Recierre_Cerro','T2', '1' )
HyWorksApi.setComponentParameter( 'CB_Recierre_Cerro','T1Pa', '1' )
HyWorksApi.setComponentParameter( 'CB_Recierre_Cerro','T1Pb', '1' )
HyWorksApi.setComponentParameter( 'CB_Recierre_Cerro','T1Pc', '1' )
HyWorksApi.setComponentParameter( 'CB_Recierre_Cerro','T2Pa', '1' )
HyWorksApi.setComponentParameter( 'CB_Recierre_Cerro','T2Pb', '1' )
HyWorksApi.setComponentParameter( 'CB_Recierre_Cerro','T2Pc', '1' )

HyWorksApi.setComponentParameter( 'CB_Recierre_Ant','EnaT1', '1' )
HyWorksApi.setComponentParameter( 'CB_Recierre_Ant','T1', '0' )
HyWorksApi.setComponentParameter( 'CB_Recierre_Ant','T1Pa', '1' )
HyWorksApi.setComponentParameter( 'CB_Recierre_Ant','T1Pb', '1' )
HyWorksApi.setComponentParameter( 'CB_Recierre_Ant','T1Pc', '1' )




ScopeViewApi.setTrig(True)
Test_Time.append( datetime.now() )

ScopeViewApi.startAcquisition()
print( 'Tiempo de inyeccion = '+ str( Test_Time[-1] ) )

HyWorksApi.loadSnapshot()
time.sleep(25)
ScopeViewApi.setTrig(False)
ScopeViewApi.startAcquisition()



print('Siguiente caso...')

time.sleep(5)




#ii= 0

for caso in range(len(Parampruebalin['caso'])):
#for caso in range( 3 ):
    
    #ii += 1
    #caso= 7+ii

    HyWorksApi.setComponentParameter( 'CP1', 'EnaGen', '1' )
    HyWorksApi.setComponentParameter( 'CP3', 'EnaGen', '0' )

    # Carga snapshot del caso base antes de inyectar el siguiente caso
    print('Ejecutando Caso %i' %( int(Parampruebalin['caso'][caso]) ) )
    para = 'Parametros--->'+'  Elemento-->'+str(Parampruebalin['elemento'][caso])+'  Tipo Falla-->'+str(Parampruebalin['tipo'][caso])+'  Distancia-->'+ str(Parampruebalin['distancia'][caso])+'  T falla-->'+str(Parampruebalin['tiempofalla'][caso])+'  T Clear-->'+str(Parampruebalin['tiempoclear'][caso])+'  R falla-->'+str(Parampruebalin['rfalla'][caso])
    print(para)
    # Cambiar parametros de acuerdo con el caso de falla HyWorksApi.setComponentParameter('componenete','parametro','valor')
    HyWorksApi.setComponentParameter(Parampruebalin['elemento'][caso],'fault_loc',str(Parampruebalin['fault_loc'][caso]))
    HyWorksApi.setComponentParameter(Parampruebalin['elemento'][caso],'RDef',str(Parampruebalin['RDef'][caso]))
    HyWorksApi.setComponentParameter(Parampruebalin['elemento'][caso],'EnaT1',str(Parampruebalin['EnaT1'][caso]))
    HyWorksApi.setComponentParameter(Parampruebalin['elemento'][caso],'EnaT2',str(Parampruebalin['EnaT2'][caso]))
    HyWorksApi.setComponentParameter(Parampruebalin['elemento'][caso],'EnaT3',str(Parampruebalin['EnaT3'][caso]))
    HyWorksApi.setComponentParameter(Parampruebalin['elemento'][caso],'EnaT4',str(Parampruebalin['EnaT4'][caso]))
    HyWorksApi.setComponentParameter(Parampruebalin['elemento'][caso],'T1',str(Parampruebalin['T1'][caso]))
    HyWorksApi.setComponentParameter(Parampruebalin['elemento'][caso],'T2',str(Parampruebalin['T2'][caso]))
    HyWorksApi.setComponentParameter(Parampruebalin['elemento'][caso],'T3',str(Parampruebalin['T3'][caso]))
    HyWorksApi.setComponentParameter(Parampruebalin['elemento'][caso],'T4',str(Parampruebalin['T4'][caso]))
    HyWorksApi.setComponentParameter(Parampruebalin['elemento'][caso],'T1Pa',str(Parampruebalin['T1Pa'][caso]))
    HyWorksApi.setComponentParameter(Parampruebalin['elemento'][caso],'T1Pb',str(Parampruebalin['T1Pb'][caso]))
    HyWorksApi.setComponentParameter(Parampruebalin['elemento'][caso],'T1Pc',str(Parampruebalin['T1Pc'][caso]))
    HyWorksApi.setComponentParameter(Parampruebalin['elemento'][caso],'T1Pg',str(Parampruebalin['T1Pg'][caso]))
    HyWorksApi.setComponentParameter(Parampruebalin['elemento'][caso],'T2Pa',str(Parampruebalin['T2Pa'][caso]))
    HyWorksApi.setComponentParameter(Parampruebalin['elemento'][caso],'T2Pb',str(Parampruebalin['T2Pb'][caso]))
    HyWorksApi.setComponentParameter(Parampruebalin['elemento'][caso],'T2Pc',str(Parampruebalin['T2Pc'][caso]))
    HyWorksApi.setComponentParameter(Parampruebalin['elemento'][caso],'T2Pg',str(Parampruebalin['T2Pg'][caso]))
    
    if Parampruebalin['caso'][caso] == 11 or Parampruebalin['caso'][caso] == 12: 
    
        HyWorksApi.setComponentParameter( 'CP1', 'EnaGen', '0' )
        HyWorksApi.setComponentParameter( 'CP3', 'EnaGen', '1' )
        HyWorksApi.setComponentParameter( 'CB2', 'EnaGen', '1' )
        HyWorksApi.setComponentParameter( 'CB1', 'EnaGen', '1' )
        

        HyWorksApi.setComponentParameter('CB2', 'EnaT1', '1' )
        #HyWorksApi.setComponentParameter('CB2', 'EnaT2', '1' )
        HyWorksApi.setComponentParameter('CB2', 'T1', str( Parampruebalin['T1'][caso]+0.07) )
        #HyWorksApi.setComponentParameter('CB2', 'T2', str( Parampruebalin['T2'][caso]) )
        HyWorksApi.setComponentParameter('CB2','T1Pa',str(1))
        HyWorksApi.setComponentParameter('CB2','T1Pb',str(1))
        HyWorksApi.setComponentParameter('CB2','T1Pc',str(1))
    
       #HyWorksApi.setComponentParameter('CB2','T2Pa',str(1))
        #HyWorksApi.setComponentParameter('CB2','T2Pb',str(1))
        #HyWorksApi.setComponentParameter('CB2','T2Pc',str(1))
    
        HyWorksApi.setComponentParameter('CB1', 'EnaT1', str( 1 ) )
        #HyWorksApi.setComponentParameter('CB1', 'EnaT2', str( 1 ) )
        HyWorksApi.setComponentParameter('CB1', 'T1', str( Parampruebalin['T1'][caso]+0.08) )
        #HyWorksApi.setComponentParameter('CB1', 'T2', str( Parampruebalin['T2'][caso]) )
        HyWorksApi.setComponentParameter('CB1','T1Pa',str(1))
        HyWorksApi.setComponentParameter('CB1','T1Pb',str(1))
        HyWorksApi.setComponentParameter('CB1','T1Pc',str(1))
 
        #HyWorksApi.setComponentParameter('CB1','T2Pa',str(1))
        #HyWorksApi.setComponentParameter('CB1','T2Pb',str(1))
        #HyWorksApi.setComponentParameter('CB1','T2Pc',str(1))
 
        

    ScopeViewApi.setTrig(True)
    Parampruebalin['t_inyect'].append(str(datetime.now()))
    Test_Time[ 'Caso_' + str( int( Parampruebalin['caso'][caso] )).zfill(2) ]= datetime.now()
    ScopeViewApi.startAcquisition()
    print('Tiempo de inyeccion = '+ Parampruebalin['t_inyect'][-1])

    HyWorksApi.loadSnapshot()
    time.sleep(25)
    ScopeViewApi.setTrig(False)
    ScopeViewApi.startAcquisition()
    
    if Parampruebalin['caso'][caso] == 11 or Parampruebalin['caso'][caso] == 12: 
        HyWorksApi.setComponentParameter( 'CB2', 'EnaGen', '0' )
        HyWorksApi.setComponentParameter( 'CB1', 'EnaGen', '0' )
    
    print('Siguiente caso...')

print('Fin del set de pruebas en líneas')


HyWorksApi.setComponentParameter( 'CP1', 'EnaGen', '0' )                # Se apagan ambas líneas para q no entren en las fallas de las barras
HyWorksApi.setComponentParameter( 'CP3', 'EnaGen', '0' )

# Loop para inyectar cada caso de prueba de barras. Se leen los parámetros de cada caso y asignan a las barras, inyecta con SV y regresa al estado estacionario inicial
for casobar in range(len(Parampruebabar['caso'])):
    
    if Parampruebabar['caso'][casobar] == 4:     
        HyWorksApi.setComponentParameter( 'Flt2', 'EnaGen', '1' )
        HyWorksApi.setComponentParameter( 'Flt3', 'EnaGen', '0' )

    elif Parampruebabar['caso'][casobar] == 5:
        HyWorksApi.setComponentParameter( 'Flt2', 'EnaGen', '0' )
        HyWorksApi.setComponentParameter( 'Flt3', 'EnaGen', '1' )
    
    # Carga snapshot del caso base antes de inyectar el siguiente caso
    print('Ejecutando Caso %i' %(int(Parampruebabar['caso'][casobar])) )
    para = 'Parametros--->' + '  Elemeto-->' + str(Parampruebabar['elemento'][casobar]) + '  Tipo Falla-->' + str(Parampruebabar['tipo'][casobar]) + '  T falla-->' + str(Parampruebabar['tiempofalla'][casobar]) + '  T Clear-->' + str(Parampruebabar['tiempoclear'][casobar]) + '  R falla-->' + str(Parampruebabar['rfalla'][casobar])
    print( para )
    
    # Cambiar parametros de acuerdo con el caso de falla HyWorksApi.setComponentParameter('componenete','parametro','valor')
    HyWorksApi.setComponentParameter(Parampruebabar['elemento'][casobar], 'RClose', str(Parampruebabar['rfalla'][casobar]))
    HyWorksApi.setComponentParameter(Parampruebabar['elemento'][casobar], 'EnaT1', str( int( Parampruebabar['EnaT1'][casobar])) )
    HyWorksApi.setComponentParameter(Parampruebabar['elemento'][casobar], 'EnaT2', str( int( Parampruebabar['EnaT2'][casobar])) )
    HyWorksApi.setComponentParameter(Parampruebabar['elemento'][casobar], 'T1', str(Parampruebabar['T1'][casobar]))
    HyWorksApi.setComponentParameter(Parampruebabar['elemento'][casobar], 'T2', str(Parampruebabar['T2'][casobar]))
    HyWorksApi.setComponentParameter(Parampruebabar['elemento'][casobar], 'T1Pa', str(Parampruebabar['T1Pa'][casobar]))
    HyWorksApi.setComponentParameter(Parampruebabar['elemento'][casobar], 'T1Pb', str(Parampruebabar['T1Pb'][casobar]))
    HyWorksApi.setComponentParameter(Parampruebabar['elemento'][casobar], 'T1Pc', str(Parampruebabar['T1Pc'][casobar]))
    HyWorksApi.setComponentParameter(Parampruebabar['elemento'][casobar], 'T1Pg', str(Parampruebabar['T1Pg'][casobar]))
    HyWorksApi.setComponentParameter(Parampruebabar['elemento'][casobar], 'T2Pa', str(Parampruebabar['T2Pa'][casobar]))
    HyWorksApi.setComponentParameter(Parampruebabar['elemento'][casobar], 'T2Pb', str(Parampruebabar['T2Pb'][casobar]))
    HyWorksApi.setComponentParameter(Parampruebabar['elemento'][casobar], 'T2Pc', str(Parampruebabar['T2Pc'][casobar]))
    HyWorksApi.setComponentParameter(Parampruebabar['elemento'][casobar], 'T2Pg', str(Parampruebabar['T2Pg'][casobar]))

    # HyWorksApi.stopSim()                                                                                  # Esto se pone si se va a simular en el Target1
    # HyWorksApi.removeDevice('POW1')
    # HyWorksApi.addDevice('Network Tools.clf', 'Point-on-wave', 11200, -9700, 1)
    # HyWorksApi.connectDeviceToBus3ph('POW1', 'net_1', 'pownode')
    # HyWorksApi.startSim()
    ScopeViewApi.setTrig(True)
    Parampruebabar['t_inyect'].append( str(datetime.now() ))
    Test_Time[ 'Caso_' + str( int( Parampruebabar['caso'][casobar] )).zfill(2) ]=  datetime.now()
    ScopeViewApi.startAcquisition()
    
    # HyWorksApi.stopSim()                                                                                  # Esto se pone si se va a simular en el Target1
    # HyWorksApi.startSim()
    #   HyWorksApi.takeSnapshot()
    #   print colored('Snapshot de Caso base tomado', 'yellow')
    #   print colored('Cargando snapshot del caso base', 'yellow')
    HyWorksApi.loadSnapshot()
    time.sleep(25)
    ScopeViewApi.setTrig(False)
    ScopeViewApi.startAcquisition()

HyWorksApi.setComponentParameter( 'Flt2', 'EnaGen', '0' )
HyWorksApi.setComponentParameter( 'Flt3', 'EnaGen', '0' )

Test_Time11= Parampruebalin['t_inyect'][:3] + Parampruebabar['t_inyect'] + Parampruebalin['t_inyect'][3:]


HyWorksApi.stopSim()
ScopeViewApi.close()
HyWorksApi.closeDesign(designPath)
HyWorksApi.closeHyperWorks()


print( 'Test Time' ) 
print( Test_Time )

print('Fin Pruebas')





# Inicia código para analizar los archivos comtrade y entrenar el Desicion Tree

input('Donwload the Comtrade Files into the specific path and Press Enter to continue')                     # Tiempo para descargar los Comtrades la carpeta especificada y presionar 'Enter' para continuar la ejecución del código
print('Next...')



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
    D3, T3= Digital( ComtradeObjec2.getDigitalASCCI(3), Time )                              # Trip CD3
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
        Ress2[ii]= [ Ress2[ii][0] ] + [ 'No Asociado' ] + [ ' - ' for i in range( len( Re2[ii] ) - 2 ) ]

   

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

                            

