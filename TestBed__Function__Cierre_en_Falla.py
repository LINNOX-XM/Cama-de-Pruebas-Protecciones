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
sv_template = r'D:\JoseM\Cama_de_Pruebas\Template_ScopeView__TestBed__Rel_Rem.svt'
#"""


# Carpeta y path para Comtrades
Carpeta= 'Ensayo CierreEnFalla_1'                                                                               # En esta carpeta se van a descargar los archivos Comtrade que se generan en el relé
pathh = 'D:\JoseM\Cama_de_Pruebas\Pruebas Fallas CID\Ensayos Automatismo Python'                                # Path de la carpeta anterior
path2= pathh + "\\" + Carpeta

if not os.path.exists( path2 ):                                                                                     # Verifica que la carpeta anterior exista, si no, la crea
    os.makedirs( path2 )





def CloseCBs(): 
    
    HyWorksApi.setComponentParameter( 'CB_Aux_Ant','EtatIniA', '1' )                                   # Abre
    HyWorksApi.setComponentParameter( 'CB_Aux_Ant','EtatIniB', '1' )
    HyWorksApi.setComponentParameter( 'CB_Aux_Ant','EtatIniC', '1' )
    
    HyWorksApi.setComponentParameter( 'CB_Aux_Cerro','EtatIniA', '1' )                                 # Abre
    HyWorksApi.setComponentParameter( 'CB_Aux_Cerro','EtatIniB', '1' )
    HyWorksApi.setComponentParameter( 'CB_Aux_Cerro','EtatIniC', '1' )
    
    
    


def Conf_CB_Aux( CB_Name, EnaT1, EnaT2, EnaT3, T1, T2, T3 ):
    
    HyWorksApi.setComponentParameter( str( CB_Name ), 'EnaT1', str( EnaT1 ) )
    HyWorksApi.setComponentParameter( str( CB_Name ), 'EnaT2', str( EnaT2 ) )
    HyWorksApi.setComponentParameter( str( CB_Name ), 'EnaT3', str( EnaT3 ) )
    HyWorksApi.setComponentParameter( str( CB_Name ), 'T1', str( T1 ) )
    HyWorksApi.setComponentParameter( str( CB_Name ), 'T2', str( T2 ) )
    HyWorksApi.setComponentParameter( str( CB_Name ), 'T3', str( T3 ) )

    HyWorksApi.setComponentParameter( str( CB_Name ), 'T1Pa', '1' )                                   # Abre
    HyWorksApi.setComponentParameter( str( CB_Name ), 'T1Pb', '1' )
    HyWorksApi.setComponentParameter( str( CB_Name ), 'T1Pc', '1' )
    
    HyWorksApi.setComponentParameter( str( CB_Name ), 'T2Pa', '1' )                                   # Abre
    HyWorksApi.setComponentParameter( str( CB_Name ), 'T2Pb', '1' )
    HyWorksApi.setComponentParameter( str( CB_Name ), 'T2Pc', '1' )
    
    HyWorksApi.setComponentParameter( str( CB_Name ), 'T3Pa', '1' )                                   # Abre
    HyWorksApi.setComponentParameter( str( CB_Name ), 'T3Pb', '1' )
    HyWorksApi.setComponentParameter( str( CB_Name ), 'T3Pc', '1' )
       




def Line_Fault_Param( Fault_loc, RDef, EnaT1, EnaT2, EnaT3, EnaT4, T1, T2, T3, T4, Fault_Type ):
    
    HyWorksApi.setComponentParameter( 'CP1','fault_loc', str( Fault_loc ) )
    HyWorksApi.setComponentParameter( 'CP1','RDef', str( RDef ) )
    HyWorksApi.setComponentParameter( 'CP1','EnaT1', str( EnaT1 ) )
    HyWorksApi.setComponentParameter( 'CP1','EnaT2', str( EnaT2 ) )
    HyWorksApi.setComponentParameter( 'CP1','EnaT3', str( EnaT3 ) )
    HyWorksApi.setComponentParameter( 'CP1','EnaT4', str( EnaT4 ) )
    HyWorksApi.setComponentParameter( 'CP1','T1', str( T1 ) )
    HyWorksApi.setComponentParameter( 'CP1','T2', str( T2 ) )                                               
    HyWorksApi.setComponentParameter( 'CP1','T3', str( T3 ) )
    HyWorksApi.setComponentParameter( 'CP1','T4', str( T4 ) )                                               
    
    
    
    if Fault_Type == 1:                                             # Falla B-G
            
        HyWorksApi.setComponentParameter( 'CP1','T1Pa', '0' )
        HyWorksApi.setComponentParameter( 'CP1','T1Pb', '1' )
        HyWorksApi.setComponentParameter( 'CP1','T1Pc', '0' )
        HyWorksApi.setComponentParameter( 'CP1','T1Pg', '1' )
        
        HyWorksApi.setComponentParameter( 'CP1','T2Pa', '0' )
        HyWorksApi.setComponentParameter( 'CP1','T2Pb', '1' )
        HyWorksApi.setComponentParameter( 'CP1','T2Pc', '0' )
        HyWorksApi.setComponentParameter( 'CP1','T2Pg', '1' )
        
        HyWorksApi.setComponentParameter( 'CP1','T3Pa', '0' )
        HyWorksApi.setComponentParameter( 'CP1','T3Pb', '1' )
        HyWorksApi.setComponentParameter( 'CP1','T3Pc', '0' )
        HyWorksApi.setComponentParameter( 'CP1','T3Pg', '1' )
        
        HyWorksApi.setComponentParameter( 'CP1','T4Pa', '0' )
        HyWorksApi.setComponentParameter( 'CP1','T4Pb', '1' )
        HyWorksApi.setComponentParameter( 'CP1','T4Pc', '0' )
        HyWorksApi.setComponentParameter( 'CP1','T4Pg', '1' )
        
    elif Fault_Type == 2:                                           # Falla AB-G
  
        HyWorksApi.setComponentParameter( 'CP1','T1Pa', '1' )
        HyWorksApi.setComponentParameter( 'CP1','T1Pb', '1' )
        HyWorksApi.setComponentParameter( 'CP1','T1Pc', '0' )
        HyWorksApi.setComponentParameter( 'CP1','T1Pg', '1' )
        
        HyWorksApi.setComponentParameter( 'CP1','T2Pa', '1' )
        HyWorksApi.setComponentParameter( 'CP1','T2Pb', '1' )
        HyWorksApi.setComponentParameter( 'CP1','T2Pc', '0' )
        HyWorksApi.setComponentParameter( 'CP1','T2Pg', '1' )
        
        HyWorksApi.setComponentParameter( 'CP1','T3Pa', '1' )
        HyWorksApi.setComponentParameter( 'CP1','T3Pb', '1' )
        HyWorksApi.setComponentParameter( 'CP1','T3Pc', '0' )
        HyWorksApi.setComponentParameter( 'CP1','T3Pg', '1' )
        
        HyWorksApi.setComponentParameter( 'CP1','T4Pa', '1' )
        HyWorksApi.setComponentParameter( 'CP1','T4Pb', '1' )
        HyWorksApi.setComponentParameter( 'CP1','T4Pc', '0' )
        HyWorksApi.setComponentParameter( 'CP1','T4Pg', '1' )

    elif Fault_Type == 3:                                           # Falla ABC-G

        HyWorksApi.setComponentParameter( 'CP1','T1Pa', '1' )
        HyWorksApi.setComponentParameter( 'CP1','T1Pb', '1' )
        HyWorksApi.setComponentParameter( 'CP1','T1Pc', '1' )
        HyWorksApi.setComponentParameter( 'CP1','T1Pg', '1' )
        
        HyWorksApi.setComponentParameter( 'CP1','T2Pa', '1' )
        HyWorksApi.setComponentParameter( 'CP1','T2Pb', '1' )
        HyWorksApi.setComponentParameter( 'CP1','T2Pc', '1' )
        HyWorksApi.setComponentParameter( 'CP1','T2Pg', '1' )
       
        HyWorksApi.setComponentParameter( 'CP1','T3Pa', '1' )
        HyWorksApi.setComponentParameter( 'CP1','T3Pb', '1' )
        HyWorksApi.setComponentParameter( 'CP1','T3Pc', '1' )
        HyWorksApi.setComponentParameter( 'CP1','T3Pg', '1' )
        
        HyWorksApi.setComponentParameter( 'CP1','T4Pa', '1' )
        HyWorksApi.setComponentParameter( 'CP1','T4Pb', '1' )
        HyWorksApi.setComponentParameter( 'CP1','T4Pc', '1' )
        HyWorksApi.setComponentParameter( 'CP1','T4Pg', '1' )





def ProcedEntreCasos():
    
    print('Stop Sim')
    HyWorksApi.stopSim() 
    HyWorksApi.clearCodeDir()
    time.sleep(5)
    print('Start Sim')
    HyWorksApi.startSim()
    time.sleep(5)    
    
    ScopeViewApi.setTrig(True)
    Test_Time.append( datetime.now() )
    
    ScopeViewApi.startAcquisition()
    print( 'Tiempo de inyeccion = ' + str( Test_Time[-1] ) )    
    
    #HyWorksApi.loadSnapshot()
    time.sleep(10)
    CloseCBs()
    ScopeViewApi.setTrig(False)
    ScopeViewApi.startAcquisition()
    time.sleep(25)





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

HyWorksApi.setComponentParameter( 'CB_Aux_Ant', 'EnaGen', '1')                                             
HyWorksApi.setComponentParameter( 'CB_Aux_Ant', 'CmdBlockSelect', 'Internal')                       
HyWorksApi.setComponentParameter( 'CB_Aux_Cerro', 'EnaGen', '1')
HyWorksApi.setComponentParameter( 'CB_Aux_Cerro', 'CmdBlockSelect', 'Internal')

HyWorksApi.setComponentParameter( 'CP1', 'EnaGen', '1')
HyWorksApi.setComponentParameter( 'CP3', 'EnaGen', '0')

HyWorksApi.startSim()
print('Simulacion en ejecución... Caso base')
HyWorksApi.takeSnapshot()
print('Snapshot de Caso base tomado')
time.sleep(5)

ScopeViewApi.openScopeView()
ScopeViewApi.loadTemplate(sv_template)
ScopeViewApi.setSync(True)
ScopeViewApi.setTrig(False)
ScopeViewApi.startAcquisition()
print('Adquisicion... Caso base')

CloseCBs()

time.sleep(3)


Test_Time= []







"""

# Caso 3A: Estado inicial: Línea abierta sin falla "pegada". Se simula una prefalla y luego de 1 segundo, se cierra el interruptor BUS. 
#  A los 200 ms se mete la falla para ver si dipara. Se espera durante 1 segundo de posfalla y finaliza el caso. Esta joda se repite pero con el TIE


print('Caso 3A: CB BUS, Falla B-G, 1 Ohm, 5%')


#"""
Conf_CB_Aux( 'CB_Aux_Ant', 1, 1, 0, 0, 3.5, 0)
Conf_CB_Aux( 'CB_Aux_Cerro', 1, 0, 0, 0, 0, 0)
Line_Fault_Param( 7, 1, 1, 1, 0, 0, 3.7, 4.7, 0, 0, 1)
#"""


ProcedEntreCasos()


print('Siguiente caso...' + '\n' )





print('Caso 3B: CB TIE, Falla B-G, 1 Ohm, 91%')


#"""
Conf_CB_Aux( 'CB_Aux_Cerro', 1, 1, 0, 0, 3.5, 0)
Conf_CB_Aux( 'CB_Aux_Ant', 1, 0, 0, 0, 0, 0)
Line_Fault_Param( 102, 1, 1, 1, 0, 0, 3.7, 4.7, 0, 0, 1)
#"""


ProcedEntreCasos()


print('Siguiente caso...' + '\n' )





print('Caso 3C: CB BUS, Falla ABC-G, 1 Ohm, 5%')


#"""
Conf_CB_Aux( 'CB_Aux_Ant', 1, 1, 0, 0, 3.5, 0)
Conf_CB_Aux( 'CB_Aux_Cerro', 1, 0, 0, 0, 0, 0)
Line_Fault_Param( 7, 1, 1, 1, 0, 0, 3.7, 4.7, 0, 0, 2)
#"""


ProcedEntreCasos()


print('Siguiente caso...' + '\n' )






"""
























# Caso 1A__BUS: Línea abierta con falla Monofásica "pegada" al 1% de la línea. Se simula una prefalla y luego de 1 segundo, se cierran el interruptor BUS 
#  para ver si la falla continúa. Se espera durante 1 segundo de posfalla y finaliza el caso. Esta joda se repite pero con el TIE

print('Caso 1A: CB BUS, Falla B-G, 1 Ohm, 5%' )

#"""
Conf_CB_Aux( 'CB_Aux_Ant', 1, 1, 0, 0, 3.5, 0)
Conf_CB_Aux( 'CB_Aux_Cerro', 1, 0, 0, 0, 0, 0)
Line_Fault_Param( 7, 1, 1, 1, 0, 0, 0.5, 4.5, 0, 0, 1)
#"""

"""
HyWorksApi.setComponentParameter( 'CP1','fault_loc', '7' )
HyWorksApi.setComponentParameter( 'CP1','RDef', '1' )
HyWorksApi.setComponentParameter( 'CP1','EnaT1', '1' )
HyWorksApi.setComponentParameter( 'CP1','EnaT2', '1' )
HyWorksApi.setComponentParameter( 'CP1','T1', '0.5' )
HyWorksApi.setComponentParameter( 'CP1','T2', '1.5' )
HyWorksApi.setComponentParameter( 'CP1','T1Pa', '0' )
HyWorksApi.setComponentParameter( 'CP1','T1Pb', '1' )
HyWorksApi.setComponentParameter( 'CP1','T1Pc', '0' )
HyWorksApi.setComponentParameter( 'CP1','T1Pg', '1' )
HyWorksApi.setComponentParameter( 'CP1','T2Pa', '0' )
HyWorksApi.setComponentParameter( 'CP1','T2Pb', '1' )
HyWorksApi.setComponentParameter( 'CP1','T2Pc', '0' )
HyWorksApi.setComponentParameter( 'CP1','T2Pg', '1' )


HyWorksApi.setComponentParameter( 'CB_Aux_Ant','EnaT1', '1' )
HyWorksApi.setComponentParameter( 'CB_Aux_Ant','EnaT2', '1' )
HyWorksApi.setComponentParameter( 'CB_Aux_Ant','T1', '0' )
HyWorksApi.setComponentParameter( 'CB_Aux_Ant','T2', '3.5' )
HyWorksApi.setComponentParameter( 'CB_Aux_Ant','T3', '5' )

HyWorksApi.setComponentParameter( 'CB_Aux_Ant','T1Pa', '1' )                                   # Abre
HyWorksApi.setComponentParameter( 'CB_Aux_Ant','T1Pb', '1' )
HyWorksApi.setComponentParameter( 'CB_Aux_Ant','T1Pc', '1' )

HyWorksApi.setComponentParameter( 'CB_Aux_Ant','T2Pa', '1' )                                   # Cierra
HyWorksApi.setComponentParameter( 'CB_Aux_Ant','T2Pb', '1' )
HyWorksApi.setComponentParameter( 'CB_Aux_Ant','T2Pc', '1' )

HyWorksApi.setComponentParameter( 'CB_Aux_Cerro','EnaT1', '1' )
HyWorksApi.setComponentParameter( 'CB_Aux_Cerro','T1', '0' )
HyWorksApi.setComponentParameter( 'CB_Aux_Cerro','T1Pa', '1' )                                 # Abre
HyWorksApi.setComponentParameter( 'CB_Aux_Cerro','T1Pb', '1' )
HyWorksApi.setComponentParameter( 'CB_Aux_Cerro','T1Pc', '1' )

"""


ScopeViewApi.setTrig(True)
Test_Time.append( datetime.now() )

ScopeViewApi.startAcquisition()
print( 'Tiempo de inyeccion = '+ str( Test_Time[-1] ) )

#HyWorksApi.loadSnapshot()
CloseCBs()
ScopeViewApi.setTrig(False)
ScopeViewApi.startAcquisition()
time.sleep(25)


print('Siguiente caso...' + '\n' )

"""
HyWorksApi.stopSim()                                                                                  # Esto se pone si se va a simular en el Target1
HyWorksApi.removeDevice('POW1')
HyWorksApi.addDevice('Network Tools.clf', 'Point-on-wave', 12300, -9900, 1)
HyWorksApi.connectDeviceToBus3ph('POW1', 'net_1', 's26')
HyWorksApi.startSim()
time.sleep(5)
"""

"""
HyWorksApi.stopSim() 
HyWorksApi.clearCodeDir()
time.sleep(5)
HyWorksApi.startSim()
time.sleep(5)
"""




# Caso 1B__TIE: Línea abierta con falla Monofásica "pegada" al 90 % de la línea. Se simula una prefalla y luego de 1 segundo, se cierra el interruptor TIE 
#  para ver si la falla continúa. Se espera durante 1 segundo de posfalla y finaliza el caso.

print('Caso 1B: CB TIE, Falla B-G, 1 Ohm, 90%' )

#"""
Conf_CB_Aux( 'CB_Aux_Cerro', 1, 1, 0, 0, 3.5, 0)
Conf_CB_Aux( 'CB_Aux_Ant', 1, 0, 0, 0, 0, 0)
Line_Fault_Param( 102, 1, 1, 1, 0, 0, 0.5, 4.5, 0, 0, 1)
#"""


ProcedEntreCasos()


print('Siguiente caso...' + '\n' )




# Caso 1C__BUS: Línea abierta con falla Bifásica "pegada" al  90% de la línea. Se simula una prefalla y luego de 1 segundo, se cierran el interruptor BUS
#  para ver si la falla continúa. Se espera durante 1 segundo de posfalla y finaliza el caso.


print('Caso 1C: CB BUS, Falla AC-G, 1 Ohm, 90%' )

#"""
Conf_CB_Aux( 'CB_Aux_Ant', 1, 1, 0, 0, 3.5, 0)
Conf_CB_Aux( 'CB_Aux_Cerro', 1, 0, 0, 0, 0, 0)
Line_Fault_Param( 102, 1, 1, 1, 0, 0, 0.5, 4.5, 0, 0, 2)
#"""


ProcedEntreCasos()


print('Siguiente caso...' + '\n' )





# Caso 1D__TIE: Línea abierta con falla Monofásica "pegada" al 50 % de la línea. Se simula una prefalla y luego de 1 segundo, se cierran el interruptor TIE 
#  para ver si la falla continúa. Se espera durante 1 segundo de posfalla y finaliza el caso.

print('Caso 1D: CB TIE, Falla B-G, 5 Ohm, 50%' )

#"""
Conf_CB_Aux( 'CB_Aux_Cerro', 1, 1, 0, 0, 3.5, 0)
Conf_CB_Aux( 'CB_Aux_Ant', 1, 0, 0, 0, 0, 0)
Line_Fault_Param( 56, 1, 1, 1, 0, 0, 0.5, 4.5, 0, 0, 1)
#"""


ProcedEntreCasos()


print('Siguiente caso...' + '\n' )





# Caso 1E__BUS: Línea abierta con falla Trifásica "pegada" al  5% de la línea. Se simula una prefalla y luego de 1 segundo, se cierran el interruptor BUS
#  para ver si la falla continúa. Se espera durante 1 segundo de posfalla y finaliza el caso.


print('Caso 1E: CB BUS, Falla ABC-G, 1 Ohm, 5%' )

#"""
Conf_CB_Aux( 'CB_Aux_Ant', 1, 1, 0, 0, 3.5, 0)
Conf_CB_Aux( 'CB_Aux_Cerro', 1, 0, 0, 0, 0, 0)
Line_Fault_Param( 7, 1, 1, 1, 0, 0, 0.5, 4.5, 0, 0, 3)
#"""


ProcedEntreCasos()


print('Siguiente caso...' + '\n' )





print('Fin Caso 1' + '\n')


time.sleep(5)








# Caso 2A: Estado inicial: Línea energizada sin falla. Se simula una falla y se queda la falla "pegada" por 1 segundo, 
#  se espera que se abran ambos extremos de la línea y se realice un recierre en ambos también. Como la falla sigue, 
#  la deben ver ambos extremos de la línea y operar al instante.

# Falla queda pegada 1 seg pa poder hacer recierre. 
# El CB_Aux_Cerro abre 50ms luego de la falla, pero debe cerrar en 500 ms luego de haber abierto el Aux_Cerro (550ms después de la falla)
# Cuando la falla sea Bipolar, el CB_Aux_Cerro tiene q cerrar a 750ms luego de la falla

#  También se debe hacer el otro subcaso, donde hay recierre sin falla y a los 200 ms se mete la falla. Es decir, Estado inicial: Línea energizada sin falla. 
#   Se simula una falla q dura 150 ms, se espera que se abran ambos extremos de la línea y se realice un recierre en ambos también. 
#   En ese momento del recierre, se espera 200 ms y se vuelve a hacer una falla de 150 ms que deben ver ambos extremos de la línea y operar al instante. 
#    El CB_Aux_Cerro va a operar a los 250 ms de la segunda falla

                          
print('Caso 2A: CB BUS, Falla B-G, 1 Ohm, 5%' )                                             # # Monofásica al "1%"


#"""
Conf_CB_Aux( 'CB_Aux_Ant', 0, 0, 0, 0, 0, 0)
Conf_CB_Aux( 'CB_Aux_Cerro', 1, 1, 0, 1.050, 1.550, 0)
Line_Fault_Param( 7, 1, 1, 1, 0, 0, 1, 2, 0, 0, 1)
#"""


ProcedEntreCasos()


print('Siguiente caso...' + '\n' )




                          
print('Caso 2B: CB BUS, Falla B-G, 1 Ohm, 90%' )    


#"""
Conf_CB_Aux( 'CB_Aux_Ant', 0, 0, 0, 0, 0, 0)
Conf_CB_Aux( 'CB_Aux_Cerro', 1, 1, 0, 1.050, 1.550, 0)
Line_Fault_Param( 95, 1, 1, 1, 0, 0, 1, 2, 0, 0, 1)
#"""


ProcedEntreCasos()


print('Siguiente caso...' + '\n' )




                          
print('Caso 2C: CB BUS, Falla AB-G, 1 Ohm, 5%' )                                             


#"""
Conf_CB_Aux( 'CB_Aux_Ant', 0, 0, 0, 0, 0, 0)
Conf_CB_Aux( 'CB_Aux_Cerro', 1, 1, 0, 1.050, 1.750, 0)
Line_Fault_Param( 7, 1, 1, 1, 0, 0, 1, 2, 0, 0, 2)
#"""


ProcedEntreCasos()


print('Siguiente caso...' + '\n' )






#  También se debe hacer el otro subcaso, donde hay recierre sin falla y a los 200 ms se mete la falla. 
#  Estado inicial: Línea energizada sin falla, se simula una falla q dura 150 ms, se espera que se abran ambos extremos de la línea y 
#   se realice un recierre de ambos también. En ese momento del recierre de Antioquia (Relé de prueba), se espera 200 ms y se vuelve a hacer una falla de 150 ms que 
#   deben ver ambos extremos de la línea y operar al instante. El CB_Aux_Cerro va a operar a los 250 ms de la segunda falla

                          
print('SubCaso 2A: CB BUS, Falla B-G, 1 Ohm, 5%' )                                             # Monofásica al "1%"


#"""
Conf_CB_Aux( 'CB_Aux_Ant', 0, 0, 0, 0, 0, 0)
Conf_CB_Aux( 'CB_Aux_Cerro', 0, 0, 0, 0, 0, 0)
Line_Fault_Param( 7, 1, 1, 1, 1, 1, 1, 1.150, 2.118, 2.268, 1)                                      # El recierre dura aprox 800 ms, por lo que la segunda falla se hace al segundo 2
#"""


ProcedEntreCasos()


print('Siguiente caso...' + '\n' )




                          
print('SubCaso 2B: CB BUS, Falla B-G, 1 Ohm, 90%' )


#"""
Conf_CB_Aux( 'CB_Aux_Ant', 0, 0, 0, 0, 0, 0)
Conf_CB_Aux( 'CB_Aux_Cerro', 0, 0, 0, 0, 0, 0)
Line_Fault_Param( 95, 1, 1, 1, 1, 1, 1, 1.150, 2.118, 2.268, 1)                                      # El recierre dura aprox 800 ms, por lo que la segunda falla se hace al segundo 2
#"""


ProcedEntreCasos()


print('Siguiente caso...' + '\n' )




                          
print('SubCaso 2C: CB BUS, Falla AB-G, 1 Ohm, 90%' )


#"""
Conf_CB_Aux( 'CB_Aux_Ant', 0, 0, 0, 0, 0, 0)
Conf_CB_Aux( 'CB_Aux_Cerro', 0, 0, 0, 0, 0, 0)
Line_Fault_Param( 95, 1, 1, 1, 1, 1, 1, 1.150, 2.118, 2.268, 2)                                      # El recierre dura aprox 800 ms, por lo que la segunda falla se hace al segundo 2
#"""


ProcedEntreCasos()


print('Siguiente caso...' + '\n' )


























"""

# Configurar la falla en la línea: Falla trifásica de 1 Ohm al 50% de la línea

HyWorksApi.setComponentParameter( 'CP1','fault_loc', '56' )
HyWorksApi.setComponentParameter( 'CP1','RDef', '1' )
HyWorksApi.setComponentParameter( 'CP1','EnaT1', '1' )
HyWorksApi.setComponentParameter( 'CP1','EnaT2', '1' )
HyWorksApi.setComponentParameter( 'CP1','T1', '1.2' )
HyWorksApi.setComponentParameter( 'CP1','T2', '1.3' )
HyWorksApi.setComponentParameter( 'CP1','T1Pa', '1' )
HyWorksApi.setComponentParameter( 'CP1','T1Pb', '1' )
HyWorksApi.setComponentParameter( 'CP1','T1Pc', '1' )
HyWorksApi.setComponentParameter( 'CP1','T1Pg', '0' )
HyWorksApi.setComponentParameter( 'CP1','T2Pa', '1' )
HyWorksApi.setComponentParameter( 'CP1','T2Pb', '1' )
HyWorksApi.setComponentParameter( 'CP1','T2Pc', '1' )
HyWorksApi.setComponentParameter( 'CP1','T2Pg', '0' )


# Apertura de CBs

HyWorksApi.setComponentParameter( 'CB_Recierre_Ant','EnaT3', '1' )
HyWorksApi.setComponentParameter( 'CB_Recierre_Ant','T1', '0' )
HyWorksApi.setComponentParameter( 'CB_Recierre_Ant','T2', '1' )
HyWorksApi.setComponentParameter( 'CB_Recierre_Ant','T3', '3' )

HyWorksApi.setComponentParameter( 'CB_Recierre_Ant','T1Pa', '1' )                               # Abre
HyWorksApi.setComponentParameter( 'CB_Recierre_Ant','T1Pb', '1' )
HyWorksApi.setComponentParameter( 'CB_Recierre_Ant','T1Pc', '1' )

HyWorksApi.setComponentParameter( 'CB_Recierre_Ant','T2Pa', '1' )                               # Cierra
HyWorksApi.setComponentParameter( 'CB_Recierre_Ant','T2Pb', '1' )
HyWorksApi.setComponentParameter( 'CB_Recierre_Ant','T2Pc', '1' )

HyWorksApi.setComponentParameter( 'CB_Recierre_Ant','T3Pa', '1' )                               # Abre
HyWorksApi.setComponentParameter( 'CB_Recierre_Ant','T3Pb', '1' )
HyWorksApi.setComponentParameter( 'CB_Recierre_Ant','T3Pc', '1' )


HyWorksApi.setComponentParameter( 'CB_Recierre_Cerro','T1', '0' )
HyWorksApi.setComponentParameter( 'CB_Recierre_Cerro','T2', '3' )

HyWorksApi.setComponentParameter( 'CB_Recierre_Cerro','T1Pa', '1' )                             # Abre
HyWorksApi.setComponentParameter( 'CB_Recierre_Cerro','T1Pb', '1' )
HyWorksApi.setComponentParameter( 'CB_Recierre_Cerro','T1Pc', '1' )






ScopeViewApi.setTrig(True)
Test_Time.append( datetime.now() )

ScopeViewApi.startAcquisition()
print( 'Tiempo de inyeccion = '+ str( Test_Time[-1] ) )

HyWorksApi.loadSnapshot()
ScopeViewApi.setTrig(False)
ScopeViewApi.startAcquisition()
time.sleep(25)


print('Siguiente caso...')

time.sleep(10)





# Caso 3B: Estado inicial: Línea abierta sin falla "pegada". Se simula una prefalla y luego de 1 segundo, se cierra el interruptor TIE. 
#  A los 200 ms se mete la falla para ver si dipara. Se espera durante 1 segundo de posfalla y finaliza el caso. 


print('Caso 3B: TIE')


# Configurar la falla en la línea: Falla trifásica de 1 Ohm al 50% de la línea

HyWorksApi.setComponentParameter( 'CP1','fault_loc', '56' )
HyWorksApi.setComponentParameter( 'CP1','RDef', '1' )
HyWorksApi.setComponentParameter( 'CP1','EnaT1', '1' )
HyWorksApi.setComponentParameter( 'CP1','EnaT2', '1' )
HyWorksApi.setComponentParameter( 'CP1','T1', '1.2' )
HyWorksApi.setComponentParameter( 'CP1','T2', '2.2' )
HyWorksApi.setComponentParameter( 'CP1','T1Pa', '1' )
HyWorksApi.setComponentParameter( 'CP1','T1Pb', '1' )
HyWorksApi.setComponentParameter( 'CP1','T1Pc', '1' )
HyWorksApi.setComponentParameter( 'CP1','T1Pg', '0' )
HyWorksApi.setComponentParameter( 'CP1','T2Pa', '1' )
HyWorksApi.setComponentParameter( 'CP1','T2Pb', '1' )
HyWorksApi.setComponentParameter( 'CP1','T2Pc', '1' )
HyWorksApi.setComponentParameter( 'CP1','T2Pg', '0' )


# Apertura de CBs

HyWorksApi.setComponentParameter( 'CB_Recierre_Cerro','EnaT3', '1' )
HyWorksApi.setComponentParameter( 'CB_Recierre_Cerro','T1', '0' )
HyWorksApi.setComponentParameter( 'CB_Recierre_Cerro','T2', '1' )
HyWorksApi.setComponentParameter( 'CB_Recierre_Cerro','T3', '3' )

HyWorksApi.setComponentParameter( 'CB_Recierre_Cerro','T1Pa', '1' )                                 # Abre
HyWorksApi.setComponentParameter( 'CB_Recierre_Cerro','T1Pb', '1' )
HyWorksApi.setComponentParameter( 'CB_Recierre_Cerro','T1Pc', '1' )

HyWorksApi.setComponentParameter( 'CB_Recierre_Cerro','T2Pa', '1' )                                 # Cierra
HyWorksApi.setComponentParameter( 'CB_Recierre_Cerro','T2Pb', '1' )
HyWorksApi.setComponentParameter( 'CB_Recierre_Cerro','T2Pc', '1' )

HyWorksApi.setComponentParameter( 'CB_Recierre_Cerro','T3Pa', '1' )                                 # Abre
HyWorksApi.setComponentParameter( 'CB_Recierre_Cerro','T3Pb', '1' )
HyWorksApi.setComponentParameter( 'CB_Recierre_Cerro','T3Pc', '1' )


HyWorksApi.setComponentParameter( 'CB_Recierre_Ant','T1', '0' )
HyWorksApi.setComponentParameter( 'CB_Recierre_Ant','T2', '3' )

HyWorksApi.setComponentParameter( 'CB_Recierre_Ant','T1Pa', '1' )                                   # Abre
HyWorksApi.setComponentParameter( 'CB_Recierre_Ant','T1Pb', '1' )
HyWorksApi.setComponentParameter( 'CB_Recierre_Ant','T1Pc', '1' )






ScopeViewApi.setTrig(True)
Test_Time.append( datetime.now() )

ScopeViewApi.startAcquisition()
print( 'Tiempo de inyeccion = '+ str( Test_Time[-1] ) )

HyWorksApi.loadSnapshot()
ScopeViewApi.setTrig(False)
ScopeViewApi.startAcquisition()
time.sleep(25)


print('Siguiente caso...')

time.sleep(10)












#ii= 0



HyWorksApi.stopSim()
ScopeViewApi.close()
HyWorksApi.closeDesign(designPath)
HyWorksApi.closeHyperWorks()


print( 'Test Time' ) 
print( Test_Time )

print('Fin Pruebas')

"""

