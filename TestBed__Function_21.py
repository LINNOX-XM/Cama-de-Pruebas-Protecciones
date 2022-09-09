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

"""
#designPath = r'D:\JoseM\Cama_de_Pruebas\BancoDePruebasV38_2___V2021__Target2\BancoDePruebasV38_2___V2021__Target2.ecf'
#hy_sensors = r'D:\JoseM\Cama_de_Pruebas\BancoDePruebasV38_2___V2021__Target2\SensoresV38_2___V2021__Target2.sig'
#sv_template = r'D:\JoseM\Cama_de_Pruebas\Template_ScopeView__TestBed.svt'
"""

#"""
designPath = r'D:\JoseM\Cama_de_Pruebas\BancoDePruebasV38_2___V2021__Target2__Rel_Remoto\BancoDePruebasV38_2___V2021__Target2__Rel_Remoto.ecf'
hy_sensors = r'D:\JoseM\Cama_de_Pruebas\BancoDePruebasV38_2___V2021__Target2__Rel_Remoto\SensoresV38_2___V2021__Target2__Rel_Remoto__Full.sig'
sv_template = r'D:\JoseM\Cama_de_Pruebas\Template_ScopeView__TestBed__Rel_Rem.svt'
#"""


# Carpeta y path para Comtrades
Carpeta= 'Ensayo Rel_Remoto_6_NN'                                                                                   # En esta carpeta se van a descargar los archivos Comtrade que se generan en el relé
pathh = 'D:\JoseM\Cama_de_Pruebas\Pruebas Fallas CID\Ensayos Automatismo Python'                                    # Path de la carpeta anterior
path2= pathh + "\\" + Carpeta

if not os.path.exists( path2 ):                                                                                     # Verifica que la carpeta anterior exista, si no, la crea
    os.makedirs( path2 )



# Inicio del Automatismo que ejecuta las pruebas en Hypersim


def search_xlw(col,string):
    for i in range(1, 1000):
        if string in str(sheet.range(col+'{}'.format(i)).value):
            filmatch = i
            break
    return filmatch

# Leer hoja de Excel con los parametros de la prueba
sheet = xlw.Book('DISTANCE_21_POTT_JM.xlsx').sheets('Input_API')

# Determina el rango de datos para las pruebas de lineas y de barras
linrowini = search_xlw('A','lineas')+2
linrowend = sheet.range('A'+str(linrowini)).end('down').row
barrowini = search_xlw('A','barras')+2
barrowend = sheet.range('A'+str(barrowini)).end('down').row

# Creación de diccionario para almacenar parametros de la prueba
Parampruebalin= {'caso':[],'tipo':[],'disporcent':[],'voltref':[],'tiempofalla':[],'tiempoclear':[],'distancia':[],'rfalla':[],'elemento':[],
                 'fault_loc':[],'RDef':[],'EnaT1':[],'EnaT2':[],'EnaT3':[],'EnaT4':[],'T1':[],'T2':[],'T3':[],'T4':[],'T1Pa':[],'T1Pb':[],'T1Pc':[],'T1Pg':[],
                 'T2Pa':[],'T2Pb':[],'T2Pc':[],'T2Pg':[],'T3Pa':[],'T3Pb':[],'T3Pc':[],'T3Pg':[],'T4Pa':[],'T4Pb':[],'T4Pc':[],'T4Pg':[],'t_inyect':[]}
Parampruebabar= {'caso':[],'tipo':[],'tiempofalla':[],'tiempoclear':[],'elemento':[],'rfalla':[],'EnaT1':[],'EnaT2':[],'T1':[],'T2':[],'T1Pa':[],
                 'T1Pb':[],'T1Pc':[],'T1Pg':[],'T2Pa':[],'T2Pb':[],'T2Pc':[],'T2Pg':[],'t_inyect':[]}

# Loop para llenar diccionario de parametros de pruebas sobre líneas
for row in range(linrowini,linrowend+1):
    Parampruebalin['caso'].append(sheet.range('A'+str(row)).value)
    Parampruebalin['tipo'].append(str(sheet.range('B'+str(row)).value))
    Parampruebalin['disporcent'].append(sheet.range('C'+str(row)).value)
    Parampruebalin['voltref'].append(str(sheet.range('D'+str(row)).value))
    Parampruebalin['tiempofalla'].append(sheet.range('E'+str(row)).value)
    Parampruebalin['tiempoclear'].append(sheet.range('F'+str(row)).value)
    Parampruebalin['distancia'].append(sheet.range('G'+str(row)).value)
    Parampruebalin['rfalla'].append(sheet.range('H'+str(row)).value)
    Parampruebalin['elemento'].append(str(sheet.range('I'+str(row)).value))
    Parampruebalin['fault_loc'].append(sheet.range('J'+str(row)).value)
    Parampruebalin['RDef'].append(sheet.range('K'+str(row)).value)
    Parampruebalin['EnaT1'].append(sheet.range('L' + str(row)).value)
    Parampruebalin['EnaT2'].append(sheet.range('M' + str(row)).value)
    Parampruebalin['EnaT3'].append(sheet.range('N' + str(row)).value)
    Parampruebalin['EnaT4'].append(sheet.range('O' + str(row)).value)
    Parampruebalin['T1'].append(sheet.range('P'+str(row)).value)
    Parampruebalin['T2'].append(sheet.range('Q'+str(row)).value)
    Parampruebalin['T3'].append(sheet.range('R'+str(row)).value)
    Parampruebalin['T4'].append(sheet.range('S'+str(row)).value)
    Parampruebalin['T1Pa'].append(int(sheet.range('T'+str(row)).value))
    Parampruebalin['T1Pb'].append(int(sheet.range('U'+str(row)).value))
    Parampruebalin['T1Pc'].append(int(sheet.range('V'+str(row)).value))
    Parampruebalin['T1Pg'].append(int(sheet.range('W'+str(row)).value))
    Parampruebalin['T2Pa'].append(int(sheet.range('X'+str(row)).value))
    Parampruebalin['T2Pb'].append(int(sheet.range('Y'+str(row)).value))
    Parampruebalin['T2Pc'].append(int(sheet.range('Z'+str(row)).value))
    Parampruebalin['T2Pg'].append(int(sheet.range('AA'+str(row)).value))
    Parampruebalin['T3Pa'].append(int(sheet.range('AB'+str(row)).value))
    Parampruebalin['T3Pb'].append(int(sheet.range('AC'+str(row)).value))
    Parampruebalin['T3Pc'].append(int(sheet.range('AD'+str(row)).value))
    Parampruebalin['T3Pg'].append(int(sheet.range('AE'+str(row)).value))
    Parampruebalin['T4Pa'].append(int(sheet.range('AF'+str(row)).value))
    Parampruebalin['T4Pb'].append(int(sheet.range('AG'+str(row)).value))
    Parampruebalin['T4Pc'].append(int(sheet.range('AH'+str(row)).value))
    Parampruebalin['T4Pg'].append(int(sheet.range('AI'+str(row)).value))

# Loop para llenar diccionario de parametros de pruebas sobre barras
for row in range(barrowini,barrowend+1):
    Parampruebabar['caso'].append(sheet.range('A'+str(row)).value)
    Parampruebabar['tipo'].append(str(sheet.range('B'+str(row)).value))
    Parampruebabar['tiempofalla'].append(sheet.range('C'+str(row)).value)
    Parampruebabar['tiempoclear'].append(sheet.range('D'+str(row)).value)
    Parampruebabar['elemento'].append(str(sheet.range('E'+str(row)).value))
    Parampruebabar['rfalla'].append(sheet.range('F'+str(row)).value)
    Parampruebabar['EnaT1'].append(sheet.range('G'+str(row)).value)
    Parampruebabar['EnaT2'].append(sheet.range('H'+str(row)).value)
    Parampruebabar['T1'].append(sheet.range('I'+str(row)).value)
    Parampruebabar['T2'].append(sheet.range('J'+str(row)).value)
    Parampruebabar['T1Pa'].append(int(sheet.range('K'+str(row)).value))
    Parampruebabar['T1Pb'].append(int(sheet.range('L'+str(row)).value))
    Parampruebabar['T1Pc'].append(int(sheet.range('M'+str(row)).value))
    Parampruebabar['T1Pg'].append(int(sheet.range('N'+str(row)).value))
    Parampruebabar['T2Pa'].append(int(sheet.range('O'+str(row)).value))
    Parampruebabar['T2Pb'].append(int(sheet.range('P'+str(row)).value))
    Parampruebabar['T2Pc'].append(int(sheet.range('Q'+str(row)).value))
    Parampruebabar['T2Pg'].append(int(sheet.range('R'+str(row)).value))



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

HyWorksApi.setComponentParameter( 'CB_Aux_Ant', 'CmdBlockSelect', 'Internal')                      # Se apagan estos interruptores y se dejan conectados siempre para q no cambien su estado durante los casos
HyWorksApi.setComponentParameter( 'CB_Aux_Ant', 'EnaGen', '0')
HyWorksApi.setComponentParameter( 'CB_Aux_Cerro', 'CmdBlockSelect', 'Internal')
HyWorksApi.setComponentParameter( 'CB_Aux_Cerro', 'EnaGen', '0')

HyWorksApi.setComponentParameter( 'CB_Aux_Ant', 'EtatIniA', '1')                       
HyWorksApi.setComponentParameter( 'CB_Aux_Ant', 'EtatIniB', '1') 
HyWorksApi.setComponentParameter( 'CB_Aux_Ant', 'EtatIniC', '1') 

HyWorksApi.setComponentParameter( 'CB_Aux_Cerro', 'EtatIniA', '1') 
HyWorksApi.setComponentParameter( 'CB_Aux_Cerro', 'EtatIniB', '1') 
HyWorksApi.setComponentParameter( 'CB_Aux_Cerro', 'EtatIniC', '1')

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

time.sleep(3)

Test_Time= {}


#ii= 0
# Loop para inyectar cada caso de prueba de líneas. Se leen los parámetros de cada caso y asignan a las líneas, inyecta con SV y regresa al estado estacionario inicial
for caso in range(len(Parampruebalin['caso'])):
#for caso in range( 3 ):
    
    #ii += 1
    #caso= 5+ii

    HyWorksApi.setComponentParameter( 'CP1', 'EnaGen', '1' )
    HyWorksApi.setComponentParameter( 'CP3', 'EnaGen', '0' )

    # Carga snapshot del caso base antes de inyectar el siguiente caso
    print('Ejecutando Caso %i' %( int(Parampruebalin['caso'][caso]) ) )
    para = 'Parametros--->'+'  Elemento-->'+str(Parampruebalin['elemento'][caso])+'  Tipo Falla-->'+str(Parampruebalin['tipo'][caso])+'  Distancia-->'+str(Parampruebalin['distancia'][caso])+'  T falla-->'+str(Parampruebalin['tiempofalla'][caso])+'  T Clear-->'+str(Parampruebalin['tiempoclear'][caso])+'  R falla-->'+str(Parampruebalin['rfalla'][caso])
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
    HyWorksApi.setComponentParameter(Parampruebalin['elemento'][caso],'T3Pa',str(Parampruebalin['T3Pa'][caso]))
    HyWorksApi.setComponentParameter(Parampruebalin['elemento'][caso],'T3Pb',str(Parampruebalin['T3Pb'][caso]))
    HyWorksApi.setComponentParameter(Parampruebalin['elemento'][caso],'T3Pc',str(Parampruebalin['T3Pc'][caso]))
    HyWorksApi.setComponentParameter(Parampruebalin['elemento'][caso],'T3Pg',str(Parampruebalin['T3Pg'][caso]))
    HyWorksApi.setComponentParameter(Parampruebalin['elemento'][caso],'T4Pa',str(Parampruebalin['T4Pa'][caso]))
    HyWorksApi.setComponentParameter(Parampruebalin['elemento'][caso],'T4Pb',str(Parampruebalin['T4Pb'][caso]))
    HyWorksApi.setComponentParameter(Parampruebalin['elemento'][caso],'T4Pc',str(Parampruebalin['T4Pc'][caso]))
    HyWorksApi.setComponentParameter(Parampruebalin['elemento'][caso],'T4Pg',str(Parampruebalin['T4Pg'][caso]))
    
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
    Test_Time[ 'Caso ' + str( int( Parampruebalin['caso'][caso] )) ]= datetime.now()
    ScopeViewApi.startAcquisition()
    print('Tiempo de inyeccion = '+ Parampruebalin['t_inyect'][-1])

    HyWorksApi.loadSnapshot()
    ScopeViewApi.setTrig(False)
    ScopeViewApi.startAcquisition()
    time.sleep(25)
    
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
    Test_Time[ 'Caso ' + str( int( Parampruebabar['caso'][casobar] )) ]=  datetime.now()
    ScopeViewApi.startAcquisition()
    
    # HyWorksApi.stopSim()                                                                                  # Esto se pone si se va a simular en el Target1
    # HyWorksApi.startSim()
    #   HyWorksApi.takeSnapshot()
    #   print colored('Snapshot de Caso base tomado', 'yellow')
    #   print colored('Cargando snapshot del caso base', 'yellow')
    HyWorksApi.loadSnapshot()
    ScopeViewApi.setTrig(False)
    ScopeViewApi.startAcquisition()
    time.sleep(25)

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
    
    if DD[k]['DO'] != var2 and k == Data2['Caso'][ii]:        
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

                            

"""
for caso in DD:
    print( caso )
    print( DD[caso] )
"""