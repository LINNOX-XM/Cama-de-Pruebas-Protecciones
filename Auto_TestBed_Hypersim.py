# -*- coding: latin-1 -*-
import os
import sys
sys.path.append(r'C:\OPAL-RT\HYPERSIM\hypersim_2021.1.2.o150\Windows\HyApi\C\py')
sys.path.append(r'C:\OPAL-RT\HYPERSIM\hypersim_2021.1.2.o150\Windows\ScopeView\lib')
from termcolor import colored
import HyWorksApi
import ScopeViewApi
import time
from datetime import datetime
#import numpy as np
#import matplotlib as pl
import xlwings as xlw
import csv

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

# Creaciï¿½n de diccionario para almacenar parametros de la prueba
Parampruebalin= {'caso':[],'tipo':[],'disporcent':[],'voltref':[],'tiempofalla':[],'tiempoclear':[],'distancia':[],'rfalla':[],'elemento':[],
                 'fault_loc':[],'RDef':[],'EnaT1':[],'EnaT2':[],'EnaT3':[],'EnaT4':[],'T1':[],'T2':[],'T3':[],'T4':[],'T1Pa':[],'T1Pb':[],'T1Pc':[],'T1Pg':[],
                 'T2Pa':[],'T2Pb':[],'T2Pc':[],'T2Pg':[],'T3Pa':[],'T3Pb':[],'T3Pc':[],'T3Pg':[],'T4Pa':[],'T4Pb':[],'T4Pc':[],'T4Pg':[],'t_inyect':[]}
Parampruebabar= {'caso':[],'tipo':[],'tiempofalla':[],'tiempoclear':[],'elemento':[],'rfalla':[],'EnaT1':[],'EnaT2':[],'T1':[],'T2':[],'T1Pa':[],
                 'T1Pb':[],'T1Pc':[],'T1Pg':[],'T2Pa':[],'T2Pb':[],'T2Pc':[],'T2Pg':[],'t_inyect':[]}

# Loop para llenar diccionario de parametros de pruebas sobre lï¿½neas
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

# print(Paramprueba)





# Arrancar Hs (Hypersim), abrir el caso base, carga archivo de sensores, abre SV (Scope View),
# carga template de mediciones, analiza, map task, compila caso de HS, corre caso base y realiza una adquisiciï¿½n para el caso base
HyWorksApi.startAndConnectHypersim()
#HyWorksApi.startHyperWorks(stdout=None, stderr=None)
#HyWorksApi.connectToHyWorks(host='localhost',timeout=180000)



#designPath = r'C:\Users\WORKSTATION03\Downloads\BancoDePruebasV38_2___V2021__Target2__Python\BancoDePruebasV38_2___V2021__Target2.ecf'
#hy_sensors = r'C:\Users\WORKSTATION03\Downloads\BancoDePruebasV38_2___V2021__Target2__Python\SensoresV38_2___V2021__Target2.sig'
#sv_template = r'C:\Users\WORKSTATION03\Downloads\BancoDePruebasV38_2___V2021__Target2__Python\Template_ScopeView__TestBed.svt'



designPath = r'D:\JoseM\Relé Siemens - Cama de Pruebas\BancoDePruebasV38_2___V2021__Target2\BancoDePruebasV38_2___V2021__Target2.ecf'
hy_sensors = r'D:\JoseM\Relé Siemens - Cama de Pruebas\BancoDePruebasV38_2___V2021__Target2\SensoresV38_2___V2021__Target2.sig'
sv_template = r'D:\JoseM\Relé Siemens - Cama de Pruebas\Template_ScopeView__TestBed.svt'



#print( 'DesignPath = ' , designPath )
#print( 'Hy_Sensors = ', hy_sensors )
#print( 'SV_Template = ', sv_template )


HyWorksApi.openDesign(designPath)
HyWorksApi.clearCodeDir()
HyWorksApi.loadSensors (hy_sensors)
HyWorksApi.analyze()
HyWorksApi.mapTask()
HyWorksApi.genCode()
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

Test_Time= {}


#ii= 0
# Loop para inyectar cada caso de prueba de lï¿½neas. Se leen los parï¿½metros de cada caso y asignan a las líneas, inyecta con SV y regresa al estado estacionario inicial
for caso in range(len(Parampruebalin['caso'])):
#for caso in range( 3 ):
    
    #ii += 1
    #caso= 7+ii

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
 
        

    # Para la simulaciï¿½n, quita y pone un nuevo POW
    # HyWorksApi.stopSim()
    # HyWorksApi.removeDevice('POW1')
    # HyWorksApi.addDevice('Network Tools.clf', 'Point-on-wave', 11200, -9700, 1)
    # HyWorksApi.connectDeviceToBus3ph('POW1','net_1','pownode')
    # HyWorksApi.startSim()
    #time.sleep(5)
    ScopeViewApi.setTrig(True)
    Parampruebalin['t_inyect'].append(str(datetime.now()))
    Test_Time[ 'Caso_' + str( int( Parampruebalin['caso'][caso] )).zfill(2) ]= datetime.now()
    ScopeViewApi.startAcquisition()
    #print colored('Tiempo de inyeccion = '+Parampruebalin['t_inyect'][-1], 'yellow')
    print('Tiempo de inyeccion = '+ Parampruebalin['t_inyect'][-1])
    # HyWorksApi.stopSim()
    # HyWorksApi.startSim()
#   HyWorksApi.takeSnapshot()
#   print colored('Snapshot de Caso base tomado', 'yellow')
#   print colored('Cargando snapshot del caso base', 'yellow')
    HyWorksApi.loadSnapshot()
    time.sleep(25)
    ScopeViewApi.setTrig(False)
    ScopeViewApi.startAcquisition()
    #print colored('Siguiente caso...', 'yellow')
    
    if Parampruebalin['caso'][caso] == 11 or Parampruebalin['caso'][caso] == 12: 
        HyWorksApi.setComponentParameter( 'CB2', 'EnaGen', '0' )
        HyWorksApi.setComponentParameter( 'CB1', 'EnaGen', '0' )
    
    print('Siguiente caso...')

print('Fin del set de pruebas en líneas')


HyWorksApi.setComponentParameter( 'CP1', 'EnaGen', '0' )                # Se apagan ambas lï¿½ï¿½neas para q no entren en las fallas de las barras
HyWorksApi.setComponentParameter( 'CP3', 'EnaGen', '0' )

# Loop para inyectar cada caso de prueba de barras. Se leen los parï¿½metros de cada caso y asignan a las barras, inyecta con SV y regresa al estado estacionario inicial
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

    # HyWorksApi.stopSim()
    # HyWorksApi.removeDevice('POW1')
    # HyWorksApi.addDevice('Network Tools.clf', 'Point-on-wave', 11200, -9700, 1)
    # HyWorksApi.connectDeviceToBus3ph('POW1', 'net_1', 'pownode')
    # HyWorksApi.startSim()
    ScopeViewApi.setTrig(True)
    Parampruebabar['t_inyect'].append( str(datetime.now() ))
    Test_Time[ 'Caso_' + str( int( Parampruebabar['caso'][casobar] )).zfill(2) ]=  datetime.now()
    #print colored('Tiempo de inyeccion = ' + Parampruebabar['t_inyect'][-1], 'yellow')
    ScopeViewApi.startAcquisition()
    # HyWorksApi.stopSim()
    # HyWorksApi.startSim()
    #   HyWorksApi.takeSnapshot()
    #   print colored('Snapshot de Caso base tomado', 'yellow')
    #   print colored('Cargando snapshot del caso base', 'yellow')
    HyWorksApi.loadSnapshot()
    time.sleep(25)
    ScopeViewApi.setTrig(False)
    ScopeViewApi.startAcquisition()
    #print colored('Siguiente caso...', 'yellow')

HyWorksApi.setComponentParameter( 'Flt2', 'EnaGen', '0' )
HyWorksApi.setComponentParameter( 'Flt3', 'EnaGen', '0' )

Test_Time11= Parampruebalin['t_inyect'][:3] + Parampruebabar['t_inyect'] + Parampruebalin['t_inyect'][3:]
#Test_Time22= Parampruebalin['t_inyect'] + Parampruebabar['t_inyect']


HyWorksApi.stopSim()
HyWorksApi.closeDesign(designPath)

print('Fin Pruebas')










