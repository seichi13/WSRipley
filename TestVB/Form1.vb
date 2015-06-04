Imports System
Imports Microsoft.VisualBasic
Imports TestVB.Log
Imports TestVB.Funciones
Imports TestVB.Tipos

Public Class Form1

    Public Shared c_TablaBloqueo As String(,)
    Private TIME_OUT_SERVER As Long = CLng(40) 'TIME OUT
    Private SERVIDOR_MIRROR_DESTINO As String = "BRVMSO" 'destination
    Private SERVIDOR_MIRROR_NODE As String = "BRVMSO" 'Mirror Node
    Private PUERTO As String = "2000" 'PUERTO
    Private PRIORIDAD_S As String = "02" 'PUERTO
    Private contadorArray As Integer = 0
    Private parameterOut As New VTAAUTO5001ParameterOUT
    'Private Tramas As String() = {"        AU            000VILLAVICENCIO VEGA, JUANA                         EL BOSQUE I                                                                                         ANCON                   LIMA            LIMA      54207-000105964-64                  13,600.00      0.00      0.00 13,600.0031/MAY/2015-30/JUN/201525/JUL/2015  3,050.00      0.00  3,050.15      0.00     84.72  3,050.15 10,550.00      0.00      0.00      0.00 10,550.00                                                              3,050.150000000000                                                                                                                                                                                                                0.00  3,050.00      0.00      0.00      0.00      0.00  3,050.15     84.72      0.00      0.00      0.00                                26/JUN/201526/JUN/2015004067COMPRA TURISMO      T    35.00  5.06%01/01                   35.0026/JUN/201526/JUN/2015004068COMPRA TURISMO      T    36.00  5.06%01/01                   36.0026/JUN/201526/JUN/2015004092COMPRA TURISMO      T    60.00  5.06%01/01                   60.0026/JUN/201526/JUN/2015004093COMPRA TURISMO      T    61.00  5.06%01/01                   61.0026/JUN/201526/JUN/2015004094COMPRA TURISMO      T    62.00  5.06%01/01                   62.0026/JUN/201526/JUN/2015004095COMPRA TURISMO      T    63.00  5.06%01/01                   63.0026/JUN/201526/JUN/2015004096COMPRA TURISMO      T    64.00  5.06%01/01                   64.0026/JUN/201526/JUN/2015004097COMPRA TURISMO      T    65.00  5.06%01/01                   65.0026/JUN/201526/JUN/2015004098COMPRA TURISMO      T    66.00  5.06%01/01                   66.0026/JUN/201526/JUN/2015004099COMPRA TURISMO      T    67.00  5.06%01/01                   67.0026/JUN/201526/JUN/2015004100COMPRA TURISMO      T    68.00  5.06%01/01                   68.0026/JUN/201526/JUN/2015004101COMPRA TURISMO      T    69.00  5.06%01/01                   69.00                                                           C",
    '                              "        AU            000VILLAVICENCIO VEGA, JUANA                         EL BOSQUE I                                                                                         ANCON                   LIMA            LIMA      54207-000105964-64                  13,600.00      0.00      0.00 13,600.0031/MAY/2015-30/JUN/201525/JUL/2015  3,050.00      0.00  3,050.15      0.00     84.72  3,050.15 10,550.00      0.00      0.00      0.00 10,550.00                                                              3,050.150000000000                                                                                                                                                                                                                0.00  3,050.00      0.00      0.00      0.00      0.00  3,050.15     84.72      0.00      0.00      0.00                                26/JUN/201526/JUN/2015004102COMPRA TURISMO      T    70.00  5.06%01/01                   70.00                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                     C",
    '                              "        AU            000VILLAVICENCIO VEGA, JUANA                         EL BOSQUE I                                                                                         ANCON                   LIMA            LIMA      54207-000105964-64                  13,600.00      0.00      0.00 13,600.0031/MAY/2015-30/JUN/201525/JUL/2015  3,050.00      0.00  3,050.15      0.00     84.72  3,050.15 10,550.00      0.00      0.00      0.00 10,550.00                                                              3,050.150000000000                                                                                                                                                                                                                0.00  3,050.00      0.00      0.00      0.00      0.00  3,050.15     84.72      0.00      0.00      0.00                                26/JUN/201526/JUN/2015004103COMPRA TURISMO      T    71.00  5.06%01/01                   71.0026/JUN/201526/JUN/2015004112COMPRA TURISMO      T    80.00  5.06%01/01                   80.0026/JUN/201526/JUN/2015004066COMPRA TURISMO      T    34.00  5.06%01/01                   34.0026/JUN/201526/JUN/2015004052COMPRA TURISMO      T    20.00  5.06%01/01                   20.0026/JUN/201526/JUN/2015004060COMPRA TURISMO      T    28.00  5.06%01/01                   28.0026/JUN/201526/JUN/2015004061COMPRA TURISMO      T    29.00  5.06%01/01                   29.0026/JUN/201526/JUN/2015004062COMPRA TURISMO      T    30.00  5.06%01/01                   30.0026/JUN/201526/JUN/2015004063COMPRA TURISMO      T    31.00  5.06%01/01                   31.0026/JUN/201526/JUN/2015004064COMPRA TURISMO      T    32.00  5.06%01/01                   32.0026/JUN/201526/JUN/2015004065COMPRA TURISMO      T    33.00  5.06%01/01                   33.0026/JUN/201526/JUN/2015004059COMPRA TURISMO      T    27.00  5.06%01/01                   27.0026/JUN/201526/JUN/2015004058COMPRA TURISMO      T    26.00  5.06%01/01                   26.00                                                           C",
    '                              "        AU            000VILLAVICENCIO VEGA, JUANA                         EL BOSQUE I                                                                                         ANCON                   LIMA            LIMA      54207-000105964-64                  13,600.00      0.00      0.00 13,600.0031/MAY/2015-30/JUN/201525/JUL/2015  3,050.00      0.00  3,050.15      0.00     84.72  3,050.15 10,550.00      0.00      0.00      0.00 10,550.00                                                              3,050.150000000000                                                                                                                                                                                                                0.00  3,050.00      0.00      0.00      0.00      0.00  3,050.15     84.72      0.00      0.00      0.00                                26/JUN/201526/JUN/2015004057COMPRA TURISMO      T    25.00  5.06%01/01                   25.00                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                     C",
    '                              "        AU            000VILLAVICENCIO VEGA, JUANA                         EL BOSQUE I                                                                                         ANCON                   LIMA            LIMA      54207-000105964-64                  13,600.00      0.00      0.00 13,600.0031/MAY/2015-30/JUN/201525/JUL/2015  3,050.00      0.00  3,050.15      0.00     84.72  3,050.15 10,550.00      0.00      0.00      0.00 10,550.00                                                              3,050.150000000000                                                                                                                                                                                                                0.00  3,050.00      0.00      0.00      0.00      0.00  3,050.15     84.72      0.00      0.00      0.00                                26/JUN/201526/JUN/2015004056COMPRA TURISMO      T    24.00  5.06%01/01                   24.0026/JUN/201526/JUN/2015004055COMPRA TURISMO      T    23.00  5.06%01/01                   23.0026/JUN/201526/JUN/2015004054COMPRA TURISMO      T    22.00  5.06%01/01                   22.0026/JUN/201526/JUN/2015004053COMPRA TURISMO      T    21.00  5.06%01/01                   21.0026/JUN/201526/JUN/2015004104COMPRA TURISMO      T    72.00  5.06%01/01                   72.0026/JUN/201526/JUN/2015004105COMPRA TURISMO      T    73.00  5.06%01/01                   73.0026/JUN/201526/JUN/2015004076COMPRA TURISMO      T    44.00  5.06%01/01                   44.0026/JUN/201526/JUN/2015004077COMPRA TURISMO      T    45.00  5.06%01/01                   45.0026/JUN/201526/JUN/2015004078COMPRA TURISMO      T    46.00  5.06%01/01                   46.0026/JUN/201526/JUN/2015004079COMPRA TURISMO      T    47.00  5.06%01/01                   47.0026/JUN/201526/JUN/2015004080COMPRA TURISMO      T    48.00  5.06%01/01                   48.0026/JUN/201526/JUN/2015004081COMPRA TURISMO      T    49.00  5.06%01/01                   49.00                                                           C",
    '                              "        AU            000VILLAVICENCIO VEGA, JUANA                         EL BOSQUE I                                                                                         ANCON                   LIMA            LIMA      54207-000105964-64                  13,600.00      0.00      0.00 13,600.0031/MAY/2015-30/JUN/201525/JUL/2015  3,050.00      0.00  3,050.15      0.00     84.72  3,050.15 10,550.00      0.00      0.00      0.00 10,550.00                                                              3,050.150000000000                                                                                                                                                                                                                0.00  3,050.00      0.00      0.00      0.00      0.00  3,050.15     84.72      0.00      0.00      0.00                                26/JUN/201526/JUN/2015004082COMPRA TURISMO      T    50.00  5.06%01/01                   50.00                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                     C",
    '                              "        AU            000VILLAVICENCIO VEGA, JUANA                         EL BOSQUE I                                                                                         ANCON                   LIMA            LIMA      54207-000105964-64                  13,600.00      0.00      0.00 13,600.0031/MAY/2015-30/JUN/201525/JUL/2015  3,050.00      0.00  3,050.15      0.00     84.72  3,050.15 10,550.00      0.00      0.00      0.00 10,550.00                                                              3,050.150000000000                                                                                                                                                                                                                0.00  3,050.00      0.00      0.00      0.00      0.00  3,050.15     84.72      0.00      0.00      0.00                                26/JUN/201526/JUN/2015004083COMPRA TURISMO      T    51.00  5.06%01/01                   51.00                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                     C",
    '                              "        AU            000VILLAVICENCIO VEGA, JUANA                         EL BOSQUE I                                                                                         ANCON                   LIMA            LIMA      54207-000105964-64                  13,600.00      0.00      0.00 13,600.0031/MAY/2015-30/JUN/201525/JUL/2015  3,050.00      0.00  3,050.15      0.00     84.72  3,050.15 10,550.00      0.00      0.00      0.00 10,550.00                                                              3,050.150000000000                                                                                                                                                                                                                0.00  3,050.00      0.00      0.00      0.00      0.00  3,050.15     84.72      0.00      0.00      0.00                                26/JUN/201526/JUN/2015004084COMPRA TURISMO      T    52.00  5.06%01/01                   52.0026/JUN/201526/JUN/2015004085COMPRA TURISMO      T    53.00  5.06%01/01                   53.0026/JUN/201526/JUN/2015004086COMPRA TURISMO      T    54.00  5.06%01/01                   54.0026/JUN/201526/JUN/2015004087COMPRA TURISMO      T    55.00  5.06%01/01                   55.0026/JUN/201526/JUN/2015004088COMPRA TURISMO      T    56.00  5.06%01/01                   56.0026/JUN/201526/JUN/2015004075COMPRA TURISMO      T    43.00  5.06%01/01                   43.0026/JUN/201526/JUN/2015004074COMPRA TURISMO      T    42.00  5.06%01/01                   42.0026/JUN/201526/JUN/2015004073COMPRA TURISMO      T    41.00  5.06%01/01                   41.0026/JUN/201526/JUN/2015004106COMPRA TURISMO      T    74.00  5.06%01/01                   74.0026/JUN/201526/JUN/2015004107COMPRA TURISMO      T    75.00  5.06%01/01                   75.0026/JUN/201526/JUN/2015004108COMPRA TURISMO      T    76.00  5.06%01/01                   76.0026/JUN/201526/JUN/2015004109COMPRA TURISMO      T    77.00  5.06%01/01                   77.00                                                           C",
    '                              "        AU            000VILLAVICENCIO VEGA, JUANA                         EL BOSQUE I                                                                                         ANCON                   LIMA            LIMA      54207-000105964-64                  13,600.00      0.00      0.00 13,600.0031/MAY/2015-30/JUN/201525/JUL/2015  3,050.00      0.00  3,050.15      0.00     84.72  3,050.15 10,550.00      0.00      0.00      0.00 10,550.00                                                              3,050.150000000000                                                                                                                                                                                                                0.00  3,050.00      0.00      0.00      0.00      0.00  3,050.15     84.72      0.00      0.00      0.00                                26/JUN/201526/JUN/2015004110COMPRA TURISMO      T    78.00  5.06%01/01                   78.00                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                     C",
    '                              "        AU            000VILLAVICENCIO VEGA, JUANA                         EL BOSQUE I                                                                                         ANCON                   LIMA            LIMA      54207-000105964-64                  13,600.00      0.00      0.00 13,600.0031/MAY/2015-30/JUN/201525/JUL/2015  3,050.00      0.00  3,050.15      0.00     84.72  3,050.15 10,550.00      0.00      0.00      0.00 10,550.00                                                              3,050.150000000000                                                                                                                                                                                                                0.00  3,050.00      0.00      0.00      0.00      0.00  3,050.15     84.72      0.00      0.00      0.00                                26/JUN/201526/JUN/2015004111COMPRA TURISMO      T    79.00  5.06%01/01                   79.0026/JUN/201526/JUN/2015004091COMPRA TURISMO      T    59.00  5.06%01/01                   59.0026/JUN/201526/JUN/2015004090COMPRA TURISMO      T    58.00  5.06%01/01                   58.0026/JUN/201526/JUN/2015004072COMPRA TURISMO      T    40.00  5.06%01/01                   40.0026/JUN/201526/JUN/2015004071COMPRA TURISMO      T    39.00  5.06%01/01                   39.0026/JUN/201526/JUN/2015004070COMPRA TURISMO      T    38.00  5.06%01/01                   38.0026/JUN/201526/JUN/2015004069COMPRA TURISMO      T    37.00  5.06%01/01                   37.0026/JUN/201526/JUN/2015004089COMPRA TURISMO      T    57.00  5.06%01/01                   57.00                                                                                                                                                                                                                                                                                                                                                                                                                                                   U"}
    'Private Tramas As String() = {"        AU            000MORENO HIDALGO, JUANA                             PLAYA 1                                                                                             ANCON                   LIMA            LIMA      54207-000143514-52                  15,000.00      0.00      0.00 15,000.0001/MAY/2015-30/JUN/201525/JUL/2015 12,299.12    181.19 12,480.93  2,790.44  4,241.09  4,241.09      0.00      0.00      0.00      0.00      0.00                                                              4,241.090000000000                                                                                                                                                                                                            2,790.44      0.00    484.38    867.06     99.00      0.00  4,241.09      0.00    484.38    867.06     99.00                                                            SALDO INICIAL                                              2790.4407/ABR/201507/ABR/2015      COMPRA EN CUOTAS    T  9014.11210.43%03/13  484.38 863.66  1348.04                            INTERES PERIODO FACT                                          0.60                            INTERES COMPENSATORI                                          2.80                            PENALIDAD PAGO VCDO.                                         99.00                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                             U"}
    'Private Tramas As String() = {"        AU            000BAZAN BARRIOS, ERICKA                             CLUB 16                                                                                             CIENEGUILLA             LIMA            LIMA      96041-000475113-80                  16,000.00      0.00      0.00 16,000.0001/MAY/2015-30/JUN/201525/JUL/2015 21,069.64    198.00 21,268.70  5,353.63  8,098.13 10,348.66      0.00      0.00      0.00      0.00      0.00                                                             10,348.660000000000                                                                                                                                                                                                            7,668.34      0.00  1,454.51  1,126.29     99.00      0.00 10,348.66      0.00  1,454.51  1,126.29     99.00                                                            SALDO INICIAL                                              7668.3422/MAR/201522/MAR/2015000005REPROGRAMADO CUOTAS T 14500.43193.90%04/10 1254.031086.18  2340.2106/ABR/201506/ABR/2015004247COMPRA EN CUOTAS    T  1200.00 18.72%03/06  200.48  11.80   212.28                            INTERES PERIODO FACT                                          6.95                            INTERES COMPENSATORI                                         21.36                            PENALIDAD PAGO VCDO.                                         99.00                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                               U"}
    'Private Tramas As String() = {"        AU            000                                                                                                                                                                                                        54207-000139778-93                  15,000.00      0.00      0.00 15,000.0008/JUL/2015-07/AGO/201501/SET/2015 21,891.03    297.00 22,189.14      0.00  9,993.26  9,993.26      0.00      0.00      0.00      0.00      0.00  8,053.32      0.00    774.04  1,165.40      0.00      0.00  9,993.26  8,053.32      0.00    774.04  1,165.40      0.00  9,993.26                                                OCT/2015  1,778.26NOV/2015  1,778.26DIC/2015  1,965.85ENE/2016  2,746.09FEB/2016  1,906.64MAR/2016  1,906.64ABR/2016  1,906.64MAY/2016  1,778.26JUN/2016  1,778.26JUL/2016  1,778.26                                                                                              SALDO INICIAL                                                                                                               8053.3227/MAR/201527/MAR/2015004230COMPRA EN CUOTAS                                                                        1400.00 18.72%05/12  114.05  14.33   128.3827/MAR/201527/MAR/2015000031ROGRAMADO CUOTAS                                                                       14000.35181.26%05/15  659.991118.27  1778.26                            INTERES PERIODO FACTURADO                                                                                                      3.82                            INTERES COMPENSATORIO                                                                                                         28.98                                                                                                                                                                                                                                                                                                                                                                                                        U"}
    'Private Tramas As String() = {"        AU            000                                                                                                                                                                                                        54207-000127184-39                  15,000.00      0.00      0.00 15,000.0008/JUL/2015-07/AGO/201501/SET/2015 22,710.41    297.00 23,008.56      0.00 12,077.82 12,077.82      0.00      0.00      0.00      0.00      0.00  9,567.10      0.00  1,019.30  1,490.82      0.00      0.00 12,077.82  9,567.10      0.00  1,019.30  1,490.82      0.00 12,077.82                                                OCT/2015  2,256.08NOV/2015  2,256.08DIC/2015  2,256.08ENE/2016  2,256.08FEB/2016  2,256.08MAR/2016  2,256.08ABR/2016  2,256.08                                                                                                                                                    SALDO INICIAL                                                                                                               9567.1027/MAR/201527/MAR/2015000030ROGRAMADO CUOTAS                                                                       15000.20213.84%05/12 1019.301236.78  2256.08                            INTERES PERIODO FACTURADO                                                                                                      4.58                            INTERES COMPENSATORIO                                                                                                        249.46                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                       U"}

    'Nuevas tramas (04706/2015)
    'Private Tramas As String() = {"        AU            000BARRABAS NOVATO, FELICIANA                        LA CURVA                                                                                            CHORRILLOS              LIMA            LIMA      52543-500223179-87                  17,000.00      0.00      0.00 17,000.0001/SET/2015-30/SET/201525/OCT/2015 -3,437.92      0.00 -3,437.92      0.00-11,062.45-11,062.45      0.00      0.00      0.00      0.00      0.00                                                            -11,062.45                                                                                                                                                                                                                                                                                                                                                                      7    2.23   15.50                            SALDO INICIAL                                                                                                               6259.4722/SET/201522/SET/2015004524PAGO BANCO RIPLEY                                                                    T                                    -17308.9422/SET/201522/SET/2015      RET.ITF PAGO                                                                         T                                         0.8522/SET/201522/SET/2015004525PAGO BANCO RIPLEY                                                                    T                                       -40.00                            INTERES COMPENSATORIO                                                                                                         27.02                                                                                                                                                                                                                                                                                                                                                                                  U"}
    Private Tramas As String() = {"        AU            000ROMERO VIRU DE LUCAS, NATALIA                     LA CURVA                                                                                            CHORRILLOS              LIMA            LIMA      52543-500223179-87                  17,000.00      0.00      0.00 17,000.0001/SET/2015-30/SET/201525/OCT/2015 -3,437.92      0.00 -3,437.92      0.00-11,062.45-11,062.45      0.00      0.00      0.00      0.00      0.00  6,259.47      0.00      0.00     27.02      0.00 17,348.94-11,062.45-11,089.47      0.00      0.00     27.02      0.00-11,062.45                                                NOV/2015  1,481.74DIC/2015  1,481.74ENE/2016  1,481.74FEB/2016  1,481.74MAR/2016  1,481.74                                                                                                                                                                5    1.21   12.20                            SALDO INICIAL                                                                                                               6259.4722/SET/201522/SET/2015004524PAGO BANCO RIPLEY                                                                    T                                    -17308.9422/MAR/201522/MAR/2015000035REPROGRAMADO CUOTAS                                                                  T 10000.00213.84%07/18  403.47 843.31  1246.7823/MAR/201523/MAR/2015000035REPROGRAMADO CUOTAS                                                                  T 10000.00213.84%08/10  402.47 841.31  1244.7822/SET/201522/SET/2015      RET.ITF PAGO                                                                         T                                         0.8522/SET/201522/SET/2015004525PAGO BANCO RIPLEY                                                                    T                                       -40.0022/SET/201522/SET/2015004525PAGO BANCO RIPLEY                                                                    T                                       -42.00                                                    C",
                                  "        AU            000ROMERO VIRU DE LUCAS, NATALIA                     LA CURVA                                                                                            CHORRILLOS              LIMA            LIMA      52543-500223179-87                  17,000.00      0.00      0.00 17,000.0001/SET/2015-30/SET/201525/OCT/2015 -3,437.92      0.00 -3,437.92      0.00-11,062.45-11,062.45      0.00      0.00      0.00      0.00      0.00  6,259.47      0.00      0.00     27.02      0.00 17,348.94-11,062.45-11,089.47      0.00      0.00     27.02      0.00-11,062.45                                                NOV/2015  1,481.74DIC/2015  1,481.74ENE/2016  1,481.74FEB/2016  1,481.74MAR/2016  1,481.74                                                                                                                                                                5    1.21   12.2022/SET/201522/SET/2015004524PAGO BANCO RIPLEY                                                                    T                                    -17308.9422/SET/201522/SET/2015004524PAGO BANCO RIPLEY                                                                    T                                    -17308.9422/MAR/201522/MAR/2015000035REPROGRAMADO CUOTAS                                                                  T 10000.00213.84%07/18  403.47 843.31  1246.7823/MAR/201523/MAR/2015000035REPROGRAMADO CUOTAS                                                                  T 10000.00213.84%08/10  402.47 841.31  1244.7822/SET/201522/SET/2015      RET.ITF PAGO                                                                         T                                         0.8522/SET/201522/SET/2015004525PAGO BANCO RIPLEY                                                                    T                                       -50.00                                                                                                                                                                                                                   C",
                                  "        AU            000ROMERO VIRU DE LUCAS, NATALIA                     LA CURVA                                                                                            CHORRILLOS              LIMA            LIMA      52543-500223179-87                  17,000.00      0.00      0.00 17,000.0001/SET/2015-30/SET/201525/OCT/2015 -3,437.92      0.00 -3,437.92      0.00-11,062.45-11,062.45      0.00      0.00      0.00      0.00      0.00  6,259.47      0.00      0.00     27.02      0.00 17,348.94-11,062.45-11,089.47      0.00      0.00     27.02      0.00-11,062.45                                                NOV/2015  1,481.74DIC/2015  1,481.74ENE/2016  1,481.74FEB/2016  1,481.74MAR/2016  1,481.74                                                                                                                                                                5    1.21   12.20                            SALDO INICIAL                                                                                                               6259.4722/SET/201522/SET/2015004524PAGO BANCO RIPLEY                                                                    T                                    -17308.9422/MAR/201522/MAR/2015000035REPROGRAMADO CUOTAS                                                                  T 10000.00213.84%07/18  403.47 843.31  1246.7823/MAR/201523/MAR/2015000035REPROGRAMADO CUOTAS                                                                  T 10000.00213.84%08/10  402.47 841.31  1244.7822/SET/201522/SET/2015      RET.ITF PAGO                                                                         T                                         0.8522/SET/201522/SET/2015004525PAGO BANCO RIPLEY                                                                    T                                       -40.0022/SET/201522/SET/2015004525PAGO BANCO RIPLEY                                                                    T                                       -49.00                                                    C",
                                  "        AU            000ROMERO VIRU DE LUCAS, NATALIA                     LA CURVA                                                                                            CHORRILLOS              LIMA            LIMA      52543-500223179-87                  17,000.00      0.00      0.00 17,000.0001/SET/2015-30/SET/201525/OCT/2015 -3,437.92      0.00 -3,437.92      0.00-11,062.45-11,062.45      0.00      0.00      0.00      0.00      0.00  6,259.47      0.00      0.00     27.02      0.00 17,348.94-11,062.45-11,089.47      0.00      0.00     27.02      0.00-11,062.45                                                NOV/2015  1,481.74DIC/2015  1,481.74ENE/2016  1,481.74FEB/2016  1,481.74MAR/2016  1,481.74                                                                                                                                                                5    1.21   12.2022/SET/201522/SET/2015004524PAGO BANCO RIPLEY                                                                    T                                    -17308.9422/SET/201522/SET/2015004524PAGO BANCO RIPLEY                                                                    T                                    -17308.9422/MAR/201522/MAR/2015000035REPROGRAMADO CUOTAS                                                                  T 10000.00213.84%07/18  403.47 843.31  1246.7823/MAR/201523/MAR/2015000035REPROGRAMADO CUOTAS                                                                  T 10000.00213.84%08/10  402.47 841.31  1244.7822/SET/201522/SET/2015      RET.ITF PAGO                                                                         T                                         0.8522/SET/201522/SET/2015004525PAGO BANCO RIPLEY                                                                    T                                       -50.00                                                                                                                                                                                                                   C",
                                  "        AU            000ROMERO VIRU DE LUCAS, NATALIA                     LA CURVA                                                                                            CHORRILLOS              LIMA            LIMA      52543-500223179-87                  17,000.00      0.00      0.00 17,000.0001/SET/2015-30/SET/201525/OCT/2015 -3,437.92      0.00 -3,437.92      0.00-11,062.45-11,062.45      0.00      0.00      0.00      0.00      0.00  6,259.47      0.00      0.00     27.02      0.00 17,348.94-11,062.45-11,089.47      0.00      0.00     27.02      0.00-11,062.45                                                NOV/2015  1,481.74DIC/2015  1,481.74ENE/2016  1,481.74FEB/2016  1,481.74MAR/2016  1,481.74                                                                                                                                                                5    1.21   12.20                            SALDO INICIAL                                                                                                               6259.4722/SET/201522/SET/2015004524PAGO BANCO RIPLEY                                                                    T                                    -17308.9422/MAR/201522/MAR/2015000035REPROGRAMADO CUOTAS                                                                  T 10000.00213.84%07/18  403.47 843.31  1246.7823/MAR/201523/MAR/2015000035REPROGRAMADO CUOTAS                                                                  T 10000.00213.84%08/10  402.47 841.31  1244.7822/SET/201522/SET/2015      RET.ITF PAGO                                                                         T                                         0.8522/SET/201522/SET/2015004525PAGO BANCO RIPLEY                                                                    T                                       -40.0022/SET/201522/SET/2015004525PAGO BANCO RIPLEY                                                                    T                                       -41.00                                                    C",
                                  "        AU            000ROMERO VIRU DE LUCAS, NATALIA                     LA CURVA                                                                                            CHORRILLOS              LIMA            LIMA      52543-500223179-87                  17,000.00      0.00      0.00 17,000.0001/SET/2015-30/SET/201525/OCT/2015 -3,437.92      0.00 -3,437.92      0.00-11,062.45-11,062.45      0.00      0.00      0.00      0.00      0.00  6,259.47      0.00      0.00     27.02      0.00 17,348.94-11,062.45-11,089.47      0.00      0.00     27.02      0.00-11,062.45                                                NOV/2015  1,481.74DIC/2015  1,481.74ENE/2016  1,481.74FEB/2016  1,481.74MAR/2016  1,481.74                                                                                                                                                                5    1.21   12.2022/SET/201522/SET/2015004524PAGO BANCO RIPLEY                                                                    T                                    -17308.9422/SET/201522/SET/2015004524PAGO BANCO RIPLEY                                                                    T                                    -17308.9422/MAR/201522/MAR/2015000035REPROGRAMADO CUOTAS                                                                  T 10000.00213.84%07/18  403.47 843.31  1246.7823/MAR/201523/MAR/2015000035REPROGRAMADO CUOTAS                                                                  T 10000.00213.84%08/10  402.47 841.31  1244.7822/SET/201522/SET/2015      RET.ITF PAGO                                                                         T                                         0.8522/SET/201522/SET/2015004525PAGO BANCO RIPLEY                                                                    T                                       -50.00                                                                                                                                                                                                                   C",
                                  "        AU            000ROMERO VIRU DE LUCAS, NATALIA                     LA CURVA                                                                                            CHORRILLOS              LIMA            LIMA      52543-500223179-87                  17,000.00      0.00      0.00 17,000.0001/SET/2015-30/SET/201525/OCT/2015 -3,437.92      0.00 -3,437.92      0.00-11,062.45-11,062.45      0.00      0.00      0.00      0.00      0.00  6,259.47      0.00      0.00     27.02      0.00 17,348.94-11,062.45-11,089.47      0.00      0.00     27.02      0.00-11,062.45                                                NOV/2015  1,481.74DIC/2015  1,481.74ENE/2016  1,481.74FEB/2016  1,481.74MAR/2016  1,481.74                                                                                                                                                                5    1.21   12.20                            SALDO INICIAL                                                                                                               6259.4722/SET/201522/SET/2015004524PAGO BANCO RIPLEY                                                                    T                                    -17308.9422/MAR/201522/MAR/2015000035REPROGRAMADO CUOTAS                                                                  T 10000.00213.84%07/18  403.47 843.31  1246.7823/MAR/201523/MAR/2015000035REPROGRAMADO CUOTAS                                                                  T 10000.00213.84%08/10  402.47 841.31  1244.7822/SET/201522/SET/2015      RET.ITF PAGO                                                                         T                                         0.8522/SET/201522/SET/2015004525PAGO BANCO RIPLEY                                                                    T                                       -40.0022/SET/201522/SET/2015004525PAGO BANCO RIPLEY                                                                    T                                       -48.00                                                    C",
                                  "        AU            000ROMERO VIRU DE LUCAS, NATALIA                     LA CURVA                                                                                            CHORRILLOS              LIMA            LIMA      52543-500223179-87                  17,000.00      0.00      0.00 17,000.0001/SET/2015-30/SET/201525/OCT/2015 -3,437.92      0.00 -3,437.92      0.00-11,062.45-11,062.45      0.00      0.00      0.00      0.00      0.00  6,259.47      0.00      0.00     27.02      0.00 17,348.94-11,062.45-11,089.47      0.00      0.00     27.02      0.00-11,062.45                                                NOV/2015  1,481.74DIC/2015  1,481.74ENE/2016  1,481.74FEB/2016  1,481.74MAR/2016  1,481.74                                                                                                                                                                5    1.21   12.20                            SALDO INICIAL                                                                                                               6259.4722/SET/201522/SET/2015004524PAGO BANCO RIPLEY                                                                    T                                    -17308.9422/MAR/201522/MAR/2015000035REPROGRAMADO CUOTAS                                                                  T 10000.00213.84%07/18  403.47 843.31  1246.7823/MAR/201523/MAR/2015000035REPROGRAMADO CUOTAS                                                                  T 10000.00213.84%08/10  402.47 841.31  1244.7822/SET/201522/SET/2015      RET.ITF PAGO                                                                         T                                         0.8522/SET/201522/SET/2015004525PAGO BANCO RIPLEY                                                                    T                                       -40.0022/SET/201522/SET/2015004525PAGO BANCO RIPLEY                                                                    T                                       -48.00                                                    U"}

    Public Function ESTADO_CUENTA_CLASICA_RSAT(ByVal sNroTarjeta As String, ByVal sNroCuenta As String, ByVal sDATA_MONITOR_KIOSCO As String, ByVal Servidor As TServidor) As String
        Dim sRespuesta As String = ""
        Dim sParametros As String = ""
        Dim sMensajeErrorUsuario As String = ""
        Dim sXML As String = ""
        Dim lContador As Long = 0
        Dim sPagina As String = ""

        'Variables nuevas
        Dim actions As Long = 1
        Dim inetputBuff As String = ""
        Dim outpputBuff As String = ""
        Dim errorMsg As String = ""

        Try
            sNroCuenta = Microsoft.VisualBasic.Strings.Right(sNroCuenta.Trim, 8)

            If sNroCuenta.Trim.Length = 8 And sNroTarjeta.Length = 16 Then

                'Instancia al mirapiweb
                '**Dim obSendMirror As ClsTxMirapi = Nothing
                '**obSendMirror = New ClsTxMirapi()

                Dim sTrama As String = ""
                Dim sXMLMOV As String = ""
                Dim sXMLMOV_FINAL As String = ""
                Dim sXMLCAB As String = ""
                Dim sXMLPIE As String = ""
                Dim sDataMov As String = ""
                Dim sDataMovAux As String = ""

                Dim lFila As Long = 0
                Dim Incrementa As Long = 1
                Dim tamanioFilaDetalle As Long = 94
                Dim tamanioFilaDetalleAnt As Long = 87

                'VARIABLES RESUMEN DE ESTADO DE CUENTA
                Dim sLineaCredito As String = ""
                Dim sLineaCreditoTemporal As String = ""
                Dim sLineaSuperEfectivo As String = ""
                Dim sLineaCreditoTotal As String = ""
                Dim sPeriodoFacturacion As String = ""
                Dim sUltimoDiaPago As String = ""
                Dim sCreditoUtilizado As String = ""
                Dim sComisionCargos As String = ""
                Dim sDeudaTotal As String = ""
                Dim sDeudaVencida As String = ""
                Dim sPagoMinimoMes As String = ""
                Dim sPagoTotalMes As String = ""
                Dim sDispCompras24Cuotas As String = ""
                Dim sDispCompras36Cuotas As String = ""
                Dim sDispSuperEfectivo As String = ""
                Dim sDisponibleCompras As String = ""
                Dim sDispEfectivoExpress As String = ""

                'CALCULO DE PAGO TOTAL DEL MES
                Dim sSaldoAnterior As String = ""
                Dim sConsumoMes As String = ""
                Dim sCuotaMes As String = ""
                Dim sInteres As String = ""
                Dim sComisionCargos_Mes As String = ""
                Dim sPagosAbonos As String = ""
                Dim sPagoTotal_Mes As String = ""

                'CALCULO DE PAGO MINIMO DEL MES
                Dim sSaldoFavor As String = ""
                Dim sMontoMinimo As String = ""
                Dim sCuotaMes_Min As String = ""
                Dim sInteres_Min As String = ""
                Dim sComisionCargos_Min As String = ""
                Dim sPagoMinimoMes_Min As String = ""


                'PLAN DE CUOTAS DE LOS PROXIMOS 3 MESES
                Dim sMes1 As String = ""
                Dim sMonto1 As String = ""
                Dim sMes2 As String = ""
                Dim sMonto2 As String = ""
                Dim sMes3 As String = ""
                Dim sMonto3 As String = ""

                'Cambios en Trama 04/06/2015
                Dim sCntMeses As String = ""
                Dim sIntPagMin As String = ""
                Dim sComPagMin As String = ""

                Dim sPeriodoFinal As String = ""
                Dim sPeriodo1 As String = ""
                Dim sPeriodo2 As String = ""
                Dim sPeriodo3 As String = ""
                Dim vFechaHoy As Date
                Dim fFechaHoyAux1 As Date
                Dim lEstadoLlamadaEECC As Long = 9 'Variable para iniciar llamadas al EECC 3 veces


                vFechaHoy = Date.Now
                fFechaHoyAux1 = vFechaHoy

                fFechaHoyAux1 = DateAdd("m", 1, fFechaHoyAux1)

                sPeriodo1 = Microsoft.VisualBasic.Strings.Right(Trim("0000" & Year(fFechaHoyAux1).ToString), 4) & Microsoft.VisualBasic.Strings.Right(Trim("00" & Month(fFechaHoyAux1).ToString), 2)

                sPeriodo2 = Microsoft.VisualBasic.Strings.Right(Trim("0000" & Year(vFechaHoy).ToString), 4) & Microsoft.VisualBasic.Strings.Right(Trim("00" & Month(vFechaHoy).ToString), 2)

                vFechaHoy = DateAdd("m", -1, vFechaHoy)

                sPeriodo3 = Microsoft.VisualBasic.Strings.Right(Trim("0000" & Year(vFechaHoy).ToString), 4) & Microsoft.VisualBasic.Strings.Right(Trim("00" & Month(vFechaHoy).ToString), 2)

                sPeriodoFinal = sPeriodo1.Trim
                'Agregar 201503 para que funcione EECC
                'sPeriodoFinal = ConfigurationSettings.AppSettings.Get("PeriodoInclusionTEA")

                Do
                    lContador = lContador + 1

                    If lContador > 9 Then
                        sPagina = lContador.ToString
                    Else
                        sPagina = "0" & lContador.ToString
                    End If

                    Dim CServidor As String = "0"

                    Select Case Servidor
                        Case TServidor.SICRON
                            CServidor = "0"
                        Case TServidor.RSAT
                            CServidor = "S"
                    End Select

                    sTrama = GetTramaR192(CServidor, sNroTarjeta, sNroCuenta, sPeriodoFinal, sPagina, outpputBuff)

                    If sTrama = "0" Then 'EXITO
                        If outpputBuff.Length > 0 Then sTrama = outpputBuff.Substring(8, outpputBuff.Length - 8)

                        'Intentos de call EECC
                        If lEstadoLlamadaEECC = 9 Then
                            lEstadoLlamadaEECC = 0 'Para que no vuelva a ingresar a esta logica

                            If Microsoft.VisualBasic.Strings.Left(sTrama, 2) = "RU" Then
                                'Segunda LLamada segundo periodo
                                sPeriodoFinal = sPeriodo2.Trim
                                sTrama = GetTramaR192(CServidor, sNroTarjeta, sNroCuenta, sPeriodoFinal, sPagina, outpputBuff)

                                If sTrama = "0" Then 'EXITO
                                    If outpputBuff.Length > 0 Then sTrama = outpputBuff.Substring(8, outpputBuff.Length - 8)

                                ElseIf sTrama = "-2" Then 'Ocurrio un error Recuperar el error
                                    'sTrama = "ERROR:" & errorMsg.Trim
                                    sTrama = ""
                                    Exit Do
                                Else
                                    'sTrama = "ERROR:Servicio no disponible."
                                    sTrama = ""
                                    Exit Do
                                End If


                            End If

                            If Microsoft.VisualBasic.Strings.Left(sTrama, 2) = "RU" Then
                                'Tercera LLamada Tercer periodo
                                sPeriodoFinal = sPeriodo3.Trim
                                sTrama = GetTramaR192(CServidor, sNroTarjeta, sNroCuenta, sPeriodoFinal, sPagina, outpputBuff)

                                If sTrama = "0" Then 'EXITO
                                    If outpputBuff.Length > 0 Then sTrama = outpputBuff.Substring(8, outpputBuff.Length - 8)

                                ElseIf sTrama = "-2" Then 'Ocurrio un error Recuperar el error
                                    'sTrama = "ERROR:" & errorMsg.Trim
                                    sTrama = ""
                                    Exit Do
                                Else
                                    'sTrama = "ERROR:Servicio no disponible."
                                    sTrama = ""
                                    Exit Do
                                End If

                            End If

                        End If 'fin de los intentos en periodos diferentes EECC

                    ElseIf sTrama = "-2" Then 'Ocurrio un error en la consulta de la transaccion
                        'sTrama = "ERROR:" & errorMsg.Trim
                        sTrama = ""
                        Exit Do
                    Else
                        'sTrama = "ERROR:Servicio no disponible."
                        sTrama = ""
                        Exit Do
                    End If


                    If Microsoft.VisualBasic.Strings.Left(sTrama, 2) <> "RU" Then
                        'CONSTRUIR CADENA DE LA CABECERA Y PIE DE PAGINA
                        If lContador = 1 Then

                            'VARIABLES RESUMEN DE ESTADO DE CUENTA
                            sLineaCredito = Mid(sTrama, 253, 10)
                            sPeriodoFacturacion = Mid(sTrama, 293, 23)

                            sUltimoDiaPago = Mid(sTrama, 316, 11)

                            sCreditoUtilizado = Mid(sTrama, 327, 10)
                            sComisionCargos = Mid(sTrama, 337, 10)
                            sDeudaTotal = Mid(sTrama, 347, 10)
                            sDeudaVencida = Mid(sTrama, 357, 10)
                            sPagoMinimoMes = Mid(sTrama, 367, 10)
                            sPagoTotalMes = Mid(sTrama, 377, 10)

                            sDisponibleCompras = Mid(sTrama, 387, 10)

                            sDispEfectivoExpress = Mid(sTrama, 427, 10)

                            'CALCULO DE PAGO TOTAL DEL MES
                            sSaldoAnterior = Mid(sTrama, 437, 10)
                            sConsumoMes = Mid(sTrama, 447, 10)
                            sCuotaMes = Mid(sTrama, 457, 10)
                            sInteres = Mid(sTrama, 467, 10)
                            sComisionCargos_Mes = Mid(sTrama, 477, 10)
                            sPagosAbonos = Mid(sTrama, 487, 10)
                            sPagoTotal_Mes = Mid(sTrama, 497, 10)

                            'CALCULO DE PAGO MINIMO DEL MES
                            sSaldoFavor = Mid(sTrama, 507, 10)
                            sMontoMinimo = Mid(sTrama, 517, 10)
                            sCuotaMes_Min = Mid(sTrama, 527, 10)
                            sInteres_Min = Mid(sTrama, 537, 10)
                            sComisionCargos_Min = Mid(sTrama, 547, 10)
                            sPagoMinimoMes_Min = Mid(sTrama, 557, 10)


                            'PLAN DE CUOTAS DE LOS PROXIMOS 3 MESES
                            sMes1 = Mid(sTrama, 615, 8)
                            sMonto1 = Mid(sTrama, 623, 10)
                            sMes2 = Mid(sTrama, 633, 8)
                            sMonto2 = Mid(sTrama, 641, 10)
                            sMes3 = Mid(sTrama, 651, 8)
                            sMonto3 = Mid(sTrama, 659, 10)

                            'Cambios en Trama 04/06/2015
                            sCntMeses = Mid(sTrama, 861, 5)
                            sIntPagMin = Mid(sTrama, 866, 8)
                            sComPagMin = Mid(sTrama, 874, 8)

                            'MOVIMIENTOS
                            sDataMovAux = ""
                            sDataMov = ""
                            sDataMov = Mid(sTrama, 882, sTrama.Length)
                            ErrorLog("Movimiento sDataMov = " & sDataMov)

                            Incrementa = 1
                            If sDataMov.Length > 0 Then

                                If sPeriodoFinal >= Constantes.PeriodoInclusionTEA Then
                                    '10-03-2015 Ocultar para pase por partes
                                    If sPeriodoFinal >= Constantes.PeriodoAgrandarGlosa Then
                                        tamanioFilaDetalle = 159

                                        If CServidor = Constantes.ServidorRSAT Then
                                            For lFila = 1 To 7

                                                If Trim(Mid(sDataMov, Incrementa, tamanioFilaDetalle)) <> "" Then
                                                    sDataMovAux = sDataMovAux & Mid(sDataMov, Incrementa, tamanioFilaDetalle) & "|\n|"
                                                    ErrorLog("sDataMovAux = " & sDataMovAux)
                                                End If

                                                Incrementa = Incrementa + tamanioFilaDetalle

                                            Next

                                            If Microsoft.VisualBasic.Strings.Right(sTrama, 1) = "C" Then
                                                ErrorLog("Llamar a TramaSiguiente")
                                                sTrama = GetTramaR192(Constantes.TramaTrece, sNroTarjeta, sNroCuenta, sPeriodoFinal, sPagina, outpputBuff)

                                                If sTrama = "0" Then 'EXITO
                                                    If outpputBuff.Length > 0 Then sTrama = outpputBuff.Substring(8, outpputBuff.Length - 8)

                                                ElseIf sTrama = "-2" Then
                                                    sTrama = ""
                                                    Exit Do
                                                Else
                                                    sTrama = ""
                                                    Exit Do
                                                End If

                                                If Microsoft.VisualBasic.Strings.Left(sTrama, 2) <> "RU" Then
                                                    sDataMov = ""
                                                    sDataMov = Mid(sTrama, 882, sTrama.Length)
                                                    ErrorLog("Movimiento sDataMov = " & sDataMov)
                                                    Incrementa = 1
                                                    If sDataMov.Length > 0 Then
                                                        For lFila = 1 To 6

                                                            If Trim(Mid(sDataMov, Incrementa, tamanioFilaDetalle)) <> "" Then
                                                                sDataMovAux = sDataMovAux & Mid(sDataMov, Incrementa, tamanioFilaDetalle) & "|\n|"
                                                                ErrorLog("sDataMovAux = " & sDataMovAux)
                                                            End If

                                                            Incrementa = Incrementa + tamanioFilaDetalle

                                                        Next
                                                    End If
                                                End If
                                            End If
                                        Else
                                            For lFila = 1 To 13

                                                If Trim(Mid(sDataMov, Incrementa, tamanioFilaDetalle)) <> "" Then
                                                    sDataMovAux = sDataMovAux & Mid(sDataMov, Incrementa, tamanioFilaDetalle) & "|\n|"
                                                    ErrorLog("sDataMovAux = " & sDataMovAux)
                                                End If

                                                Incrementa = Incrementa + tamanioFilaDetalle

                                            Next
                                        End If

                                    Else
                                        If CServidor = Constantes.ServidorRSAT Then
                                            For lFila = 1 To 12

                                                If Trim(Mid(sDataMov, Incrementa, tamanioFilaDetalle)) <> "" Then
                                                    sDataMovAux = sDataMovAux & Mid(sDataMov, Incrementa, tamanioFilaDetalle) & "|\n|"
                                                    ErrorLog("sDataMovAux = " & sDataMovAux)
                                                End If

                                                Incrementa = Incrementa + tamanioFilaDetalle

                                            Next

                                            If Microsoft.VisualBasic.Strings.Right(sTrama, 1) = "C" Then
                                                ErrorLog("Llamar a TramaSiguiente")
                                                sTrama = GetTramaR192(Constantes.TramaTrece, sNroTarjeta, sNroCuenta, sPeriodoFinal, sPagina, outpputBuff)

                                                If sTrama = "0" Then 'EXITO
                                                    If outpputBuff.Length > 0 Then sTrama = outpputBuff.Substring(8, outpputBuff.Length - 8)

                                                ElseIf sTrama = "-2" Then
                                                    sTrama = ""
                                                    Exit Do
                                                Else
                                                    sTrama = ""
                                                    Exit Do
                                                End If

                                                If Microsoft.VisualBasic.Strings.Left(sTrama, 2) <> "RU" Then
                                                    sDataMov = ""
                                                    sDataMov = Mid(sTrama, 861, sTrama.Length)
                                                    ErrorLog("Movimiento sDataMov = " & sDataMov)
                                                    If sDataMov.Length > 0 Then
                                                        If Trim(Mid(sDataMov, 1, tamanioFilaDetalle)) <> "" Then
                                                            sDataMovAux = sDataMovAux & Mid(sDataMov, 1, tamanioFilaDetalle) & "|\n|"
                                                            ErrorLog("sDataMovAux = " & sDataMovAux)
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        Else
                                            For lFila = 1 To 13

                                                If Trim(Mid(sDataMov, Incrementa, tamanioFilaDetalle)) <> "" Then
                                                    sDataMovAux = sDataMovAux & Mid(sDataMov, Incrementa, tamanioFilaDetalle) & "|\n|"
                                                    ErrorLog("sDataMovAux = " & sDataMovAux)
                                                End If

                                                Incrementa = Incrementa + tamanioFilaDetalle

                                            Next
                                        End If
                                    End If

                                Else
                                    For lFila = 1 To 13

                                        If Trim(Mid(sDataMov, Incrementa, tamanioFilaDetalleAnt)) <> "" Then
                                            sDataMovAux = sDataMovAux & Mid(sDataMov, Incrementa, tamanioFilaDetalleAnt) & "|\n|"
                                            ErrorLog("sDataMovAux = " & sDataMovAux)
                                        End If

                                        Incrementa = Incrementa + tamanioFilaDetalleAnt

                                    Next
                                End If

                            End If

                            sXMLCAB = ""
                            sXMLCAB = sLineaCredito.Trim & "|\t|" & sDispEfectivoExpress.Trim & "|\t|" & sDisponibleCompras.Trim
                            sXMLCAB = sXMLCAB & "|\t|" & sPeriodoFacturacion.Trim & "|\t|" & sUltimoDiaPago.Trim
                            sXMLCAB = sXMLCAB & "|\t|" & sCreditoUtilizado.Trim & "|\t|" & sComisionCargos.Trim & "|\t|" & sDeudaTotal.Trim
                            sXMLCAB = sXMLCAB & "|\t|" & sDeudaVencida.Trim & "|\t|" & sPagoMinimoMes.Trim & "|\t|" & sPagoTotalMes.Trim
                            ErrorLog("sXMLCAB=" & sXMLCAB)

                            sXMLPIE = ""
                            sXMLPIE = sSaldoAnterior.Trim & "|\t|" & sConsumoMes.Trim & "|\t|" & sCuotaMes.Trim
                            sXMLPIE = sXMLPIE & "|\t|" & sInteres.Trim & "|\t|" & sComisionCargos_Mes.Trim & "|\t|" & sPagosAbonos.Trim
                            sXMLPIE = sXMLPIE & "|\t|" & sPagoTotal_Mes.Trim & "|\t|" & sSaldoFavor.Trim & "|\t|" & sMontoMinimo.Trim
                            sXMLPIE = sXMLPIE & "|\t|" & sCuotaMes_Min.Trim & "|\t|" & sInteres_Min.Trim & "|\t|" & sComisionCargos_Min.Trim
                            sXMLPIE = sXMLPIE & "|\t|" & sPagoMinimoMes_Min.Trim & "|\t|" & sMes1.Trim & "|\t|" & sMonto1.Trim
                            sXMLPIE = sXMLPIE & "|\t|" & sMes2.Trim & "|\t|" & sMonto2.Trim & "|\t|" & sMes3.Trim & "|\t|" & sMonto3.Trim
                            ErrorLog("sXMLPIE=" & sXMLPIE)


                        Else
                            'EVALUAR LA SEGUNDA CALL SOLO MOVIMIENTOS CONTADOR DE LLAMADAS

                            'MOVIMIENTOS
                            sDataMovAux = ""
                            sDataMov = ""
                            sDataMov = Mid(sTrama, 882, sTrama.Length)
                            ErrorLog("Movimiento sDataMov = " & sDataMov)

                            Incrementa = 1
                            If sDataMov.Length > 0 Then

                                If sPeriodoFinal >= Constantes.PeriodoInclusionTEA Then
                                    '10-03-2015 Ocultar para pase por partes
                                    If sPeriodoFinal >= Constantes.PeriodoAgrandarGlosa Then
                                        tamanioFilaDetalle = 159

                                        If CServidor = Constantes.ServidorRSAT Then
                                            For lFila = 1 To 7

                                                If Trim(Mid(sDataMov, Incrementa, tamanioFilaDetalle)) <> "" Then
                                                    sDataMovAux = sDataMovAux & Mid(sDataMov, Incrementa, tamanioFilaDetalle) & "|\n|"
                                                    ErrorLog("sDataMovAux = " & sDataMovAux)
                                                End If

                                                Incrementa = Incrementa + tamanioFilaDetalle

                                            Next

                                            If lContador Mod 4 <> 0 Then
                                                If Microsoft.VisualBasic.Strings.Right(sTrama, 1) = "C" Then
                                                    ErrorLog("Llamar a TramaSiguiente")
                                                    sTrama = GetTramaR192(Constantes.TramaTrece, sNroTarjeta, sNroCuenta, sPeriodoFinal, sPagina, outpputBuff)

                                                    If sTrama = "0" Then 'EXITO
                                                        If outpputBuff.Length > 0 Then sTrama = outpputBuff.Substring(8, outpputBuff.Length - 8)

                                                    ElseIf sTrama = "-2" Then
                                                        sTrama = ""
                                                        Exit Do
                                                    Else
                                                        sTrama = ""
                                                        Exit Do
                                                    End If

                                                    If Microsoft.VisualBasic.Strings.Left(sTrama, 2) <> "RU" Then
                                                        sDataMov = ""
                                                        sDataMov = Mid(sTrama, 882, sTrama.Length)
                                                        ErrorLog("Movimiento sDataMov = " & sDataMov)
                                                        Incrementa = 1

                                                        If sDataMov.Length > 0 Then

                                                            For lFila = 1 To 6

                                                                If Trim(Mid(sDataMov, Incrementa, tamanioFilaDetalle)) <> "" Then
                                                                    sDataMovAux = sDataMovAux & Mid(sDataMov, Incrementa, tamanioFilaDetalle) & "|\n|"
                                                                    ErrorLog("sDataMovAux = " & sDataMovAux)
                                                                End If

                                                                Incrementa = Incrementa + tamanioFilaDetalle

                                                            Next

                                                        End If
                                                    End If
                                                End If
                                            End If

                                        Else
                                            For lFila = 1 To 13

                                                If Trim(Mid(sDataMov, Incrementa, tamanioFilaDetalle)) <> "" Then
                                                    sDataMovAux = sDataMovAux & Mid(sDataMov, Incrementa, tamanioFilaDetalle) & "|\n|"
                                                    ErrorLog("sDataMovAux = " & sDataMovAux)
                                                End If

                                                Incrementa = Incrementa + tamanioFilaDetalle

                                            Next
                                        End If

                                    Else

                                        If CServidor = Constantes.ServidorRSAT Then
                                            For lFila = 1 To 12

                                                If Trim(Mid(sDataMov, Incrementa, tamanioFilaDetalle)) <> "" Then
                                                    sDataMovAux = sDataMovAux & Mid(sDataMov, Incrementa, tamanioFilaDetalle) & "|\n|"
                                                    ErrorLog("sDataMovAux = " & sDataMovAux)
                                                End If

                                                Incrementa = Incrementa + tamanioFilaDetalle

                                            Next

                                            If lContador Mod 4 <> 0 Then
                                                If Microsoft.VisualBasic.Strings.Right(sTrama, 1) = "C" Then
                                                    ErrorLog("Llamar a TramaSiguiente")
                                                    sTrama = GetTramaR192(Constantes.TramaTrece, sNroTarjeta, sNroCuenta, sPeriodoFinal, sPagina, outpputBuff)

                                                    If sTrama = "0" Then 'EXITO
                                                        If outpputBuff.Length > 0 Then sTrama = outpputBuff.Substring(8, outpputBuff.Length - 8)

                                                    ElseIf sTrama = "-2" Then
                                                        sTrama = ""
                                                        Exit Do
                                                    Else
                                                        sTrama = ""
                                                        Exit Do
                                                    End If

                                                    If Microsoft.VisualBasic.Strings.Left(sTrama, 2) <> "RU" Then
                                                        sDataMov = ""
                                                        sDataMov = Mid(sTrama, 861, sTrama.Length)
                                                        ErrorLog("Movimiento sDataMov = " & sDataMov)
                                                        If sDataMov.Length > 0 Then
                                                            If Trim(Mid(sDataMov, 1, tamanioFilaDetalle)) <> "" Then
                                                                sDataMovAux = sDataMovAux & Mid(sDataMov, 1, tamanioFilaDetalle) & "|\n|"
                                                                ErrorLog("sDataMovAux = " & sDataMovAux)
                                                            End If
                                                        End If
                                                    End If
                                                End If
                                            End If

                                        Else
                                            For lFila = 1 To 13

                                                If Trim(Mid(sDataMov, Incrementa, tamanioFilaDetalle)) <> "" Then
                                                    sDataMovAux = sDataMovAux & Mid(sDataMov, Incrementa, tamanioFilaDetalle) & "|\n|"
                                                    ErrorLog("sDataMovAux = " & sDataMovAux)
                                                End If

                                                Incrementa = Incrementa + tamanioFilaDetalle

                                            Next
                                        End If

                                    End If

                                Else
                                    For lFila = 1 To 13

                                        If Trim(Mid(sDataMov, Incrementa, tamanioFilaDetalleAnt)) <> "" Then
                                            sDataMovAux = sDataMovAux & Mid(sDataMov, Incrementa, tamanioFilaDetalleAnt) & "|\n|"
                                            ErrorLog("sDataMovAux = " & sDataMovAux)
                                        End If

                                        Incrementa = Incrementa + tamanioFilaDetalleAnt

                                    Next
                                End If

                            End If



                        End If 'FIN DE VALIDACION DEL CONTADOR
                        If Microsoft.VisualBasic.Strings.Right(sTrama, 1) = "C" Then
                            'If Left(sTrama, 2) = "AC" Then 'HAY MOVIMIENTOS PENDIENTES POR LLAMAR

                            sXMLMOV = sXMLMOV & sDataMovAux
                            ErrorLog("sXMLMOV = " & sXMLMOV)

                        End If


                        If Microsoft.VisualBasic.Strings.Right(sTrama, 1) = "U" Then
                            'If Left(sTrama, 2) = "AU" Then 'FIN DE LA SOLICITUD

                            If sDataMovAux.Length > 0 Then
                                sDataMovAux = Microsoft.VisualBasic.Strings.Left(sDataMovAux, sDataMovAux.Length - 4)
                                sXMLMOV = sXMLMOV & sDataMovAux
                                ErrorLog("sXMLMOV = " & sXMLMOV)
                            Else
                                sDataMovAux = "" 'NO HAY MOVIMIENTOS
                                sXMLMOV = sXMLMOV & sDataMovAux
                                ErrorLog("sXMLMOV = " & sXMLMOV)

                            End If

                            If sXMLMOV.Length > 0 Then
                                'Armar las columnas de los movimientos de EECC
                                sXMLMOV_FINAL = FUN_DETALLE_EECC_TEA(sXMLMOV, sPeriodoFinal)
                                ErrorLog("sXMLMOV_FINAL = " & sXMLMOV_FINAL)
                            End If

                        End If


                        'CADENA FINAL CON LOS DATOS FINALES
                        sRespuesta = sXMLCAB.Trim & "*$¿*" & sXMLPIE.Trim & "*$¿*" & sXMLMOV_FINAL.Trim

                    End If

                    If lContador > 20 Then
                        Exit Do
                    End If

                Loop Until Microsoft.VisualBasic.Strings.Right(sTrama, 1) <> "C"


                'Evaluar Respuesta si es ERROR,  RU, AU (Respuesta correcta)
                If sTrama.Trim.Length > 0 Then
                    If Microsoft.VisualBasic.Strings.Left(sTrama.Trim, 2) = "RU" Then 'Mensaje de Validacion al usuario. RU            000CUENTA BLOQUEADA
                        sMensajeErrorUsuario = Mid(sTrama.Trim, 18, Len(sTrama.Trim))
                        sRespuesta = "ERROR:NODATA-" & sMensajeErrorUsuario.Trim
                    End If
                Else 'Sino devuelve nada
                    sRespuesta = "ERROR:Servicio no disponible."
                End If
            Else
                'Mostrar Mensaje de Error
                sRespuesta = "ERROR:Parametros Incompletos"
            End If

        Catch ex As Exception
            'save log error

            sRespuesta = "ERROR:" & ex.Message.Trim

        End Try


        'Save Monitor
        Dim sDataMonitor As String = ""
        Dim FECHA_TX As String = DateTime.Now.Year.ToString("0000").Trim & DateTime.Now.Month.ToString("00").Trim + DateTime.Now.Day.ToString("00").Trim
        Dim HORA_TX As String = DateTime.Now.Hour.ToString("00").Trim & DateTime.Now.Minute.ToString("00").Trim & DateTime.Now.Second.ToString("00").Trim
        Dim sNroCuentax As String = ""
        Dim sNroTarjetax As String = ""
        Dim sNombreCliente As String = ""
        Dim sModoEntrada As String = ""
        Dim sSaldoDisponible As String = ""
        Dim sTotalDeuda As String = ""
        Dim sIDSucursal As String = ""
        Dim sIDTerminal As String = ""
        Dim sCodigoTransaccion As String = ""

        Dim sRespuestaServidor As String = ""
        Dim sCanalAtencion As String = "01" 'RipleyMatico
        Dim sEstadoCuenta As String = ""



        If sDATA_MONITOR_KIOSCO.Length > 0 Then

            'RECORRER ARREGLO
            Dim ADATA_MONITOR As Array
            ADATA_MONITOR = Split(sDATA_MONITOR_KIOSCO, "|\t|", , CompareMethod.Text)


            sNroCuentax = ADATA_MONITOR(0)
            sNroTarjetax = ADATA_MONITOR(1)
            sNombreCliente = ADATA_MONITOR(2)
            sModoEntrada = ADATA_MONITOR(3)
            sSaldoDisponible = ADATA_MONITOR(4)
            sTotalDeuda = ADATA_MONITOR(5)
            sIDSucursal = ADATA_MONITOR(6)
            sIDTerminal = ADATA_MONITOR(7)
            sCodigoTransaccion = ADATA_MONITOR(8)

            'GUADAR LOG CONSULTAS
            'SAVE_LOG_CONSULTAS(sIDSucursal.Trim, sIDTerminal.Trim, "32", FUN_BUSCAR_TIPO_TARJETA(sNroTarjetax.Trim))

            sNroCuentax = Microsoft.VisualBasic.Strings.Right("          " & sNroCuentax.Trim, 10)
            sNroTarjetax = Microsoft.VisualBasic.Strings.Right(Strings.StrDup(20, " ") & sNroTarjetax.Trim, 20)
            sNombreCliente = Microsoft.VisualBasic.Strings.Right(Strings.StrDup(26, " ") & sNombreCliente.Trim, 26)

            sSaldoDisponible = Microsoft.VisualBasic.Strings.Right("      " & sSaldoDisponible.Trim, 6)
            sTotalDeuda = Microsoft.VisualBasic.Strings.Right("      " & sTotalDeuda.Trim, 6)


            'armar variable desde flash con: numero de cuenta, numero de tarjeta, nombre del cliente, modo de entrada, canal de atencion, id de sucursal,codigo kiosco y codigo de transaccion (opcion)
            If Microsoft.VisualBasic.Strings.Left(sRespuesta.Trim, 5) <> "ERROR" Then
                'VALIDACION CORRECTA

                sEstadoCuenta = "01" 'Estado de la cuenta
                sRespuestaServidor = "01" 'Atendido

                'ENVIAR_MONITOR
                sDataMonitor = ""
                sDataMonitor = FECHA_TX & HORA_TX & sNroCuentax & sNroTarjetax & sNombreCliente & sModoEntrada & sCanalAtencion & sSaldoDisponible & sTotalDeuda & sIDSucursal & sIDTerminal & sEstadoCuenta & sRespuestaServidor & sCodigoTransaccion

                'ENVIAR_MONITOR(sDataMonitor)

            Else
                'ENVIAR_MONITOR

                sEstadoCuenta = "01" 'Estado de la cuenta
                sRespuestaServidor = "02" 'Rechazado

                sDataMonitor = ""
                sDataMonitor = FECHA_TX & HORA_TX & sNroCuentax & sNroTarjetax & sNombreCliente & sModoEntrada & sCanalAtencion & sSaldoDisponible & sTotalDeuda & sIDSucursal & sIDTerminal & sEstadoCuenta & sRespuestaServidor & sCodigoTransaccion

                'ENVIAR_MONITOR(sDataMonitor)

            End If

        End If

        Return sRespuesta

    End Function

    Private Function GetTramaR192(ByVal CServidor As String, ByVal sNroTarjeta As String, ByVal sNroCuenta As String, ByVal sPeriodoFinal As String, ByVal sPagina As String, ByRef outpputBuff As String) As String
        Dim sParametros As String = ""
        'Variables nuevas
        Dim actions As Long = 1
        Dim inetputBuff As String = ""
        Dim errorMsg As String = ""

        'Instancia al mirapiweb
        'Dim obSendMirror As ClsTxMirapi = Nothing
        'obSendMirror = New ClsTxMirapi()

        Dim sTrama As String = ""

        sParametros = "0000000000" & CServidor & sNroTarjeta & sNroCuenta.Trim & sPeriodoFinal.Trim & sPagina
        inetputBuff = "      " + "R192" + sParametros
        ErrorLog("inetputBuff=" & inetputBuff)

        sTrama = ""
        'sTrama = obSendMirror.ExecuteTX(actions.ToString, inetputBuff, outpputBuff, TIME_OUT_SERVER, SERVIDOR_MIRROR_DESTINO, SERVIDOR_MIRROR_NODE, PUERTO, errorMsg)
        sTrama = "0"
        outpputBuff = Tramas(contadorArray)
        contadorArray += 1

        ErrorLog("sTrama=" & sTrama)

        outpputBuff = RTrim(outpputBuff)
        ErrorLog("outpputBuff=" & outpputBuff)

        Return sTrama
    End Function

    Private Sub btnCalcular_Click(sender As System.Object, e As System.EventArgs) Handles btnCalcular.Click

        Dim fecha As Date = DateTime.Now
        Dim lol23 As String = Format(fecha, "g")

        Dim sRespuesta As String = ""
        'VARIABLES RESUMEN DE ESTADO DE CUENTA
        Dim sLineaCredito As String = ""
        Dim sLineaCreditoTemporal As String = ""
        Dim sLineaSuperEfectivo As String = ""
        Dim sLineaCreditoTotal As String = ""
        Dim sPeriodoFacturacion As String = ""
        Dim sUltimoDiaPago As String = ""
        Dim sCreditoUtilizado As String = ""
        Dim sComisionCargos As String = ""
        Dim sDeudaTotal As String = ""
        Dim sDeudaVencida As String = ""
        Dim sPagoMinimoMes As String = ""
        Dim sPagoTotalMes As String = ""
        Dim sDispCompras24Cuotas As String = ""
        Dim sDispCompras36Cuotas As String = ""
        Dim sDispSuperEfectivo As String = ""
        Dim sDisponibleCompras As String = ""
        Dim sDispEfectivoExpress As String = ""

        'CALCULO DE PAGO TOTAL DEL MES
        Dim sSaldoAnterior As String = ""
        Dim sConsumoMes As String = ""
        Dim sCuotaMes As String = ""
        Dim sInteres As String = ""
        Dim sComisionCargos_Mes As String = ""
        Dim sPagosAbonos As String = ""
        Dim sPagoTotal_Mes As String = ""

        'CALCULO DE PAGO MINIMO DEL MES
        Dim sSaldoFavor As String = ""
        Dim sMontoMinimo As String = ""
        Dim sCuotaMes_Min As String = ""
        Dim sInteres_Min As String = ""
        Dim sComisionCargos_Min As String = ""
        Dim sPagoMinimoMes_Min As String = ""

        'PLAN DE CUOTAS DE LOS PROXIMOS 3 MESES
        Dim sMes1 As String = ""
        Dim sMonto1 As String = ""
        Dim sMes2 As String = ""
        Dim sMonto2 As String = ""
        Dim sMes3 As String = ""
        Dim sMonto3 As String = ""

        Dim sPeriodoFinal As String = ""
        Dim sPeriodo1 As String = ""
        Dim sPeriodo2 As String = ""
        Dim sPeriodo3 As String = ""
        Dim vFechaHoy As Date
        Dim fFechaHoyAux1 As Date
        Dim lEstadoLlamadaEECC As Long = 9 'Variable para iniciar llamadas al EECC 3 veces

        Dim sTrama As String = ""
        Dim sXMLMOV As String = ""
        Dim sXMLMOV_FINAL As String = ""
        Dim sXMLCAB As String = ""
        Dim sXMLPIE As String = ""
        Dim sDataMov As String = ""
        Dim sDataMovAux As String = ""

        Dim lFila As Long = 0
        Dim Incrementa As Long = 1
        Dim tamanioFilaDetalle As Long = 94
        Dim tamanioFilaDetalleAnt As Long = 87

        Dim outpputBuff As String = "        AU            008CABALLERO ZEÑA, DANIEL                            JR.SANTA MARIA 338 PS.04 INT.B                                                                      SAN MARTIN DE PORRES    LIMA            LIMA      52543-500328093-61                   1,500.00      0.00      0.00  1,500.0001/FEB/2015-31/MAR/201525/ABR/2015  2,136.81     58.77  2,195.69    186.16    392.47    392.47      0.00      0.00      0.00      0.00      0.00                                                                392.47                                                                                                                                                                                                                        186.16      5.90    101.48     59.91     39.00      0.00    392.47     11.80    101.48     59.91     39.00                                                            SALDO INICIAL                                               186.1620/FEB/201520/FEB/2015004125RET. EFEC CUOTAS TDAT  1000.05 43.74%02/24   29.10  30.91    60.0121/FEB/201521/FEB/2015004131RET. EFEC CUOTAS TDAT  1000.05 41.25%02/12   72.38  28.10   100.4831/MAR/201531/MAR/2015000100SEGURO DESGRAVAMEN  T     5.90       01/01                    5.90                            INTERES PERIODO FACT                                          0.31                            INTERES COMPENSATORI                                          0.59                            PENALIDAD PAGO VCDO.                                         39.00                                                                                                                                                                                                                                                                                                                                                                                                                                                                                          U    "
        outpputBuff = "        AU            006APARICIO SORIA, RONALD                            AV.CAMINO REAL 348 PISO 12                                                                          1401SAN ISIDRO          0015LIMA        LIMA      96041-005000244-91                   7,100.00      0.00      0.00  7,100.0026/NOV/2014-26/DIC/201420/ENE/2015    400.86      6.90    407.78      0.00     36.90    407.78  6,692.24      0.00      0.00      0.00  6,692.24    188.83    712.03      0.00      0.00      6.90    500.00    407.780000000000     30.00      0.00      0.00      6.90     36.90                                                                                                                                                                                                                                                                                                                                  SALDO INICIAL                                        188.8308/DIC/201409/DIC/2014370162MAKRO SUPERMAYORISTAT   706.1301/01                  706.1319/DIC/201419/DIC/2014049057PAGO TIENDA  RIPLEY T                               -500.0026/DIC/201426/DIC/2014000100SEGURO DESGRAVAMEN  T     5.9001/01                    5.90                            ENVIO EE.CC. MENSUAL                                   6.90                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                         U    "
        sPeriodoFinal = "201412"
        outpputBuff = RTrim(outpputBuff)
        sTrama = "0"
        If sTrama = "0" Then 'EXITO
            If outpputBuff.Length > 0 Then
                sTrama = outpputBuff.Substring(8, outpputBuff.Length - 8)
            End If

        End If





        MessageBox.Show(sTrama.Length)

        'VARIABLES RESUMEN DE ESTADO DE CUENTA
        sLineaCredito = Mid(sTrama, 253, 10)
        sDispEfectivoExpress = Mid(sTrama, 427, 10)
        sDisponibleCompras = Mid(sTrama, 387, 10)
        sPeriodoFacturacion = Mid(sTrama, 293, 23)
        sUltimoDiaPago = Mid(sTrama, 316, 11)

        sCreditoUtilizado = Mid(sTrama, 327, 10)
        sComisionCargos = Mid(sTrama, 337, 10)
        sDeudaTotal = Mid(sTrama, 347, 10)
        sDeudaVencida = Mid(sTrama, 357, 10)
        sPagoMinimoMes = Mid(sTrama, 367, 10)
        sPagoTotalMes = Mid(sTrama, 377, 10)


        'CALCULO DE PAGO TOTAL DEL MES
        sSaldoAnterior = Mid(sTrama, 437, 10)
        sConsumoMes = Mid(sTrama, 447, 10)
        sCuotaMes = Mid(sTrama, 457, 10)
        sInteres = Mid(sTrama, 467, 10)
        sComisionCargos_Mes = Mid(sTrama, 477, 10)
        sPagosAbonos = Mid(sTrama, 487, 10)
        sPagoTotal_Mes = Mid(sTrama, 497, 10)

        'CALCULO DE PAGO MINIMO DEL MES
        sSaldoFavor = Mid(sTrama, 507, 10)
        sMontoMinimo = Mid(sTrama, 517, 10)
        sCuotaMes_Min = Mid(sTrama, 527, 10)
        sInteres_Min = Mid(sTrama, 537, 10)
        sComisionCargos_Min = Mid(sTrama, 547, 10)
        sPagoMinimoMes_Min = Mid(sTrama, 557, 10)


        'PLAN DE CUOTAS DE LOS PROXIMOS 3 MESES
        sMes1 = Mid(sTrama, 615, 8)
        sMonto1 = Mid(sTrama, 623, 10)
        sMes2 = Mid(sTrama, 633, 8)
        sMonto2 = Mid(sTrama, 641, 10)
        sMes3 = Mid(sTrama, 651, 8)
        sMonto3 = Mid(sTrama, 659, 10)

        'MOVIMIENTOS
        sDataMovAux = ""
        sDataMov = ""
        sDataMov = Mid(sTrama, 861, sTrama.Length)


        Incrementa = 1
        If sDataMov.Length > 0 Then

            If sPeriodoFinal >= Constantes.PeriodoInclusionTEA Then
                For lFila = 1 To 13

                    If Trim(Mid(sDataMov, Incrementa, tamanioFilaDetalle)) <> "" Then
                        sDataMovAux = sDataMovAux & Mid(sDataMov, Incrementa, tamanioFilaDetalle) & "|\n|"
                    End If

                    Incrementa = Incrementa + tamanioFilaDetalle

                Next
            Else
                For lFila = 1 To 13

                    If Trim(Mid(sDataMov, Incrementa, tamanioFilaDetalleAnt)) <> "" Then
                        sDataMovAux = sDataMovAux & Mid(sDataMov, Incrementa, tamanioFilaDetalleAnt) & "|\n|"
                    End If

                    Incrementa = Incrementa + tamanioFilaDetalleAnt

                Next
            End If


        End If

        sXMLCAB = ""
        sXMLCAB = sLineaCredito.Trim & "|\t|" & sDispEfectivoExpress.Trim & "|\t|" & sDisponibleCompras.Trim
        sXMLCAB = sXMLCAB & "|\t|" & sPeriodoFacturacion.Trim & "|\t|" & sUltimoDiaPago.Trim
        sXMLCAB = sXMLCAB & "|\t|" & sCreditoUtilizado.Trim & "|\t|" & sComisionCargos.Trim & "|\t|" & sDeudaTotal.Trim
        sXMLCAB = sXMLCAB & "|\t|" & sDeudaVencida.Trim & "|\t|" & sPagoMinimoMes.Trim & "|\t|" & sPagoTotalMes.Trim


        sXMLPIE = ""
        sXMLPIE = sSaldoAnterior.Trim & "|\t|" & sConsumoMes.Trim & "|\t|" & sCuotaMes.Trim
        sXMLPIE = sXMLPIE & "|\t|" & sInteres.Trim & "|\t|" & sComisionCargos_Mes.Trim & "|\t|" & sPagosAbonos.Trim
        sXMLPIE = sXMLPIE & "|\t|" & sPagoTotal_Mes.Trim & "|\t|" & sSaldoFavor.Trim & "|\t|" & sMontoMinimo.Trim
        sXMLPIE = sXMLPIE & "|\t|" & sCuotaMes_Min.Trim & "|\t|" & sInteres_Min.Trim & "|\t|" & sComisionCargos_Min.Trim
        sXMLPIE = sXMLPIE & "|\t|" & sPagoMinimoMes_Min.Trim & "|\t|" & sMes1.Trim & "|\t|" & sMonto1.Trim
        sXMLPIE = sXMLPIE & "|\t|" & sMes2.Trim & "|\t|" & sMonto2.Trim & "|\t|" & sMes3.Trim & "|\t|" & sMonto3.Trim

        Dim lol As String = Microsoft.VisualBasic.Strings.Right(sTrama, 1)
        If Microsoft.VisualBasic.Strings.Right(sTrama, 1) = "U" Then
            'If Left(sTrama, 2) = "AU" Then 'FIN DE LA SOLICITUD

            If sDataMovAux.Length > 0 Then
                sDataMovAux = Microsoft.VisualBasic.Strings.Left(sDataMovAux, sDataMovAux.Length - 4)
                sXMLMOV = sXMLMOV & sDataMovAux

            Else
                sDataMovAux = "" 'NO HAY MOVIMIENTOS
                sXMLMOV = sXMLMOV & sDataMovAux


            End If

            If sXMLMOV.Length > 0 Then
                'Armar las columnas de los movimientos de EECC
                sXMLMOV_FINAL = FUN_DETALLE_EECC_TEA(sXMLMOV, sPeriodoFinal)

            End If

        End If

        'CADENA FINAL CON LOS DATOS FINALES
        sRespuesta = sXMLCAB.Trim & "*$¿*" & sXMLPIE.Trim & "*$¿*" & sXMLMOV_FINAL.Trim
        lblRespuesta.Text = sRespuesta
    End Sub

    'DEVUELVE DETALLE DE EECC
    '<WebMethod(Description:="Devuelve el detalle del EECC Clasica")> _
    Private Function FUN_DETALLE_EECC_TEA2(ByVal sDataDetalle As String, ByVal sPeriodoFinal As String) As String

        Dim sResult As String = ""
        Dim SDATA_MOVIMIENTOS As String = ""
        Dim ADATA_MOVIMIENTO As Array
        Dim lIndice As Long = 0
        Dim sRegistroAux As String = ""
        Dim sRegistro As String = ""
        Dim sMOV As String = ""

        'CAMPOS MOVIMIENTOS
        Dim sFechaConsumo As String = ""
        Dim sFechaProceso As String = ""
        Dim sNroTicket As String = ""
        Dim sDescripcion As String = ""
        Dim sTA As String = ""
        Dim sMonto As String = ""
        Dim sTEA As String = ""
        Dim sNroCuotas As String = ""
        Dim sCapital As String = ""
        Dim sInteres As String = ""
        Dim sTotal As String = ""


        SDATA_MOVIMIENTOS = sDataDetalle

        If SDATA_MOVIMIENTOS.Trim.Length > 0 Then 'IF HAY REGISTROS
            ADATA_MOVIMIENTO = Split(SDATA_MOVIMIENTOS, "|\n|", , CompareMethod.Text)

            For lIndice = 0 To ADATA_MOVIMIENTO.Length - 1

                sRegistro = ""
                sRegistro = ADATA_MOVIMIENTO(lIndice)

                sFechaConsumo = Mid(sRegistro, 1, 11)
                sFechaProceso = Mid(sRegistro, 12, 11)
                sNroTicket = Mid(sRegistro, 23, 6)
                sDescripcion = Mid(sRegistro, 29, 20)
                sTA = Mid(sRegistro, 49, 1)
                sMonto = Mid(sRegistro, 50, 9)
                If sPeriodoFinal >= Constantes.PeriodoInclusionTEA Then
                    sTEA = Mid(sRegistro, 59, 7)
                    sNroCuotas = Mid(sRegistro, 66, 5)
                    sCapital = Mid(sRegistro, 71, 8)
                    sInteres = Mid(sRegistro, 79, 8)
                    sTotal = Mid(sRegistro, 87, 8)
                Else
                    sTEA = "       "
                    sNroCuotas = Mid(sRegistro, 59, 5)
                    sCapital = Mid(sRegistro, 64, 8)
                    sInteres = Mid(sRegistro, 72, 8)
                    sTotal = Mid(sRegistro, 80, 8)
                End If

                sRegistroAux = ""
                sRegistroAux = sFechaConsumo.Trim & "|\t|" & sFechaProceso.Trim & "|\t|" & sNroTicket.Trim & "|\t|" & sDescripcion.Trim & "|\t|" & sTA.Trim & "|\t|"
                sRegistroAux = sRegistroAux & sMonto.Trim & "|\t|" & sTEA.Trim & "|\t|" & sNroCuotas.Trim & "|\t|" & sCapital.Trim & "|\t|" & sInteres.Trim & "|\t|"
                sRegistroAux = sRegistroAux & sTotal.Trim & "|\n|"

                sMOV = sMOV & sRegistroAux.Trim

            Next


            If sMOV.Trim.Length > 0 Then
                sMOV = Microsoft.VisualBasic.Left(sMOV, sMOV.Length - 4)
            Else
                sMOV = "" 'NO HAY MOVIMIENTOS
            End If

            sResult = sMOV

        End If

        Return sResult.Trim

    End Function

    ''DEVUELVE DETALLE DE EECC
    ''<WebMethod(Description:="Devuelve el detalle del EECC Clasica")> _
    'Private Function FUN_DETALLE_EECC_TEA(ByVal sDataDetalle As String, ByVal sPeriodoFinal As String) As String

    '    Dim sResult As String = ""
    '    Dim SDATA_MOVIMIENTOS As String = ""
    '    Dim ADATA_MOVIMIENTO As Array
    '    Dim lIndice As Long = 0
    '    Dim sRegistroAux As String = ""
    '    Dim sRegistro As String = ""
    '    Dim sMOV As String = ""

    '    'CAMPOS MOVIMIENTOS
    '    Dim sFechaConsumo As String = ""
    '    Dim sFechaProceso As String = ""
    '    Dim sNroTicket As String = ""
    '    Dim sDescripcion As String = ""
    '    Dim sTA As String = ""
    '    Dim sMonto As String = ""
    '    Dim sTEA As String = ""
    '    Dim sNroCuotas As String = ""
    '    Dim sCapital As String = ""
    '    Dim sInteres As String = ""
    '    Dim sTotal As String = ""


    '    SDATA_MOVIMIENTOS = sDataDetalle

    '    If SDATA_MOVIMIENTOS.Trim.Length > 0 Then 'IF HAY REGISTROS
    '        ADATA_MOVIMIENTO = Split(SDATA_MOVIMIENTOS, "|\n|", , CompareMethod.Text)

    '        For lIndice = 0 To ADATA_MOVIMIENTO.Length - 1

    '            sRegistro = ""
    '            sRegistro = ADATA_MOVIMIENTO(lIndice)

    '            sFechaConsumo = Mid(sRegistro, 1, 11)
    '            sFechaProceso = Mid(sRegistro, 12, 11)
    '            sNroTicket = Mid(sRegistro, 23, 6)
    '            sDescripcion = Mid(sRegistro, 29, 85)
    '            sTA = Mid(sRegistro, 114, 1)
    '            sMonto = Mid(sRegistro, 115, 9)

    '            If sPeriodoFinal >= Constantes.PeriodoInclusionTEA Then
    '                sTEA = Mid(sRegistro, 124, 7)
    '                sNroCuotas = Mid(sRegistro, 131, 5)
    '                sCapital = Mid(sRegistro, 136, 8)
    '                sInteres = Mid(sRegistro, 144, 8)
    '                sTotal = Mid(sRegistro, 152, 8)
    '            Else
    '                sTEA = "       "
    '                sNroCuotas = Mid(sRegistro, 124, 5)
    '                sCapital = Mid(sRegistro, 129, 8)
    '                sInteres = Mid(sRegistro, 137, 8)
    '                sTotal = Mid(sRegistro, 145, 8)
    '            End If



    '            sRegistroAux = ""
    '            sRegistroAux = sFechaConsumo.Trim & "|\t|" & sFechaProceso.Trim & "|\t|" & sNroTicket.Trim & "|\t|" & sDescripcion.Trim & "|\t|" & sTA.Trim & "|\t|"
    '            sRegistroAux = sRegistroAux & sMonto.Trim & "|\t|" & sTEA.Trim & "|\t|" & sNroCuotas.Trim & "|\t|" & sCapital.Trim & "|\t|" & sInteres.Trim & "|\t|"
    '            sRegistroAux = sRegistroAux & sTotal.Trim & "|\n|"

    '            sMOV = sMOV & sRegistroAux.Trim

    '        Next


    '        If sMOV.Trim.Length > 0 Then
    '            sMOV = Microsoft.VisualBasic.Strings.Left(sMOV, sMOV.Length - 4)
    '        Else
    '            sMOV = "" 'NO HAY MOVIMIENTOS
    '        End If

    '        sResult = sMOV

    '    End If

    '    Return sResult.Trim

    'End Function

    '10-03-2015 Ocultar para pase por partes
    'DEVUELVE DETALLE DE EECC
    '<WebMethod(Description:="Devuelve el detalle del EECC Clasica")> _
    Private Function FUN_DETALLE_EECC_TEA(ByVal sDataDetalle As String, ByVal sPeriodoFinal As String) As String

        Dim sResult As String = ""
        Dim SDATA_MOVIMIENTOS As String = ""
        Dim ADATA_MOVIMIENTO As Array
        Dim lIndice As Long = 0
        Dim sRegistroAux As String = ""
        Dim sRegistro As String = ""
        Dim sMOV As String = ""

        'CAMPOS MOVIMIENTOS
        Dim sFechaConsumo As String = ""
        Dim sFechaProceso As String = ""
        Dim sNroTicket As String = ""
        Dim sDescripcion As String = ""
        Dim sTA As String = ""
        Dim sMonto As String = ""
        Dim sTEA As String = ""
        Dim sNroCuotas As String = ""
        Dim sCapital As String = ""
        Dim sInteres As String = ""
        Dim sTotal As String = ""


        SDATA_MOVIMIENTOS = sDataDetalle

        If SDATA_MOVIMIENTOS.Trim.Length > 0 Then 'IF HAY REGISTROS
            ADATA_MOVIMIENTO = Split(SDATA_MOVIMIENTOS, "|\n|", , CompareMethod.Text)

            For lIndice = 0 To ADATA_MOVIMIENTO.Length - 1

                sRegistro = ""
                sRegistro = ADATA_MOVIMIENTO(lIndice)

                sFechaConsumo = Mid(sRegistro, 1, 11)
                sFechaProceso = Mid(sRegistro, 12, 11)
                sNroTicket = Mid(sRegistro, 23, 6)

                If sPeriodoFinal >= Constantes.PeriodoAgrandarGlosa Then
                    sDescripcion = Mid(sRegistro, 29, 85)
                    sTA = Mid(sRegistro, 114, 1)
                    sMonto = Mid(sRegistro, 115, 9)

                    sTEA = Mid(sRegistro, 124, 7)
                    sNroCuotas = Mid(sRegistro, 131, 5)
                    sCapital = Mid(sRegistro, 136, 8)
                    sInteres = Mid(sRegistro, 144, 7)
                    sTotal = Mid(sRegistro, 151, 9)

                Else
                    sDescripcion = Mid(sRegistro, 29, 20)
                    sTA = Mid(sRegistro, 49, 1)
                    sMonto = Mid(sRegistro, 50, 9)

                    If sPeriodoFinal >= Constantes.PeriodoInclusionTEA Then
                        sTEA = Mid(sRegistro, 59, 7)
                        sNroCuotas = Mid(sRegistro, 66, 5)
                        sCapital = Mid(sRegistro, 71, 8)
                        sInteres = Mid(sRegistro, 79, 8)
                        sTotal = Mid(sRegistro, 87, 8)
                    Else
                        sTEA = "       "
                        sNroCuotas = Mid(sRegistro, 59, 5)
                        sCapital = Mid(sRegistro, 64, 8)
                        sInteres = Mid(sRegistro, 72, 8)
                        sTotal = Mid(sRegistro, 80, 8)
                    End If

                End If

                sRegistroAux = ""
                sRegistroAux = sFechaConsumo.Trim & "|\t|" & sFechaProceso.Trim & "|\t|" & sNroTicket.Trim & "|\t|" & sDescripcion.Trim & "|\t|" & sTA.Trim & "|\t|"
                sRegistroAux = sRegistroAux & sMonto.Trim & "|\t|" & sTEA.Trim & "|\t|" & sNroCuotas.Trim & "|\t|" & sCapital.Trim & "|\t|" & sInteres.Trim & "|\t|"
                sRegistroAux = sRegistroAux & sTotal.Trim & "|\n|"

                sMOV = sMOV & sRegistroAux.Trim

            Next


            If sMOV.Trim.Length > 0 Then
                sMOV = Microsoft.VisualBasic.Strings.Left(sMOV, sMOV.Length - 4)
            Else
                sMOV = "" 'NO HAY MOVIMIENTOS
            End If

            sResult = sMOV

        End If

        Return sResult.Trim

    End Function


    Public Function ObtenerFechaFacturacion(ByVal ultimoDiaPago As String, ByVal mes As String, ByVal anio As String) As String
        Dim diasMes As Integer = 0
        diasMes = Date.DaysInMonth(CInt(anio), CInt(mes))
        Dim respuesta As String = String.Empty
        Try

            Select Case ultimoDiaPago
                Case 1
                    Select Case diasMes
                        Case 28
                            respuesta = "04"
                        Case 29
                            respuesta = "05"
                        Case 30
                            respuesta = "06"
                        Case 31
                            respuesta = "07"
                        Case Else
                            respuesta = "XX"
                    End Select
                Case 5
                    Select Case diasMes
                        Case 28
                            respuesta = "08"
                        Case 29
                            respuesta = "09"
                        Case 30
                            respuesta = "10"
                        Case 31
                            respuesta = "11"
                        Case Else
                            respuesta = "XX"
                    End Select
                Case 10
                    Select Case diasMes
                        Case 28
                            respuesta = "13"
                        Case 29
                            respuesta = "14"
                        Case 30
                            respuesta = "15"
                        Case 31
                            respuesta = "16"
                        Case Else
                            respuesta = "XX"
                    End Select
                Case 15
                    Select Case diasMes
                        Case 28
                            respuesta = "18"
                        Case 29
                            respuesta = "19"
                        Case 30
                            respuesta = "20"
                        Case 31
                            respuesta = "21"
                        Case Else
                            respuesta = "XX"
                    End Select
                Case 20
                    Select Case diasMes
                        Case 28
                            respuesta = "23"
                        Case 29
                            respuesta = "24"
                        Case 30
                            respuesta = "25"
                        Case 31
                            respuesta = "26"
                        Case Else
                            respuesta = "XX"
                    End Select
                Case 25
                    Select Case diasMes
                        Case 28
                            respuesta = "28"
                        Case 29
                            respuesta = "29"
                        Case 30
                            respuesta = "30"
                        Case 31
                            respuesta = "31"
                        Case Else
                            respuesta = "XX"
                    End Select
                Case Else
                    respuesta = "XX"

            End Select

            If respuesta.Contains("XX") Then
                respuesta = "ERROR:NODATA-Último día de pago es incorrecto."
            Else
                respuesta = anio & mes & respuesta
            End If

        Catch ex As Exception
            respuesta = "ERROR:NODATA-Hubo error en llamada a método ObtenerFechaFacturacion " & ex.Message
        End Try

        Return respuesta
    End Function

    Enum TServidor

        SICRON = 1
        RSAT = 2

    End Enum

    Private Sub ObtenerTramaEECC_Click(sender As System.Object, e As System.EventArgs) Handles ObtenerTramaEECC.Click
        'Dim respuesta As String = ESTADO_CUENTA_CLASICA_RSAT("5420700010596464", "05174142", "", TServidor.RSAT)
        Dim respuesta As String = ESTADO_CUENTA_CLASICA_RSAT("5254350068889667", "000005278239", "000005278239|\t|5254350068889667|\t|CASTILLO ALFARO, MIJAIL ALEXANDER|\t|00|\t|      |\t|      |\t|2|\t|SIRIP-99|\t|13", TServidor.RSAT)
        ErrorLog("respuesta = " & respuesta)
        txtSalidaEECC.Text = respuesta
    End Sub

    Public Function ObtenerOfertaCambioProducto(ByVal contrato As String) As Response
        Dim respuesta As New Response
        respuesta.Success = False
        respuesta.Warning = False
        Dim oferta As New OfertaCambioProducto
        Try
            ErrorLog("Entro al método ObtenerOfertaCambioProducto(" & contrato & ")")
            'oferta = BNOfertaCambioProducto.Instancia.ObtenerOferta(contrato)
            oferta.ContratoTarjeta = contrato
            oferta.DatosTarjeta = "014002TARJETA PLATINUM VISA SAT                         TARJETA CLASICA SAT                               "
            oferta.DatosSEF = "01"
            If oferta.ContratoTarjeta <> String.Empty Then
                Dim validaTarjeta As String
                Dim productoOrigenTarjeta As String
                Dim productoDestinoTarjeta As String
                Dim descripcionProductoOrigenTarjeta As String
                Dim descripcionProductoDestinoTarjeta As String

                Dim validaSEF As String
                Dim productoOrigenSEF As String
                Dim productoDestinoSEF As String
                Dim descripcionProductoOrigenSEF As String
                Dim descripcionProductoDestinoSEF As String

                If oferta.DatosTarjeta.Length = 106 Then
                    validaTarjeta = oferta.DatosTarjeta.Substring(0, 2)
                    productoOrigenTarjeta = oferta.DatosTarjeta.Substring(2, 2)
                    productoDestinoTarjeta = oferta.DatosTarjeta.Substring(4, 2)
                    descripcionProductoOrigenTarjeta = oferta.DatosTarjeta.Substring(6, 50)
                    descripcionProductoDestinoTarjeta = oferta.DatosTarjeta.Substring(56, 50)
                ElseIf oferta.DatosTarjeta.Length = 2 Then
                    validaTarjeta = oferta.DatosTarjeta.Substring(0, 2)
                End If

                If oferta.DatosSEF.Length = 106 Then
                    validaSEF = oferta.DatosSEF.Substring(0, 2)
                    productoOrigenSEF = oferta.DatosSEF.Substring(2, 2)
                    productoDestinoSEF = oferta.DatosSEF.Substring(4, 2)
                    descripcionProductoOrigenSEF = oferta.DatosSEF.Substring(6, 50)
                    descripcionProductoDestinoSEF = oferta.DatosSEF.Substring(56, 50)
                ElseIf oferta.DatosSEF.Length = 2 Then
                    validaSEF = oferta.DatosSEF.Substring(0, 2)
                End If
            Else
                respuesta.Warning = True
                respuesta.Message = "No tiene oferta de cambio de producto"
            End If

            respuesta.Success = True
        Catch ex As Exception
            respuesta.Message = "En estos momentos no podemos atenderle. Por favor inténtelo más tarde"
            ErrorLog("ObtenerOfertaCambioProducto Error =" & ex.Message)
        End Try

        Return respuesta
    End Function

    Private Sub btnCambioProducto_Click(sender As System.Object, e As System.EventArgs) Handles btnCambioProducto.Click

        AddOUTParameters(New VTAAUTO5001ParameterOUT)
        Dim lo232 As String = parameterOut.ToString()

        Dim tramaRespuesta As String = "0000000000VTAAUTO5001                                       0000010240001000004SIRIP-990188VACP2015-02-11RMPES                                                                                                                                                                                                                                                                                                                                                                                                                 000 ES                                                                                        00010001000005351359C 10085380    2SAT     188         SIRIP-99    10.25.4.6      04801100000000000000000000000200EL CONTRATO CUMPLE CON LAS CONDICIONES DE CAMBIO DE PRODUCTO. CODIGO DE CAMBIO DE PRODUCYO (0480) ENCONTRADO"
        Dim tananio As Integer = tramaRespuesta.Length
        Dim variante As Integer = 92
        Dim CuentaENC As String = tramaRespuesta.Substring(639, 1)
        Dim CodCambioENC As String = tramaRespuesta.Substring(640, 1)
        Dim TieneSEF As String = tramaRespuesta.Substring(641, 1)
        Dim ContratoSEF As String = tramaRespuesta.Substring(695 - variante, 20)
        Dim CambioOK As String = tramaRespuesta.Substring(715 - variante, 1)

        Dim Paso As String = tramaRespuesta.Substring(716 - variante, 2)
        Dim CodRpta As String = tramaRespuesta.Substring(718 - variante, 2)
        Dim DesRpta As String = tramaRespuesta.Substring(720 - variante, 350)

        Dim fechaActual As String = Date.Now.ToShortDateString()
        FormatearFecha(fechaActual)
        ObtenerOfertaCambioProducto("00010001000000011111")
    End Sub

    Private Function FormatearFecha(ByVal sFecha As String) As String
        Dim sResul As String = ""
        Dim sDia As String = ""
        Dim sMes As String = ""
        Dim sAnio As String = ""

        If sFecha.Trim.Length > 0 Then
            sDia = Microsoft.VisualBasic.Strings.Left(sFecha.Trim, 2)
            sMes = Mid(sFecha.Trim, 3, 2)
            sAnio = Microsoft.VisualBasic.Strings.Right(sFecha.Trim, 4)

            sResul = sDia.Trim & "/" & sMes.Trim & "/" & sAnio.Trim

        End If


        Return sResul.Trim

    End Function

    Sub AddOUTParameters(ByVal parametrosOUT As VTAAUTO5001ParameterOUT)

        parameterOut.CuentaENC = fx_Completar_Campo("0", 1, parametrosOUT.CuentaENC, TYPE_ALINEAR.DERECHA)
        parameterOut.ProductoO = fx_Completar_Campo("0", 2, parametrosOUT.ProductoO, TYPE_ALINEAR.DERECHA)
        parameterOut.CodCambioENC = fx_Completar_Campo("0", 1, parametrosOUT.CodCambioENC, TYPE_ALINEAR.DERECHA)
        parameterOut.CuentaPlataformaEnc = fx_Completar_Campo("0", 1, parametrosOUT.CuentaPlataformaEnc, TYPE_ALINEAR.DERECHA)
        parameterOut.TieneSEF = fx_Completar_Campo("0", 1, parametrosOUT.TieneSEF, TYPE_ALINEAR.DERECHA)
        parameterOut.ContratoSEF = fx_Completar_Campo("0", 20, parametrosOUT.ContratoSEF, TYPE_ALINEAR.DERECHA)
        parameterOut.CambioOK = fx_Completar_Campo("0", 1, parametrosOUT.CambioOK, TYPE_ALINEAR.DERECHA)
        parameterOut.Paso = fx_Completar_Campo("0", 2, parametrosOUT.Paso, TYPE_ALINEAR.DERECHA)
        parameterOut.CodRpta = fx_Completar_Campo("0", 2, parametrosOUT.CodRpta, TYPE_ALINEAR.DERECHA)

    End Sub

    Private Sub Form1_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        Dim lol As String = "27/MAR/201527/MAR/2015004230COMPRA EN CUOTAS                                                                        1400.00"
        Dim rpta As String = Validar()
    End Sub

    Public Function Validar() As String
        Dim TipProducto As String = "2"
        Dim TipDocumento As String = "C"
        Dim NroDocumento As String = "44780645"
        Dim sRespuesta As String = "0000000000SFSCANC0040                                       000002000PE00010000020027X8  0       C44780645    1                                           00APAZA               CAHUANA             ANTHONY SILVERIO    026756840200010001000002615323050001-01-0101TI0000000000024500340005509259      46                              3000010001000002615323050001-01-0101TI0000000000014500340005509267      00                              30                    00          00  000000000000                      00                                                    00          00  000000000000                      00                                                    00          00  000000000000                      00                                                    00          00  000000000000                      00                                                    00          00  000000000000                      00                                                    00          00  000000000000                      00                                                    00          00  000000000000                      00                                                    00          00  000000000000                      00                                                    00          00  000000000000                      00                                                    00          00  000000000000                      00                                                    00          00  000000000000                      00                                                    00          00  000000000000                      00                                                    00          00  000000000000                      00                                                    00          00  000000000000                      00                                                    00          00  000000000000                      00                                      "
        Dim sNroTarjeta As String = String.Empty
        Dim sNroDocumento As String = "44780645"
        Dim sApellidoParterno As String = String.Empty
        Dim sApellidoMaterno As String = String.Empty
        Dim sNombres As String = String.Empty
        Dim sCliente As String = String.Empty
        Dim sFechaNac As String = String.Empty
        Dim sTipoProducto As String = String.Empty 'Segun el bin de la tarjeta 6 primeros numeros de la tarjeta
        Dim sTipoDocumento As String = String.Empty
        Dim sNroCuenta As String = String.Empty
        Dim sXMLRespuesta As String = String.Empty
        Dim Nro_Contrato As String
        Dim Tipo_Tarjeta_Ofertas As String

        If sRespuesta.Trim.Length > 0 Then
            If Microsoft.VisualBasic.Strings.Left(sRespuesta.Trim, 5) = "ERROR" Then
                'Save Error
                sRespuesta = "ERROR:Servicio no disponible."

            Else
                ErrorLog("Inicio de Corte")
                '---- INICIO CORTE ----'
                Dim fnEstMigra As String
                fnEstMigra = Mid(sRespuesta, 155, 2)
                If fnEstMigra = "00" Then
                    ErrorLog("fnEstMigra = 00")
                    sNombres = Mid$(sRespuesta, 197, 20)            '-> Nombres
                    sApellidoParterno = Mid$(sRespuesta, 157, 20)   '-> Apellido Paterno
                    sApellidoMaterno = Mid$(sRespuesta, 177, 20)    '-> Apellido Materno
                    sCliente = sApellidoParterno.Trim & " " & sApellidoMaterno.Trim & ", " & sNombres.Trim
                    Nro_Contrato = Mid$(sRespuesta, 227, 20)
                    Tipo_Tarjeta_Ofertas = "RSAT"
                    Dim int_cont, iIndEle, iLonBucle As Integer
                    Dim bCondCli As Boolean

                    Dim vg_RSAT_Sitarj, vg_RSAT_CodBloqueo, vg_RSAT_FecBaj, vg_RSAT_CodTitularAdicional As String

                    int_cont = Val(Mid(sRespuesta, 225, 2))
                    iLonBucle = 104
                    bCondCli = False

                    For iIndEle = 1 To int_cont
                        sNroCuenta = Mid(sRespuesta, 235 + iLonBucle * (iIndEle - 1), 12)
                        vg_RSAT_Sitarj = Mid(sRespuesta, 247 + iLonBucle * (iIndEle - 1), 2) 'SITUACION TARJETA
                        vg_RSAT_FecBaj = Mid(sRespuesta, 249 + iLonBucle * (iIndEle - 1), 10)
                        vg_RSAT_CodTitularAdicional = Mid(sRespuesta, 261 + iLonBucle * (iIndEle - 1), 2)
                        sNroTarjeta = Mid(sRespuesta, 275 + iLonBucle * (iIndEle - 1), 22)
                        vg_RSAT_CodBloqueo = Mid(sRespuesta, 297 + iLonBucle * (iIndEle - 1), 2)

                        If sNroCuenta.Length = 0 Then
                            sRespuesta = "ERROR: DNI no encontrado"
                            Exit For
                        End If

                        If VALIDAR_BLOQUEDO_TARJETA_RSAT(vg_RSAT_CodBloqueo) = False And vg_RSAT_Sitarj = "05" And vg_RSAT_FecBaj = "0001-01-01" And vg_RSAT_CodTitularAdicional = "TI" Then
                            bCondCli = True
                        Else
                            If VALIDAR_BLOQUEDO_TARJETA_RSAT(vg_RSAT_CodBloqueo) = False And vg_RSAT_Sitarj = "05" And vg_RSAT_FecBaj = "0001-01-01" And vg_RSAT_CodTitularAdicional = "BE" Then
                                bCondCli = True
                            End If
                        End If

                        '<INI TCK-563699-01 DHERRERA 20-03-2014>
                        'If bCondCli Then
                        If bCondCli = True And sNroTarjeta.Substring(6, 3) <> "761" Then
                            '<FIN TCK-563699-01 DHERRERA 20-03-2014>
                            If getTipProducto_AbiertaRSAT(sNroTarjeta.Substring(0, 6)).ToString = TipProducto Then

                                If TipDocumento = "C" Then 'DNI
                                    sTipoDocumento = "1"
                                Else
                                    sTipoDocumento = "2" 'Carnet de Extranjeria
                                End If

                                sTipoProducto = Mid(sNroTarjeta.Trim, 1, 6)
                                sRespuesta = sNroTarjeta.Trim & "|\t|" & NroDocumento.Trim & "|\t|" & sCliente.Trim & "|\t|" & sFechaNac.Trim & "|\t|" & sTipoProducto.Trim & "|\t||\t|RSAT"
                                Return sRespuesta
                            Else
                                sXMLRespuesta = "ERROR: DNI no encontrado"
                            End If

                            If iIndEle = int_cont Then
                                Return sXMLRespuesta
                            End If

                        Else

                            If iIndEle = int_cont Then
                                Return "ERROR:Servicio no disponible."
                            End If

                        End If


                    Next iIndEle



                Else
                    sRespuesta = "ERROR: DNI no encontrado"
                End If
            End If

        Else
            'Sino Encontro DNI en RSAT
            sRespuesta = "ERROR: DNI no encontrado"
        End If

        Return sRespuesta
    End Function

    Private Function VALIDAR_BLOQUEDO_TARJETA_RSAT(ByVal cod_bloquedo As String) As Boolean

        Dim ind As Byte
        Dim bTFEncontrada As Boolean
        bTFEncontrada = False

        Try
            For ind = 0 To c_TablaBloqueo.GetLength(0) - 1
                If cod_bloquedo = "00" Then
                    Exit For
                ElseIf cod_bloquedo = c_TablaBloqueo(ind, 0) And c_TablaBloqueo(ind, 3) = "S" Then
                    bTFEncontrada = True
                    Exit For
                End If
            Next
        Catch ex As Exception
            bTFEncontrada = False
        End Try

        Return bTFEncontrada

    End Function

    Private Function getTipProducto_AbiertaRSAT(ByVal BINN As String) As Integer

        Dim codigo As Integer

        Select Case BINN

            Case "525435"
                codigo = 2
            Case "542070"
                codigo = 1
            Case "450000"
                codigo = 5
            Case "450007"
                codigo = 4
                '<INI TCK-563699-01 DHERRERA 20-03-2014>
            Case "542020"
                codigo = 1
            Case "525474"
                codigo = 2
                '<FIN TCK-563699-01 DHERRERA 20-03-2014>
        End Select

        Return codigo

    End Function

    Private Sub BtnTitular_Click(sender As System.Object, e As System.EventArgs) Handles BtnTitular.Click
        txtDia.Text = Mid(txtSalidaEECC.Text, 261, 2)
    End Sub

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        Dim fechaActual As String = Date.Now.ToShortDateString()
        Dim dia As String = String.Empty
        Dim mes As String = String.Empty
        Dim anio As String = String.Empty
        Dim fechaResultante As String = String.Empty

        If fechaActual.Trim.Length > 0 Then
            dia = Microsoft.VisualBasic.Strings.Left(fechaActual.Trim, 2)
            mes = Mid(fechaActual.Trim, 4, 2)
            anio = Microsoft.VisualBasic.Strings.Right(fechaActual.Trim, 4)

            fechaResultante = anio.Trim() & "/" & mes.Trim() & "/" & dia.Trim()

        End If
    End Sub
End Class

