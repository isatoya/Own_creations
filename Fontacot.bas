Attribute VB_Name = "Fontacot"
Option Explicit
'variables para todo el proyecto

    Dim Fecha1, Fecha2, mes, Mes_Texto, año, dia As String
    Dim ruta, Ruta_Año, Ruta_Mes, Ruta_soporte As String
    Dim mes_f_inicial, mes_f_fin, cator_f_inicial, cator_f_fin, sem_f_inicial, sem_f_fin As String
    Dim wbConsolidado, wbCreditosInicial As Workbook
    Dim wbMensual, wbCatorcenal, wbSemanal, wbAnalisis As Workbook
    Dim CelsFecha As Range
    Dim lastRow, i As Long
    Dim CelsAG As Object
    
Sub Ejecutar_fanacot_P1()

    ' Verificar si hay datos en las celdas I8 y M8
    If ThisWorkbook.Sheets("Principal").Range("L7").Value = "" Or ThisWorkbook.Sheets("Principal").Range("H7").Value = "" Then
        MsgBox "Datos incompletos, por favor ingrese los datos antes de ejecutar.", vbExclamation
        Exit Sub
    End If

    ' Llama a cada una de las funciones
    CrearCarpetas
    Consolidado
    sap
    Fanacot
    
    MsgBox "El informe ha sido completado. Por favor, proceda con la revisión de las fórmulas faltantes que están vinculadas a otros archivos, así como con cualquier otro ajuste necesario para el análisis.", vbInformation


End Sub

Sub Ejecutar_fanacot_P2()

    ' Verificar si hay datos en las celdas I8 y M8
    If ThisWorkbook.Sheets("Principal").Range("L7").Value = "" Or ThisWorkbook.Sheets("Principal").Range("H7").Value = "" Then
        MsgBox "Datos incompletos, por favor ingrese los datos antes de ejecutar.", vbExclamation
        Exit Sub
    End If

    ' Llama a cada una de las funciones
    CrearCarpetas
    Analisis_consolidado
    
    MsgBox "El informe ha sido completado y se encuentra abierto. Por favor complete los campos de mes y los pegarlos # de empleado que van en el papel de trabajo", vbInformation

End Sub


Sub Ejecutar_fanacot_P3()

    ' Verificar si hay datos en las celdas I8 y M8
    If ThisWorkbook.Sheets("Principal").Range("L7").Value = "" Or ThisWorkbook.Sheets("Principal").Range("H7").Value = "" Then
        MsgBox "Datos incompletos, por favor ingrese los datos antes de ejecutar.", vbExclamation
        Exit Sub
    End If

    ' Llama a cada una de las funciones
    CrearCarpetas
    ultima_parte
    
    MsgBox "El informe ha sido completado y se encuentra abierto. Por favo ingrese a realizar el analisis  y revisiones correspondientes.", vbInformation

End Sub

Sub InicializarVariables()
'Definicion de las variables
    
    'Fechas
    mes = ThisWorkbook.Sheets("Principal").Range("M7").Text
    Mes_Texto = ThisWorkbook.Sheets("Principal").Range("H11").Value
    año = ThisWorkbook.Sheets("Principal").Range("M13").Value
    Fecha1 = ThisWorkbook.Sheets("Principal").Range("H7").Value
    Fecha2 = ThisWorkbook.Sheets("Principal").Range("I7").Value
    
    'Rutas
    'ruta = ThisWorkbook.Path & "\"
    ruta = "G:\H2R\Mexico\PAYROLL\Conciliacion Cuentas\"
    'Ruta_Año = ruta & "FONACOT"
    Ruta_Año = ruta & año & "\" & "FONACOT"
    Ruta_Mes = Ruta_Año & "\" & Mes_Texto & " " & año
    Ruta_soporte = Ruta_Mes & "\" & "Soportes"
    'ORIGINAL: 'G:\H2R\Mexico\PAYROLL\Conciliacion Cuentas\2024\FONACOT
    

End Sub

Sub CrearCarpetas()
    
    'Creación y validacion de las carpetas
    InicializarVariables
    
    '''''''
    ''AÑO''
    '''''''
    Ruta_Año = ruta & "FONACOT"
    If Dir(Ruta_Año, vbDirectory + vbHidden) = "" Then
        'Comprueba que la carpeta no exista para crear el directorio.
        If Dir(Ruta_Año & vbDirectory + vbHidden) = "" Then MkDir Ruta_Año
    End If
    '''''''
    ''MES''
    '''''''
    Ruta_Mes = Ruta_Año & "\" & Mes_Texto & " " & año
    If Dir(Ruta_Mes, vbDirectory + vbHidden) = "" Then
        'Comprueba que la carpeta no exista para crear el directorio.
        If Dir(Ruta_Mes & vbDirectory + vbHidden) = "" Then MkDir Ruta_Mes
    End If
    ''''''''''''''''''''
    ''SOPORTES''
    ''''''''''''''''''''
    Ruta_soporte = Ruta_Mes & "\" & "Soportes"
    If Dir(Ruta_soporte, vbDirectory + vbHidden) = "" Then
        'Comprueba que la carpeta no exista para crear el directorio.
        If Dir(Ruta_soporte & vbDirectory + vbHidden) = "" Then MkDir Ruta_soporte
    End If
        
End Sub


Sub Consolidado()
    
    InicializarVariables
    Dim wbPlantilla As Workbook
    Dim wbPlantilla_creditos As Workbook
    Dim wbMX02 As Workbook
    Dim wbMX08 As Workbook
    Dim wbMX021 As Workbook
    Dim wbMX081 As Workbook
    Dim lastRowA As Long
    Dim lastRowC As Long
    
    'Elimina una copia si existe
    On Error Resume Next
    Kill Ruta_Mes & "\" & "Consolidado " & Mes_Texto & " " & año & ".xlsx"
    On Error GoTo 0
    
    'Hace copia de la plantilla
    Set wbPlantilla = Workbooks.Open(Ruta_Año & "\" & "Consolidado Plantilla.xlsx")
    wbPlantilla.Activate
    ActiveWorkbook.SaveCopyAs Filename:=Ruta_Mes & "\" & "Consolidado " & Mes_Texto & " " & año & ".xlsx"
    wbPlantilla.Close False
    
    'Abrir archivos
    Set wbConsolidado = Workbooks.Open(Ruta_Mes & "\" & "Consolidado " & Mes_Texto & " " & año & ".xlsx")
    Set wbMX02 = Workbooks.Open(Ruta_Mes & "\" & "MX02.csv")
    Set wbMX021 = Workbooks.Open(Ruta_Mes & "\" & "MX02-1.csv")
    Set wbMX08 = Workbooks.Open(Ruta_Mes & "\" & "MX08.csv")
    Set wbMX081 = Workbooks.Open(Ruta_Mes & "\" & "MX08-1.csv")
    
 '--------------- Crea copia de la plantilla de creditos activos ---------------
    
    'Elimina una copia si existe
    On Error Resume Next
    Kill Ruta_soporte & "\" & "Creditos activos " & Mes_Texto & " " & año & "Inicial " & ".xlsx"
    On Error GoTo 0
    
    'Hace copia de la plantilla
    Set wbPlantilla_creditos = Workbooks.Open(Ruta_Año & "\" & "Creditos activos Plantilla.xlsx")
    wbPlantilla_creditos.Activate
    ActiveWorkbook.SaveCopyAs Filename:=Ruta_soporte & "\" & "Creditos activos " & Mes_Texto & " " & año & " Inicial " & ".xlsx"
    wbPlantilla_creditos.Close False
    
    
'--------------- Copia los datos de MX02 ---------------
    wbMX02.Activate
    Sheets(1).Activate
    lastRowA = ActiveSheet.Cells(Rows.Count, 1).End(xlUp).Row
    Range("A2:M" & lastRowA).Select
    Selection.Copy
    wbConsolidado.Activate
    Sheets("Cedula descargada").Activate
    ActiveSheet.Range("C3").PasteSpecial xlPasteAll
    
    'Coloca el indicativo del archivo
    lastRowC = ActiveSheet.Cells(Rows.Count, 3).End(xlUp).Row
    ActiveSheet.Range("A3:A" & lastRowC).Value = "MX02"
    wbConsolidado.Save
    wbMX02.Close False
    
'--------------- Copia los datos de MX02-1 ---------------
    wbMX021.Activate
    Sheets(1).Activate
    lastRowA = ActiveSheet.Cells(Rows.Count, 1).End(xlUp).Row
    Range("A2:M" & lastRowA).Select
    Selection.Copy
    wbConsolidado.Activate
    Sheets("Cedula descargada").Activate
    lastRowC = ActiveSheet.Cells(Rows.Count, 3).End(xlUp).Row
    ActiveSheet.Range("C" & lastRowC + 1).PasteSpecial xlPasteAll

    'Coloca el indicativo del archivo
    lastRowA = ActiveSheet.Cells(ActiveSheet.Rows.Count, "A").End(xlUp).Row
    lastRowC = ActiveSheet.Cells(ActiveSheet.Rows.Count, "C").End(xlUp).Row
    ActiveSheet.Range("A" & lastRowA + 1 & ":A" & lastRowC).Value = "MX02-1"
    wbConsolidado.Save
    wbMX021.Close False
    
'--------------- Copia los datos de MX08 ---------------
    wbMX08.Activate
    Sheets(1).Activate
    lastRowA = ActiveSheet.Cells(Rows.Count, 1).End(xlUp).Row
    Range("A2:M" & lastRowA).Select
    Selection.Copy
    wbConsolidado.Activate
    Sheets("Cedula descargada").Activate
    lastRowC = ActiveSheet.Cells(Rows.Count, 3).End(xlUp).Row
    ActiveSheet.Range("C" & lastRowC + 1).PasteSpecial xlPasteAll
    
    'Coloca el indicativo del archivo
    lastRowA = ActiveSheet.Cells(ActiveSheet.Rows.Count, "A").End(xlUp).Row
    lastRowC = ActiveSheet.Cells(ActiveSheet.Rows.Count, "C").End(xlUp).Row
    ActiveSheet.Range("A" & lastRowA + 1 & ":A" & lastRowC).Value = "MX08"
    wbConsolidado.Save
    wbMX08.Close False
    
'--------------- Copia los datos de MX08-1 ---------------
    wbMX081.Activate
    Sheets(1).Activate
    lastRowA = ActiveSheet.Cells(Rows.Count, 1).End(xlUp).Row
    Range("A2:M" & lastRowA).Select
    Selection.Copy
    wbConsolidado.Activate
    Sheets("Cedula descargada").Activate
    lastRowC = ActiveSheet.Cells(Rows.Count, 3).End(xlUp).Row
    ActiveSheet.Range("C" & lastRowC + 1).PasteSpecial xlPasteAll
    
    'Coloca el indicativo del archivo
    lastRowA = ActiveSheet.Cells(ActiveSheet.Rows.Count, "A").End(xlUp).Row
    lastRowC = ActiveSheet.Cells(ActiveSheet.Rows.Count, "C").End(xlUp).Row
    ActiveSheet.Range("A" & lastRowA + 1 & ":A" & lastRowC).Value = "MX08-1"
    wbConsolidado.Save
    wbMX081.Close False
    
'--------------- Termina ---------------
    MsgBox "Archivos consolidados exitosamente.", vbInformation
    
End Sub

Sub sap()


'-------------------- DESCARGA ZHRMX27 --------------------
'Variante ACTIVOS FONACO de JULNOVOA

    InicializarVariables
    'Descargar datos
    Dim SapGuiAuto As Object
    Dim App As Object
    Dim Connection As Object
    Dim session As Object
    
    ' Conexion con SAP
    Application.DisplayAlerts = False
    Set SapGuiAuto = GetObject("SAPGUI")
    Set App = SapGuiAuto.GetScriptingEngine
    Set Connection = App.Children(0)
    Set session = Connection.Children(0)
    
    'Vuelve a la pantalla inicial
    session.findById("wnd[0]/tbar[0]/okcd").Text = "/n"
    session.findById("wnd[0]").sendVKey 0

    'Entra a la transacion
    session.findById("wnd[0]/tbar[0]/okcd").Text = "/nZHRMX27"
    session.findById("wnd[0]").sendVKey 0
    
    'Variante
    session.findById("wnd[0]/tbar[1]/btn[17]").press
    session.findById("wnd[1]/usr/txtV-LOW").Text = "ACTIVOS FONACO"
    session.findById("wnd[1]/usr/txtENAME-LOW").Text = "JULNOVOA"
    session.findById("wnd[1]/usr/txtV-LOW").caretPosition = 8
    session.findById("wnd[1]/tbar[0]/btn[8]").press
    
    'Borra el "3" para quitar los activos
    session.findById("wnd[0]/usr/btn%_PNPSTAT2_%_APP_%-VALU_PUSH").press
    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE").Columns.elementAt(1).Selected = True
    session.findById("wnd[1]/tbar[0]/btn[16]").press
    session.findById("wnd[1]/tbar[0]/btn[8]").press
    
    'Ejecuta
    session.findById("wnd[0]/tbar[1]/btn[8]").press
    
    'Descarga
    session.findById("wnd[0]/mbar/menu[0]/menu[3]/menu[2]").Select
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").Select
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").SetFocus
    session.findById("wnd[1]/tbar[0]/btn[0]").press
    session.findById("wnd[1]/usr/ctxtDY_PATH").Text = Ruta_soporte
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = "ZHRMX27.XLS"
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 11
    session.findById("wnd[1]/tbar[0]/btn[11]").press
    
    'Crea copia. pasa de xls a xlsx
    Application.CutCopyMode = False
    Workbooks.Open Ruta_soporte & "\" & "ZHRMX27.XLS"
    ActiveSheet.Cells.Select
    Selection.Copy
    Workbooks.Add
    ActiveSheet.Paste
    ActiveWorkbook.SaveAs Ruta_soporte & "\" & "ZHRMX27.XLSX"
    ActiveWorkbook.Close SaveChanges:=True
    Workbooks("ZHRMX27.XLS").Close
    Kill Ruta_soporte & "\" & "ZHRMX27.XLS"
    
    'Cambios de formato
    Workbooks.Open Ruta_soporte & "\" & "ZHRMX27.XLSX"
    Workbooks("ZHRMX27.XLSX").Activate
    Rows("1:7").Delete
    Rows("2").Delete
    Columns("A").Delete
    Columns("A:O").AutoFit
    
    'Convierte fechas de contab. de 01.01.9999 a 01/01/9999
    Sheets(1).Activate
    Dim CelsC As Range
    Dim UltimaFilaC As Long
    UltimaFilaC = ActiveSheet.Cells(Rows.Count, 1).End(xlUp).Row
    
    If UltimaFilaC >= 2 Then
        For Each CelsC In ActiveSheet.Range("G2:G" & UltimaFilaC)
            If Len(CelsC.Value) = 10 And Mid(CelsC.Value, 3, 1) = "." And Mid(CelsC.Value, 6, 1) = "." Then
                CelsC.Value = DateSerial(Right(CelsC.Value, 4), Mid(CelsC.Value, 4, 2), Left(CelsC.Value, 2))
                CelsC.NumberFormat = "dd/mm/yyyy"
            End If
        Next CelsC
    End If

    ActiveWorkbook.Save


'-------------------- DESCARGA QUERY DE LA SQ01 --------------------
'Query JULNOVOA_FONAC en el H6

    'Vuelve a la pantalla inicial
    session.findById("wnd[0]/tbar[0]/okcd").Text = "/n"
    session.findById("wnd[0]").sendVKey 0
    
    'Entra a la transacion SQ01
    session.findById("wnd[0]/tbar[0]/okcd").Text = "SQ01"
    session.findById("wnd[0]").sendVKey 0
    
    'Busca el grupo
    session.findById("wnd[0]/tbar[1]/btn[19]").press
    session.findById("wnd[1]/usr/cntlGRID1/shellcont/shell").currentCellRow = -1
    session.findById("wnd[1]/usr/cntlGRID1/shellcont/shell").selectColumn "DBGBNUM"
    session.findById("wnd[1]/tbar[0]/btn[29]").press
    session.findById("wnd[2]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN001-LOW").Text = "/SAPQUERY/H6"
    session.findById("wnd[2]/tbar[0]/btn[0]").press
    session.findById("wnd[1]/usr/cntlGRID1/shellcont/shell").selectedRows = "0"
    session.findById("wnd[1]/tbar[0]/btn[0]").press
    
    'Ingresa el query IT377
    session.findById("wnd[0]/usr/ctxtRS38R-QNUM").Text = "JULNOVOA_FONAC"
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]/tbar[1]/btn[8]").press
    
    'Ejecuta
    session.findById("wnd[0]/tbar[1]/btn[8]").press
    
    'Descarga en formato XLS
    session.findById("wnd[0]/mbar/menu[0]/menu[4]/menu[2]").Select
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").Select
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").SetFocus
    session.findById("wnd[1]/tbar[0]/btn[0]").press
    session.findById("wnd[1]/usr/ctxtDY_PATH").Text = Ruta_soporte
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = "Creditos activos.XLS"
    session.findById("wnd[1]/tbar[0]/btn[11]").press
    
    'Cambia formato de xls a xlsx
    Workbooks.Open Ruta_soporte & "\" & "Creditos activos.XLS"
    ActiveSheet.Cells.Select
    Selection.Copy
    Workbooks.Add
    ActiveSheet.Paste
    ActiveWorkbook.SaveAs Ruta_soporte & "\" & "Creditos activos.XLSX"
    ActiveWorkbook.Close SaveChanges:=True
    Workbooks("Creditos activos.XLS").Close
    Kill Ruta_soporte & "\" & "Creditos activos.XLS"
    Workbooks.Open Ruta_soporte & "\" & "Creditos activos.XLSX"
    Workbooks("Creditos activos.XLSX").Activate

    'Cambios de formato
    Rows("1:4").Delete
    Rows("2").Delete
    Columns("A").Delete
    Columns("C").Delete
    Columns("A:Q").AutoFit
    
    'Cambia formatos de fecha
    Sheets(1).Activate
    lastRow = ActiveSheet.Cells(Rows.Count, 1).End(xlUp).Row
    
    If lastRow >= 2 Then
        For Each CelsFecha In ActiveSheet.Range("C2:C" & lastRow)
            If Len(CelsFecha.Value) = 10 And Mid(CelsFecha.Value, 3, 1) = "." And Mid(CelsFecha.Value, 6, 1) = "." Then
                CelsFecha.Value = DateSerial(Right(CelsFecha.Value, 4), Mid(CelsFecha.Value, 4, 2), Left(CelsFecha.Value, 2))
                CelsFecha.NumberFormat = "dd/mm/yyyy"
            End If
        Next CelsFecha
    End If
    
    If lastRow >= 2 Then
        For Each CelsFecha In ActiveSheet.Range("H2:H" & lastRow)
            If Len(CelsFecha.Value) = 10 And Mid(CelsFecha.Value, 3, 1) = "." And Mid(CelsFecha.Value, 6, 1) = "." Then
                CelsFecha.Value = DateSerial(Right(CelsFecha.Value, 4), Mid(CelsFecha.Value, 4, 2), Left(CelsFecha.Value, 2))
                CelsFecha.NumberFormat = "dd/mm/yyyy"
            End If
        Next CelsFecha
    End If
    
    If lastRow >= 2 Then
        For Each CelsFecha In ActiveSheet.Range("I2:I" & lastRow)
            If Len(CelsFecha.Value) = 10 And Mid(CelsFecha.Value, 3, 1) = "." And Mid(CelsFecha.Value, 6, 1) = "." Then
                CelsFecha.Value = DateSerial(Right(CelsFecha.Value, 4), Mid(CelsFecha.Value, 4, 2), Left(CelsFecha.Value, 2))
                CelsFecha.NumberFormat = "dd/mm/yyyy"
            End If
        Next CelsFecha
    End If
    
    Workbooks("Creditos activos.XLSX").Save
    


    
End Sub

Sub Fanacot()

    InicializarVariables

'-------------- Hace formulas de en el archivo del consolidado --------------
    Set wbConsolidado = Workbooks("Consolidado " & Mes_Texto & " " & año & ".xlsx")
    Workbooks("ZHRMX27.XLSX").Activate
    Sheets(1).Activate
    Columns("A:O").Select
    Selection.Copy
    wbConsolidado.Activate
    Sheets("Base empleados").Activate
    ActiveSheet.Range("A1").PasteSpecial Paste:=xlPasteValues
    Columns("A:O").AutoFit
    Application.CutCopyMode = False
    Workbooks("ZHRMX27.XLSX").Close False
    wbConsolidado.Save
    
    'Formulas
    wbConsolidado.Activate
    Sheets("Cedula descargada").Activate
    Range("A2").Value = "Sociedad"
    lastRow = Cells(Rows.Count, "C").End(xlUp).Row
    Range("B3").AutoFill Destination:=Range("B3:B" & lastRow), Type:=xlFillDefault
    Range("P3").AutoFill Destination:=Range("P3:P" & lastRow), Type:=xlFillDefault
    Range("Q3").AutoFill Destination:=Range("Q3:Q" & lastRow), Type:=xlFillDefault
    Range("R3").AutoFill Destination:=Range("R3:R" & lastRow), Type:=xlFillDefault
    Range("S3").AutoFill Destination:=Range("S3:S" & lastRow), Type:=xlFillDefault
    Range("T3").AutoFill Destination:=Range("T3:T" & lastRow), Type:=xlFillDefault
    Range("U3").AutoFill Destination:=Range("U3:U" & lastRow), Type:=xlFillDefault
    Range("W3").AutoFill Destination:=Range("W3:W" & lastRow), Type:=xlFillDefault
    Range("Y3").AutoFill Destination:=Range("Y3:Y" & lastRow), Type:=xlFillDefault
    Range("Z3").AutoFill Destination:=Range("Z3:Z" & lastRow), Type:=xlFillDefault
    Range("AA3").AutoFill Destination:=Range("AA3:AA" & lastRow), Type:=xlFillDefault
    
    
    'Cambia formato de la columna C
    

'-------------- Hace formulas de en el archivo de creditos iniciales --------------
    Set wbCreditosInicial = Workbooks.Open(Ruta_soporte & "\" & "Creditos activos " & Mes_Texto & " " & año & " Inicial " & ".xlsx")
    Set wbCreditosInicial = Workbooks("Creditos activos " & Mes_Texto & " " & año & " Inicial " & ".xlsx")
    
    Workbooks("Creditos activos.XLSX").Activate
    Sheets(1).Activate
    Columns("A:P").Select
    Selection.Copy
    wbCreditosInicial.Activate
    Sheets(1).Activate
    ActiveSheet.Range("A1").PasteSpecial Paste:=xlPasteValues
    Columns("A:O").AutoFit
    Application.CutCopyMode = False
    Workbooks("Creditos activos.XLSX").Close False
    wbCreditosInicial.Save
    
    
    'Formulas
    wbCreditosInicial.Activate
    Sheets(1).Activate
    lastRow = Cells(Rows.Count, "A").End(xlUp).Row
    Range("Q2").AutoFill Destination:=Range("Q2:Q" & lastRow), Type:=xlFillDefault
    Range("R2").AutoFill Destination:=Range("R2:R" & lastRow), Type:=xlFillDefault
    Range("S2").AutoFill Destination:=Range("S2:S" & lastRow), Type:=xlFillDefault
    Range("T2").AutoFill Destination:=Range("T2:T" & lastRow), Type:=xlFillDefault
    Range("U2").AutoFill Destination:=Range("U2:U" & lastRow), Type:=xlFillDefault
    Range("V2").AutoFill Destination:=Range("V2:V" & lastRow), Type:=xlFillDefault
    Range("Z2").AutoFill Destination:=Range("Z2:Z" & lastRow), Type:=xlFillDefault
    Range("AA2").AutoFill Destination:=Range("AA2:AA" & lastRow), Type:=xlFillDefault
    Range("AB2").AutoFill Destination:=Range("AB2:AB" & lastRow), Type:=xlFillDefault
    
    'Fechas
    Range("AJ1").Value = "01" & "/" & mes & "/" & año
    Range("AJ2").Value = "01" & "/" & mes & "/" & año
    Columns("AJ:AJ").NumberFormat = "DD/MM/YYYY"
    Columns("H:H").NumberFormat = "DD/MM/YYYY"
    Columns("I:I").NumberFormat = "DD/MM/YYYY"
    Columns("Q:Q").NumberFormat = "DD/MM/YYYY"
   
'-------------- Crear la hoja Batch --------------

    wbConsolidado.Activate
    lastRow = Sheets("Cedula descargada").Cells(Rows.Count, "B").End(xlUp).Row
    
    'Sociedad
    Sheets("Cedula descargada").Range("A3:A" & lastRow).Copy
    Sheets("Modelo Batch").Range("A2").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    
    'Tipo empleado
    Sheets("Cedula descargada").Range("Q3:Q" & lastRow).Copy
    Sheets("Modelo Batch").Range("B2").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    
    'N empleado
    Sheets("Cedula descargada").Range("B3:B" & lastRow).Copy
    Sheets("Modelo Batch").Range("C2").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    
    'Nombre
    Sheets("Cedula descargada").Range("E3:E" & lastRow).Copy
    Sheets("Modelo Batch").Range("D2").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    
    'N Fonacot
    Sheets("Cedula descargada").Range("C3:C" & lastRow).Copy
    Sheets("Modelo Batch").Range("E2").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    
    'N Credito
    Sheets("Cedula descargada").Range("F3:F" & lastRow).Copy
    Sheets("Modelo Batch").Range("F2").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    
    'Plazo
    Sheets("Cedula descargada").Range("I3:I" & lastRow).Copy
    Sheets("Modelo Batch").Range("G2").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    
    'Cuotas Pagadas
    Sheets("Cedula descargada").Range("J3:J" & lastRow).Copy
    Sheets("Modelo Batch").Range("H2").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    
    'Retencion real
    Sheets("Cedula descargada").Range("K3:K" & lastRow).Copy
    Sheets("Modelo Batch").Range("I2").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    
    wbConsolidado.Save
    
    'Alarga las formulas
    Sheets("Modelo Batch").Activate
    lastRow = Cells(Rows.Count, "A").End(xlUp).Row
    Range("K2").AutoFill Destination:=Range("K2:K" & lastRow), Type:=xlFillDefault
    Range("L2").AutoFill Destination:=Range("L2:L" & lastRow), Type:=xlFillDefault
    Range("N2").AutoFill Destination:=Range("N2:N" & lastRow), Type:=xlFillDefault
    Range("O2").AutoFill Destination:=Range("O2:O" & lastRow), Type:=xlFillDefault
    Range("P2").AutoFill Destination:=Range("P2:P" & lastRow), Type:=xlFillDefault
    Range("Q2").AutoFill Destination:=Range("Q2:Q" & lastRow), Type:=xlFillDefault
    Range("R2").AutoFill Destination:=Range("R2:R" & lastRow), Type:=xlFillDefault
    Range("S2").AutoFill Destination:=Range("S2:S" & lastRow), Type:=xlFillDefault
    Range("T2").AutoFill Destination:=Range("T2:T" & lastRow), Type:=xlFillDefault
    Range("U2").AutoFill Destination:=Range("U2:U" & lastRow), Type:=xlFillDefault
    Range("V2").AutoFill Destination:=Range("V2:V" & lastRow), Type:=xlFillDefault
    
    'Guardar y cerrar
    wbConsolidado.Save
    wbCreditosInicial.Save
    wbConsolidado.Close
    wbCreditosInicial.Close
     
     
'-------------- Va y trabaja en el archivo del analisis del consolidado --------------
    
    
End Sub

'--------------------------------------
'-------------------------------------- INICIO PARTE DOS
'--------------------------------------

Sub Analisis_consolidado()
'Abre formulario para descargar los datos de las nominas mensual, semanal y catorcenal en la 025 con las variantes de julian
'Luego organiza documentos descargados

    InicializarVariables
    
    ' Abrir el formulario, en esta parte entra a realizar todas las descargas de sap restantes
    MsgBox "Por favor ingrese los datos que se le piden en el siguiente formulario.", vbInformation
    UserForm1.Show

    Set wbConsolidado = Workbooks.Open(Ruta_Mes & "\" & "Consolidado " & Mes_Texto & " " & año & ".xlsx")
    Set wbConsolidado = Workbooks("Consolidado " & Mes_Texto & " " & año & ".xlsx")

'--------- ORGANIZA EL REPORTE DE LAS MENSUALES ---------
     InicializarVariables
    'Organiza documentos
    Set wbMensual = Workbooks.Open(Ruta_soporte & "\" & "ZPYMX025 FONACOT MES.XLSX")
    Set wbMensual = Workbooks("ZPYMX025 FONACOT MES.XLSX")
    wbMensual.Activate
    Sheets(1).Activate
    Rows("1").Delete
    Rows("2").Delete
    Columns("A").Delete
    Columns("A:Z").AutoFit
    
    'Organiza fechas
    Dim Cels As Range
    Dim UltimaFila As Long
    UltimaFila = ActiveSheet.Cells(Rows.Count, 1).End(xlUp).Row
    If UltimaFila >= 2 Then

        For Each Cels In ActiveSheet.Range("Q2:Q" & UltimaFila)
            If Len(Cels.Value) = 10 And Mid(Cels.Value, 3, 1) = "." And Mid(Cels.Value, 6, 1) = "." Then
                Cels.Value = DateSerial(Right(Cels.Value, 4), Mid(Cels.Value, 4, 2), Left(Cels.Value, 2))
                Cels.NumberFormat = "dd/mm/yyyy"
            End If
        Next Cels
    
    End If
    
    'Cambios de formato para la cantidad
    Columns("W:W").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("W1").Value = "Cantidad"
    lastRow = ActiveSheet.Cells(ActiveSheet.Rows.Count, "A").End(xlUp).Row
    If lastRow >= 2 Then
        Range("W2:W" & lastRow).Formula = "=VALUE(SUBSTITUTE(RC[1], CHAR(160),""""))"
    End If
    Columns("W:W").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("X:X").Select
    Selection.Delete

    'Cambios de formato para el importe
    Columns("X:X").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("X1").Value = "Importe"
        If lastRow >= 2 Then
        Range("X2:X" & lastRow).Formula = "=VALUE(SUBSTITUTE(RC[1], CHAR(160),""""))"
    End If
    Columns("X:X").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("Y:Y").Select
    Selection.Delete
    Columns("X:X").NumberFormat = "$#,##0.00"
    Columns("A:Z").AutoFit

    'Elimina filas con importe 0
    lastRow = ActiveSheet.Cells(ActiveSheet.Rows.Count, "X").End(xlUp).Row
    For i = lastRow To 2 Step -1
        If ActiveSheet.Cells(i, "X").Value = 0 Then
            ActiveSheet.Rows(i).EntireRow.Delete
        End If
    Next i
    ActiveWorkbook.Save
    

'--------- ORGANIZA EL REPORTE DE LAS CATORCENALES ---------

    'Organiza documentos
    Set wbCatorcenal = Workbooks.Open(Ruta_soporte & "\" & "ZPYMX025 FONACOT CATORCENA.XLSX")
    Set wbCatorcenal = Workbooks("ZPYMX025 FONACOT CATORCENA.XLSX")
    wbCatorcenal.Activate
    Sheets(1).Activate
    Rows("1").Delete
    Rows("2").Delete
    Columns("A").Delete
    Columns("A:Z").AutoFit
    
    'Organiza fechas
    UltimaFila = ActiveSheet.Cells(Rows.Count, 1).End(xlUp).Row
    If UltimaFila >= 2 Then

        For Each Cels In ActiveSheet.Range("Q2:Q" & UltimaFila)
            If Len(Cels.Value) = 10 And Mid(Cels.Value, 3, 1) = "." And Mid(Cels.Value, 6, 1) = "." Then
                Cels.Value = DateSerial(Right(Cels.Value, 4), Mid(Cels.Value, 4, 2), Left(Cels.Value, 2))
                Cels.NumberFormat = "dd/mm/yyyy"
            End If
        Next Cels
    
    End If
    
    'Cambios de formato para la cantidad
    Columns("W:W").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("W1").Value = "Cantidad"
    lastRow = ActiveSheet.Cells(ActiveSheet.Rows.Count, "A").End(xlUp).Row
    If lastRow >= 2 Then
        Range("W2:W" & lastRow).Formula = "=VALUE(SUBSTITUTE(RC[1], CHAR(160),""""))"
    End If
    Columns("W:W").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("X:X").Select
    Selection.Delete

    'Cambios de formato para el importe
    Columns("X:X").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("X1").Value = "Importe"
        If lastRow >= 2 Then
        Range("X2:X" & lastRow).Formula = "=VALUE(SUBSTITUTE(RC[1], CHAR(160),""""))"
    End If
    Columns("X:X").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("Y:Y").Select
    Selection.Delete
    Columns("X:X").NumberFormat = "$#,##0.00"
    Columns("A:Z").AutoFit

    'Elimina filas con importe 0
    lastRow = ActiveSheet.Cells(ActiveSheet.Rows.Count, "X").End(xlUp).Row
    For i = lastRow To 2 Step -1
        If ActiveSheet.Cells(i, "X").Value = 0 Then
            ActiveSheet.Rows(i).EntireRow.Delete
        End If
    Next i
    ActiveWorkbook.Save


'--------- ORGANIZA EL REPORTE DE LAS SEMANALES ---------

    'Organiza documentos
    Set wbSemanal = Workbooks.Open(Ruta_soporte & "\" & "ZPYMX025 FONACOT SEMANA.XLSX")
    Set wbSemanal = Workbooks("ZPYMX025 FONACOT SEMANA.XLSX")
    wbSemanal.Activate
    Sheets(1).Activate
    Rows("1").Delete
    Rows("2").Delete
    Columns("A").Delete
    Columns("A:Z").AutoFit
    
    'Organiza fechas
    UltimaFila = ActiveSheet.Cells(Rows.Count, 1).End(xlUp).Row
    If UltimaFila >= 2 Then

        For Each Cels In ActiveSheet.Range("Q2:Q" & UltimaFila)
            If Len(Cels.Value) = 10 And Mid(Cels.Value, 3, 1) = "." And Mid(Cels.Value, 6, 1) = "." Then
                Cels.Value = DateSerial(Right(Cels.Value, 4), Mid(Cels.Value, 4, 2), Left(Cels.Value, 2))
                Cels.NumberFormat = "dd/mm/yyyy"
            End If
        Next Cels
    
    End If
    
    'Cambios de formato para la cantidad
    Columns("W:W").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("W1").Value = "Cantidad"
    lastRow = ActiveSheet.Cells(ActiveSheet.Rows.Count, "A").End(xlUp).Row
    If lastRow >= 2 Then
        Range("W2:W" & lastRow).Formula = "=VALUE(SUBSTITUTE(RC[1], CHAR(160),""""))"
    End If
    Columns("W:W").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("X:X").Select
    Selection.Delete

    'Cambios de formato para el importe
    Columns("X:X").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("X1").Value = "Importe"
        If lastRow >= 2 Then
        Range("X2:X" & lastRow).Formula = "=VALUE(SUBSTITUTE(RC[1], CHAR(160),""""))"
    End If
    Columns("X:X").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("Y:Y").Select
    Selection.Delete
    Columns("X:X").NumberFormat = "$#,##0.00"
    Columns("A:Z").AutoFit

    'Elimina filas con importe 0
    lastRow = ActiveSheet.Cells(ActiveSheet.Rows.Count, "X").End(xlUp).Row
    For i = lastRow To 2 Step -1
        If ActiveSheet.Cells(i, "X").Value = 0 Then
            ActiveSheet.Rows(i).EntireRow.Delete
        End If
    Next i
    ActiveWorkbook.Save

'--------- Pedirle al usuario que abra el analisis del consolidado del mes pasado ---------

    Dim archivoSeleccionado As Variant
    Dim nombre_anterior As String
    Dim archivoExistente As String
    
    
    ' Solicitar al usuario que seleccione el archivo de análisis del mes anterior
    archivoSeleccionado = Application.GetOpenFilename("Archivos de Excel (*.xlsx; *.xlsb), *.xlsx; *.xlsb", , "Selecciona el archivo de análisis del mes anterior")
    
    ' Salir si el usuario cancela la selección
    If archivoSeleccionado = False Then Exit Sub
    
    ' Obtener el nombre del archivo seleccionado
    nombre_anterior = Dir(archivoSeleccionado)
    
    archivoExistente = Dir(Ruta_Mes & "\" & "Consolidado Analisis Fonacot " & Mes_Texto & " " & año & ".xlsb")

    ' Eliminar el archivo existente si se encuentra
    If archivoExistente <> "" Then
        Kill Ruta_Mes & "\" & "Consolidado Analisis Fonacot " & Mes_Texto & " " & año & ".xlsb"
    End If
    
    ' Crear una copia del archivo en la ruta de soporte
    FileCopy archivoSeleccionado, Ruta_Mes & "\" & nombre_anterior
    
    ' Renombrar la copia como "consolidado_analisis_mes_actual.xlsx" '------ corregir a que si ya existe borrarlo
    Name Ruta_Mes & "\" & nombre_anterior As Ruta_Mes & "\" & "Consolidado Analisis Fonacot " & Mes_Texto & " " & año & ".xlsb"
    
    'Abre archivo del analisis
    Set wbAnalisis = Workbooks.Open(Ruta_Mes & "\" & "Consolidado Analisis Fonacot " & Mes_Texto & " " & año & ".xlsb")
    Set wbAnalisis = Workbooks("Consolidado Analisis Fonacot " & Mes_Texto & " " & año & ".xlsb")

    
    'Copia y pega los datos de los archivos descargados al archivo de analisis
    'Mensual
    wbMensual.Activate
    Sheets(1).Activate
    lastRow = ActiveSheet.Cells(Rows.Count, 1).End(xlUp).Row
    Range("A2:Z" & lastRow).Select
    Selection.Copy
    wbAnalisis.Activate
    Sheets("CC Nom").Activate
    lastRow = ActiveSheet.Cells(Rows.Count, 1).End(xlUp).Row
    ActiveSheet.Range("A" & lastRow + 1).PasteSpecial xlPasteAll
    Application.CutCopyMode = False
    wbMensual.Save
    wbMensual.Close
    wbAnalisis.Save
    
    'Catorcenal
    wbCatorcenal.Activate
    Sheets(1).Activate
    lastRow = ActiveSheet.Cells(Rows.Count, 1).End(xlUp).Row
    Range("A2:Z" & lastRow).Select
    Selection.Copy
    wbAnalisis.Activate
    Sheets("CC Nom").Activate
    lastRow = ActiveSheet.Cells(Rows.Count, 1).End(xlUp).Row
    ActiveSheet.Range("A" & lastRow + 1).PasteSpecial xlPasteAll
    Application.CutCopyMode = False
    wbCatorcenal.Save
    wbCatorcenal.Close
    wbAnalisis.Save
    
    'Semanal
    wbSemanal.Activate
    Sheets(1).Activate
    lastRow = ActiveSheet.Cells(Rows.Count, 1).End(xlUp).Row
    Range("A2:Z" & lastRow).Select
    Selection.Copy
    wbAnalisis.Activate
    Sheets("CC Nom").Activate
    lastRow = ActiveSheet.Cells(Rows.Count, 1).End(xlUp).Row
    ActiveSheet.Range("A" & lastRow + 1).PasteSpecial xlPasteAll
    Application.CutCopyMode = False
    wbSemanal.Save
    wbSemanal.Close
    wbAnalisis.Save
    wbAnalisis.Activate
    Sheets("CC Nom").Activate


'--------- Pega las cedulas del archivo de consolidados al archivo de analisis ---------

    Dim lastRowA, lastRowC As Long
    wbConsolidado.Activate
    Sheets("Cedula descargada").Activate
    If ActiveSheet.FilterMode Then
        ActiveSheet.ShowAllData
    End If
    lastRowA = ActiveSheet.Cells(Rows.Count, 1).End(xlUp).Row
    Range("A3:T" & lastRowA).Select
    Selection.Copy
    wbAnalisis.Activate
    Sheets("Cedulas").Activate
    lastRowC = ActiveSheet.Cells(Rows.Count, 3).End(xlUp).Row
    ActiveSheet.Range("C" & lastRowC + 1).PasteSpecial xlPasteValues
    lastRowA = ActiveSheet.Cells(Rows.Count, 2).End(xlUp).Row
    
    wbAnalisis.Save
    wbConsolidado.Save
    wbConsolidado.Close
    
    wbAnalisis.Activate
    Sheets("Cedulas").Activate
    If ActiveSheet.FilterMode Then
        ActiveSheet.ShowAllData
    End If
    lastRow = ActiveSheet.Cells(ActiveSheet.Rows.Count, "A").End(xlUp).Row
    ActiveSheet.Range("A1:Z" & lastRow).AutoFilter Field:=23, Criteria1:="Pegar en papel de trabajo"
    wbAnalisis.Save

End Sub

Sub ultima_parte()

    InicializarVariables
'--------- Pega los datos que tengan "papel de trabajo" ---------
    Set wbConsolidado = Workbooks("Consolidado Analisis Fonacot " & Mes_Texto & " " & año & ".xlsb")
'    wbAnalisis.Activate
'    Sheets("Cedulas").Activate
'    If ActiveSheet.FilterMode Then
'        ActiveSheet.ShowAllData
'    End If
'    lastRow = ActiveSheet.Cells(ActiveSheet.Rows.Count, "A").End(xlUp).Row
'    ActiveSheet.Range("A1:Z" & lastRow).AutoFilter Field:=23, Criteria1:="Pegar en papel de trabajo"
'    wbAnalisis.Save
'    lastRow = ActiveSheet.Cells(Rows.Count, 1).End(xlUp).Row
'    Range("D2:D" & lastRow).Select
'    Selection.SpecialCells(xlCellTypeVisible).Select
'    Selection.Copy
'    Sheets("Papel de trabajo").Activate
'    If ActiveSheet.FilterMode Then
'        ActiveSheet.ShowAllData
'    End If
'    lastRow = ActiveSheet.Cells(Rows.Count, 1).End(xlUp).Row
'    ActiveSheet.Range("A" & lastRow + 1).PasteSpecial xlPasteValues
'
'    wbAnalisis.Save


'--------- Descarga maestro de sap ---------

    Dim SapGuiAuto As Object
    Dim App As Object
    Dim Connection As Object
    Dim session As Object

    ' Conexion con SAP
    Application.DisplayAlerts = False
    Set SapGuiAuto = GetObject("SAPGUI")
    Set App = SapGuiAuto.GetScriptingEngine
    Set Connection = App.Children(0)
    Set session = Connection.Children(0)

    'Vuelve a la pantalla inicial
    session.findById("wnd[0]/tbar[0]/okcd").Text = "/n"
    session.findById("wnd[0]").sendVKey 0

    'Entra a la transacion
    session.findById("wnd[0]/tbar[0]/okcd").Text = "/nZHRMX27"
    session.findById("wnd[0]").sendVKey 0

    'Variante
    session.findById("wnd[0]/tbar[1]/btn[17]").press
    session.findById("wnd[1]/usr/txtV-LOW").Text = "ACTIVOS FONACOO"
    session.findById("wnd[1]/usr/txtENAME-LOW").Text = "JULNOVOA"
    session.findById("wnd[1]/usr/txtV-LOW").caretPosition = 8
    session.findById("wnd[1]/tbar[0]/btn[8]").press

    'Borra el "3" para quitar los activos
    session.findById("wnd[0]/usr/btn%_PNPSTAT2_%_APP_%-VALU_PUSH").press
    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE").Columns.elementAt(1).Selected = True
    session.findById("wnd[1]/tbar[0]/btn[16]").press
    session.findById("wnd[1]/tbar[0]/btn[8]").press

    'Selecciona que sea hoy la fecha
    session.findById("wnd[0]/usr/radPNPTIMR1").Select
    session.findById("wnd[0]/usr/radPNPTIMR1").SetFocus
    session.findById("wnd[0]/usr/btn%_PNPPERNR_%_APP_%-VALU_PUSH").press

    'Coloca los numeros del los empleados
    session.findById("wnd[0]/usr/ctxtPNPENDDA").SetFocus
    session.findById("wnd[0]/usr/ctxtPNPENDDA").caretPosition = 0
    session.findById("wnd[0]/usr/btn%_PNPPERNR_%_APP_%-VALU_PUSH").press
    session.findById("wnd[1]/tbar[0]/btn[16]").press

    wbAnalisis.Activate
    Sheets("Papel de trabajo").Activate
    If ActiveSheet.FilterMode Then
        ActiveSheet.ShowAllData
    End If
    Range("A4").Select
    lastRow = Cells(Rows.Count, "A").End(xlUp).Row

    If lastRow >= 3 Then
            Columns("A").Resize(lastRow - 1).Offset(2, 0).Copy
        Else
            MsgBox "No hay datos en la columna A.", vbExclamation
    End If

    session.findById("wnd[1]/tbar[0]/btn[24]").press
    session.findById("wnd[1]/tbar[0]/btn[8]").press

    'Ejecuta
    session.findById("wnd[0]/tbar[1]/btn[8]").press

    'Descarga
    session.findById("wnd[0]/mbar/menu[0]/menu[3]/menu[2]").Select
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").Select
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").SetFocus
    session.findById("wnd[1]/tbar[0]/btn[0]").press
    session.findById("wnd[1]/usr/ctxtDY_PATH").Text = Ruta_soporte
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = "ZHRMX27_MAESTRO.XLS"
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 11
    session.findById("wnd[1]/tbar[0]/btn[11]").press

    'Crea copia. pasa de xls a xlsx
    Application.CutCopyMode = False
    Workbooks.Open Ruta_soporte & "\" & "ZHRMX27_MAESTRO.XLS"
    ActiveSheet.Cells.Select
    Selection.Copy
    Workbooks.Add
    ActiveSheet.Paste
    ActiveWorkbook.SaveAs Ruta_soporte & "\" & "ZHRMX27_MAESTRO.XLSX"
    ActiveWorkbook.Close SaveChanges:=True
    Workbooks("ZHRMX27_MAESTRO.XLS").Close
    Kill Ruta_soporte & "\" & "ZHRMX27_MAESTRO.XLS"

    'Cambios de formato
    Workbooks.Open Ruta_soporte & "\" & "ZHRMX27_MAESTRO.XLSX"
    Workbooks("ZHRMX27_MAESTRO.XLSX").Activate
    Rows("1:7").Delete
    Rows("2").Delete
    Columns("A").Delete
    Columns("A:O").AutoFit
    Columns("K:O").Delete
    Columns("A").Delete
    Columns("B").Delete
    Columns("C:F").Delete
    Columns("A:E").AutoFit
    Workbooks("ZHRMX27_MAESTRO.XLSX").Save

    'copia  y pega en el archivo de analisis
    Workbooks("ZHRMX27_MAESTRO.XLSX").Activate
    Sheets(1).Activate
    lastRow = Sheets(1).Cells(Sheets(1).Rows.Count, "A").End(xlUp).Row
    Sheets(1).Range("A2:A" & lastRow).Copy
    wbAnalisis.Activate
    Sheets("Datos").Range("O3").PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False

    Workbooks("ZHRMX27_MAESTRO.XLSX").Activate
    Sheets(1).Activate
    lastRow = Sheets(1).Cells(Sheets(1).Rows.Count, "A").End(xlUp).Row
    Sheets(1).Range("B2:B" & lastRow).Copy
    wbAnalisis.Activate
    Sheets("Datos").Range("Q3").PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False

    Workbooks("ZHRMX27_MAESTRO.XLSX").Activate
    Sheets(1).Activate
    lastRow = Sheets(1).Cells(Sheets(1).Rows.Count, "A").End(xlUp).Row
    Sheets(1).Range("C2:C" & lastRow).Copy
    wbAnalisis.Activate
    Sheets("Datos").Range("P3").PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False

    Workbooks("ZHRMX27_MAESTRO.XLSX").Activate
    Sheets(1).Activate
    lastRow = Sheets(1).Cells(Sheets(1).Rows.Count, "A").End(xlUp).Row
    Sheets(1).Range("D2:D" & lastRow).Copy
    wbAnalisis.Activate
    Sheets("Datos").Range("R3").PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    Workbooks("ZHRMX27_MAESTRO.XLSX").Save
    Workbooks("ZHRMX27_MAESTRO.XLSX").Close
    wbAnalisis.Save
    

End Sub

