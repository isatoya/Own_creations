Attribute VB_Name = "Módulo1"
Option Explicit
'variables para todo el proyecto

    Dim MAESTRO, wbPlantilla, wbAuditoria As Workbook
    Dim Fecha1, Fecha2, Trimestre, año, fechaActual, fechaFormato, Trimestre_letra, cellValue, incio_trimestre, infinito As String
    Dim ruta, Ruta_Año, Ruta_Trimestre, ruta_maestroA As String
    Dim lastRow, i As Long
    Dim CelsAG As Object
    Dim rng, row, visibleRange, cell, rowToColor As Range
    
Sub InicializarVariables()
'Definicion de las variables
    
    '------------ Variables originales del archivo -----------
    Trimestre = ThisWorkbook.Sheets("Principal").Range("E5").Value
    año = ThisWorkbook.Sheets("Principal").Range("E7").Value
    Trimestre_letra = ThisWorkbook.Sheets("Principal").Range("F5").Text
    Fecha1 = ThisWorkbook.Sheets("Principal").Range("H5").Value
    Fecha2 = ThisWorkbook.Sheets("Principal").Range("H6").Value
    fechaActual = Date
    fechaFormato = Format(fechaActual, "dd.mm.yyyy")
    
    '------------ Rutas -----------
    ruta = ThisWorkbook.Path & "\"
    Ruta_Año = ruta & año
    Ruta_Trimestre = Ruta_Año & "\" & Trimestre
    

End Sub

Sub Macro_control_trimestral()

    ' Verificar si hay datos en las celdas I8 y M8
    If ThisWorkbook.Sheets("Principal").Range("E5").Value = "" Or ThisWorkbook.Sheets("Principal").Range("E7").Value = "" Then
        MsgBox "Datos incompletos, por favor ingrese los datos antes de ejecutar.", vbExclamation
        Exit Sub
    End If

    ' Llama a cada una de las funciones
    CrearCarpetas
    Plantilla
    ZBC033
    ZHR929
    ZBC033_Criticas
    Desarrollo_organizar
    
    MsgBox "Reporte finalizado. Este se encuentra abierto en una de las ventanas de excel", vbInformation

End Sub

Sub CrearCarpetas()
    
    InicializarVariables

    '''''''
    ''AÑO''
    '''''''
    Ruta_Año = ruta & año
    If Dir(Ruta_Año, vbDirectory + vbHidden) = "" Then
        'Comprueba que la carpeta no exista para crear el directorio.
        If Dir(Ruta_Año & vbDirectory + vbHidden) = "" Then MkDir Ruta_Año
    End If
    
    '''''''''''''
    ''TRIMESTRE''
    '''''''''''''
    Ruta_Trimestre = Ruta_Año & "\" & Trimestre
    If Dir(Ruta_Trimestre, vbDirectory + vbHidden) = "" Then
        'Comprueba que la carpeta no exista para crear el directorio.
        If Dir(Ruta_Trimestre & vbDirectory + vbHidden) = "" Then MkDir Ruta_Trimestre
    End If
        
End Sub


Sub ZBC033()
InicializarVariables

'-------------------- DESCARGA LOS DATOS DE LA ZBC033 PARA LA REVISION DE CARGOS --------------------
    
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
    session.findById("wnd[0]/tbar[0]/okcd").Text = "/nZBC033"
    session.findById("wnd[0]").sendVKey 0
    
    'Entra a la opcion de cargos
    session.findById("wnd[0]/tbar[1]/btn[8]").press
    session.findById("wnd[0]/usr/btn%_S_USER_%_APP_%-VALU_PUSH").press
    
    'Borra si hay datos
    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE").Columns.elementAt(1).Selected = True
    session.findById("wnd[1]/tbar[0]/btn[16]").press
    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE").deselectAllColumns
    
    'Le pide al usuario abrir el informe de conflictos para copiar los datos de los usuarios
    MsgBox "Por favor selecciene el archivo de los conflictos descargado desde GRC.", vbInformation
    ruta_maestroA = Application.GetOpenFilename("Archivos Excel (*.xls; *.xlsx), *.xls; *.xlsx")
    Application.AskToUpdateLinks = False
    
    If ruta_maestroA <> "Falso" Then
            
            Set MAESTRO = Workbooks.Open(Filename:=ruta_maestroA, UpdateLinks:=0)
            MAESTRO.Activate
            Worksheets(1).Activate
            lastRow = ActiveSheet.Cells(ActiveSheet.Rows.Count, "A").End(xlUp).row
            On Error Resume Next
            ActiveSheet.ShowAllData
            On Error GoTo 0
            Range("A18:A" & lastRow).Select
            Selection.Copy
            
    End If
    
    'pega los datos
    session.findById("wnd[1]/tbar[0]/btn[24]").press
    session.findById("wnd[1]/tbar[0]/btn[8]").press
    
    'Ejecuta
    session.findById("wnd[0]/tbar[1]/btn[8]").press
    
    'Descarga el archivo
    session.findById("wnd[0]/usr/cntlCONT9000/shellcont/shell").pressToolbarContextButton "&MB_EXPORT"
    session.findById("wnd[0]/usr/cntlCONT9000/shellcont/shell").selectContextMenuItem "&PC"
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").Select
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").SetFocus
    session.findById("wnd[1]/tbar[0]/btn[0]").press
    session.findById("wnd[1]/usr/ctxtDY_PATH").Text = Ruta_Trimestre
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = "Cargos_" & fechaFormato & ".XLS"
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 9
    session.findById("wnd[1]/tbar[0]/btn[11]").press
    
    'Guarda y cierra el archivo de conflictos
    MAESTRO.Save
    MAESTRO.Close
    
    'Organizar documento
    Application.CutCopyMode = False
    Workbooks.Open Ruta_Trimestre & "\" & "Cargos_" & fechaFormato & ".XLS"
    ActiveSheet.Cells.Select
    Selection.Copy
    Workbooks.Add
    ActiveSheet.Paste
    ActiveWorkbook.SaveAs Ruta_Trimestre & "\" & "Cargos_" & fechaFormato & ".XLSX"
    ActiveWorkbook.Close SaveChanges:=True
    Workbooks("Cargos_" & fechaFormato & ".XLS").Close
    Kill Ruta_Trimestre & "\" & "Cargos_" & fechaFormato & ".XLS"
    
End Sub

Sub ZHR929()
InicializarVariables

'-------------------- DESCARGA LOS DATOS DE LA ZHR929 CON EL LEYOUT /MCS25.06 --------------------
    
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
    session.findById("wnd[0]/tbar[0]/okcd").Text = "/nZHR929"
    session.findById("wnd[0]").sendVKey 0
    
    'Cambia fechas
    session.findById("wnd[0]/usr/ctxtPNPBEGDA").Text = Fecha1
    session.findById("wnd[0]/usr/ctxtPNPENDDA").Text = Fecha2
    
    'Coloca que sean solo empleados de abs y que esten activos
    session.findById("wnd[0]/usr/ctxtPNPSTAT2-LOW").Text = "3"
    session.findById("wnd[0]/usr/ctxtPNPBUKRS-LOW").Text = "tc*"
    
    'Cambia layout
    session.findById("wnd[0]/usr/ctxtP_VAR").Text = "/MCS26.05"
    session.findById("wnd[0]/usr/ctxtP_VAR").SetFocus
    session.findById("wnd[0]/usr/ctxtP_VAR").caretPosition = 9
    
    'Ejecuta
    session.findById("wnd[0]/tbar[1]/btn[8]").press
    
    'Descarga el archivo
    session.findById("wnd[0]/mbar/menu[0]/menu[3]/menu[2]").Select
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").Select
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").SetFocus
    session.findById("wnd[1]/tbar[0]/btn[0]").press
    session.findById("wnd[1]/usr/ctxtDY_PATH").Text = Ruta_Trimestre
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = "ZHR929_" & fechaFormato & ".XLS"
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 16
    session.findById("wnd[1]/tbar[0]/btn[11]").press
    
    'Organizar documento
    Application.CutCopyMode = False
    Workbooks.Open Ruta_Trimestre & "\" & "ZHR929_" & fechaFormato & ".XLS"
    ActiveSheet.Cells.Select
    Selection.Copy
    Workbooks.Add
    ActiveSheet.Paste
    ActiveWorkbook.SaveAs Ruta_Trimestre & "\" & "ZHR929_" & fechaFormato & ".XLSX"
    ActiveWorkbook.Close SaveChanges:=True
    Workbooks("ZHR929_" & fechaFormato & ".XLS").Close
    Kill Ruta_Trimestre & "\" & "ZHR929_" & fechaFormato & ".XLS"
    

End Sub

Sub ZHR1252()
InicializarVariables

'-------------------- DESCARGA LOS DATOS DE LA ZHR1252 --------------------
    
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
    session.findById("wnd[0]/tbar[0]/okcd").Text = "/nZHR1252"
    session.findById("wnd[0]").sendVKey 0
    
    'Pone a fecha de hoy
    session.findById("wnd[0]/usr/txtS_USRID-LOW").SetFocus
    session.findById("wnd[0]/usr/txtS_USRID-LOW").caretPosition = 0
    session.findById("wnd[0]/usr/btn%_S_USRID_%_APP_%-VALU_PUSH").press
    
    'Abre para pegar los datos de los empleados
    session.findById("wnd[0]/usr/radPNPTIMR1").Select
    session.findById("wnd[0]/usr/btn%_S_USRID_%_APP_%-VALU_PUSH").press
    
    Workbooks("MCS 25.06.01 " & Trimestre_letra & " " & "H2R" & " " & año & ".xlsx").Activate
    Sheets("Revisión de Cargos H2R-HR-IT").Activate
    On Error Resume Next
    ActiveSheet.ShowAllData
    On Error GoTo 0
    lastRow = ActiveSheet.Cells(ActiveSheet.Rows.Count, "A").End(xlUp).row
    Range("A2:A" & lastRow).Select
    Selection.Copy
    
    'Pega la info
    session.findById("wnd[1]/tbar[0]/btn[24]").press
    session.findById("wnd[1]/tbar[0]/btn[8]").press
    
    'Ejecuta
    session.findById("wnd[0]/tbar[1]/btn[8]").press

    'Descarga el archivo
    session.findById("wnd[0]/mbar/menu[0]/menu[3]/menu[2]").Select
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").Select
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").SetFocus
    session.findById("wnd[1]/tbar[0]/btn[0]").press
    session.findById("wnd[1]/usr/ctxtDY_PATH").Text = Ruta_Trimestre
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = "CARGO SAP GLOBAL_" & fechaFormato & ".XLS"
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 16
    session.findById("wnd[1]/tbar[0]/btn[11]").press
    
    'Organizar documento
    Application.CutCopyMode = False
    Workbooks.Open Ruta_Trimestre & "\" & "CARGO SAP GLOBAL_" & fechaFormato & ".XLS"
    ActiveSheet.Cells.Select
    Selection.Copy
    Workbooks.Add
    ActiveSheet.Paste
    ActiveWorkbook.SaveAs Ruta_Trimestre & "\" & "CARGO SAP GLOBAL_" & fechaFormato & ".XLSX"
    ActiveWorkbook.Close SaveChanges:=True
    Workbooks("CARGO SAP GLOBAL_" & fechaFormato & ".XLS").Close
    Kill Ruta_Trimestre & "\" & "CARGO SAP GLOBAL_" & fechaFormato & ".XLS"
    

End Sub

Sub ZBC033_Criticas()
InicializarVariables

'-------------------- DESCARGA LOS DATOS DE LA ZBC033 DESCARGAR LAS TRX CRITICAS --------------------
    
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
    session.findById("wnd[0]/tbar[0]/okcd").Text = "/nZBC033"
    session.findById("wnd[0]").sendVKey 0
    
    'Entra a transaccion por usuario
    session.findById("wnd[0]/usr/radRB_TCODE").Select
    session.findById("wnd[0]/usr/radRB_TCODE").SetFocus
    
    'Ingresa las transcciones criticas
    session.findById("wnd[0]/tbar[1]/btn[8]").press
    session.findById("wnd[0]/usr/btn%_S_LOW_%_APP_%-VALU_PUSH").press
    
    ThisWorkbook.Sheets("TRX").Activate
    lastRow = ActiveSheet.Cells(ActiveSheet.Rows.Count, "A").End(xlUp).row
    Range("A2:A" & lastRow).Select
    Selection.Copy
    
    'Pega
    session.findById("wnd[1]/tbar[0]/btn[24]").press
    session.findById("wnd[1]/tbar[0]/btn[8]").press
    
    'Ejecuta
    session.findById("wnd[0]/tbar[1]/btn[8]").press
    
    'Descarga el archivo
    session.findById("wnd[0]/usr/cntlCONT9000/shellcont/shell").pressToolbarContextButton "&MB_EXPORT"
    session.findById("wnd[0]/usr/cntlCONT9000/shellcont/shell").selectContextMenuItem "&PC"
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").Select
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").SetFocus
    session.findById("wnd[1]/tbar[0]/btn[0]").press
    session.findById("wnd[1]/usr/ctxtDY_PATH").Text = Ruta_Trimestre
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = "TRANSACCIONES CRÍTICAS_" & fechaFormato & ".XLS"
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 9
    session.findById("wnd[1]/tbar[0]/btn[11]").press
        
    'Organizar documento
    Application.CutCopyMode = False
    Workbooks.Open Ruta_Trimestre & "\" & "TRANSACCIONES CRÍTICAS_" & fechaFormato & ".XLS"
    ActiveSheet.Cells.Select
    Selection.Copy
    Workbooks.Add
    ActiveSheet.Paste
    ActiveWorkbook.SaveAs Ruta_Trimestre & "\" & "TRANSACCIONES CRÍTICAS_" & fechaFormato & ".XLSX"
    ActiveWorkbook.Close SaveChanges:=True
    Workbooks("TRANSACCIONES CRÍTICAS_" & fechaFormato & ".XLS").Close
    Kill Ruta_Trimestre & "\" & "TRANSACCIONES CRÍTICAS_" & fechaFormato & ".XLS"

End Sub
Sub Plantilla()

    InicializarVariables
    
    'Elimina el archivo si ya existe
    On Error Resume Next
    Kill Ruta_Trimestre & "\" & "MCS 25.06.01 " & Trimestre_letra & " " & año & ".xlsx"
    On Error GoTo 0
    
    'Hace copia de la plantilla del ICS
    Set wbPlantilla = Workbooks.Open(ruta & "\" & "MCS 25.06.01 PLANTILLA.xlsx")
    wbPlantilla.Activate
    ActiveWorkbook.SaveCopyAs Filename:=Ruta_Trimestre & "\" & "MCS 25.06.01 " & Trimestre_letra & " " & "H2R" & " " & año & ".xlsx"
    wbPlantilla.Close False
    
End Sub


Sub Desarrollo_organizar()

    InicializarVariables
    
    'Abre el achivo del control
    Set wbAuditoria = Workbooks.Open(Ruta_Trimestre & "\" & "MCS 25.06.01 " & Trimestre_letra & " " & "H2R" & " " & año & ".xlsx")
    
    '---------------------------- PARTE DE ORGANIZARLA HOJA DE CARGOS TRANSACCIONES ----------------------------
    
    'Abre el archivo del los cargos
    Workbooks.Open Ruta_Trimestre & "\" & "Cargos_" & fechaFormato & ".XLSX"
    Workbooks("Cargos_" & fechaFormato & ".XLSX").Activate
    Sheets(1).Activate
    Columns("A").Delete
    Rows("1:4").Delete
    Rows("2").Delete
    ActiveWorkbook.Save
    lastRow = ActiveSheet.Cells(ActiveSheet.Rows.Count, "B").End(xlUp).row
    On Error Resume Next
    ActiveSheet.ShowAllData
    On Error GoTo 0
    
    'Pega funcion
    Range("A2:A" & lastRow).Select
    Selection.Copy
    wbAuditoria.Activate
    Sheets("CARGOS TRANSACCIONES").Activate
    Range("B2").PasteSpecial Paste:=xlPasteAll
    
    'Pega Usuario
    Workbooks("Cargos_" & fechaFormato & ".XLSX").Activate
    Sheets(1).Activate
    Range("B2:B" & lastRow).Select
    Selection.Copy
    wbAuditoria.Activate
    Sheets("CARGOS TRANSACCIONES").Activate
    Range("A2").PasteSpecial Paste:=xlPasteAll
    
    'Pega nombre de usuario
    Workbooks("Cargos_" & fechaFormato & ".XLSX").Activate
    Sheets(1).Activate
    Range("C2:C" & lastRow).Select
    Selection.Copy
    wbAuditoria.Activate
    Sheets("CARGOS TRANSACCIONES").Activate
    Range("C2").PasteSpecial Paste:=xlPasteAll
    Workbooks("Cargos_" & fechaFormato & ".XLSX").Save
    Workbooks("Cargos_" & fechaFormato & ".XLSX").Close
    
    'Hace la formula de extraer y la copia y la pega como valores
    wbAuditoria.Activate
    Sheets("CARGOS TRANSACCIONES").Activate
    lastRow = ActiveSheet.Cells(Rows.Count, 1).End(xlUp).row
    Range("D2:D" & lastRow) = "=MID(RC[-2],5,3)"
    Columns("D").Select
    Selection.Copy
    Columns("D").Select
    Selection.PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    
    'Cambiar nombre de las fabrics
    Sheets("CARGOS TRANSACCIONES").Activate
    lastRow = ActiveSheet.Cells(Rows.Count, 1).End(xlUp).row
    For i = 2 To lastRow
        cellValue = Sheets("CARGOS TRANSACCIONES").Cells(i, "D").Value
        Select Case cellValue
            Case "B2R"
                Sheets("CARGOS TRANSACCIONES").Cells(i, "D").Value = "R2R"
            Case "HRC", "SHR"
                Sheets("CARGOS TRANSACCIONES").Cells(i, "D").Value = "HR"
            Case "ITC"
                Sheets("CARGOS TRANSACCIONES").Cells(i, "D").Value = "IT"
        End Select
    Next i
    
    
    'Elimina las filas que tengan el evaluador de dialogo: Este rol lo puede tener cualquiera dado que se relaciona con el funcionamiento de mi llave
    Sheets("CARGOS TRANSACCIONES").Activate
    lastRow = ActiveSheet.Cells(Rows.Count, 1).End(xlUp).row
    For i = lastRow To 2 Step -1
        If Sheets("CARGOS TRANSACCIONES").Cells(i, "B").Value = "HLTCHRCF_EVALUADOR_DIALOGO" Then
            Sheets("CARGOS TRANSACCIONES").Rows(i).Delete
        End If
    Next i
    
    'Filtrar los de las fabricas que nos interesan
    Sheets("CARGOS TRANSACCIONES").Activate
    If ActiveSheet.AutoFilterMode Then
        ActiveSheet.AutoFilterMode = False
    End If
    Rows("1:1").AutoFilter Field:=4, Criteria1:=Array("H2R", "HR", "IT"), Operator:=xlFilterValues
    
    'pega los datos de las fabricas a la hoja "Revisión de Cargos H2R-HR-IT"
    lastRow = ActiveSheet.Cells(ActiveSheet.Rows.Count, "A").End(xlUp).row
    Range("A2:D" & lastRow).SpecialCells(xlCellTypeVisible).Copy
    Sheets("Revisión de Cargos H2R-HR-IT").Activate
    Range("A2").PasteSpecial Paste:=xlPasteAll
    Sheets("CARGOS TRANSACCIONES").Activate
    Application.CutCopyMode = False
    If ActiveSheet.AutoFilterMode Then
        ActiveSheet.AutoFilterMode = False
    End If
    
    
    '---------------------------- PARTE DE ORGANIZARLA HOJA DE LA ZHR929 ----------------------------
    
    'Abre el archivo del los cargos
    Workbooks.Open Ruta_Trimestre & "\" & "ZHR929_" & fechaFormato & ".XLSX"
    Workbooks("ZHR929_" & fechaFormato & ".XLSX").Activate
    Sheets(1).Activate
    Columns("A").Delete
    Rows("1:4").Delete
    Rows("2").Delete
    ActiveWorkbook.Save
    lastRow = ActiveSheet.Cells(ActiveSheet.Rows.Count, "A").End(xlUp).row
    On Error Resume Next
    ActiveSheet.ShowAllData
    On Error GoTo 0
    
    
    'Pega las primeras 4 columnas
    Range("A2:D" & lastRow).Select
    Selection.Copy
    wbAuditoria.Activate
    Sheets("ZHR929").Activate
    Range("A2").PasteSpecial Paste:=xlPasteAll
    Application.CutCopyMode = False
    
    'Pega Usuario
    Workbooks("ZHR929_" & fechaFormato & ".XLSX").Activate
    Sheets(1).Activate
    lastRow = ActiveSheet.Cells(ActiveSheet.Rows.Count, "A").End(xlUp).row
    Range("E2:N" & lastRow).Select
    Selection.Copy
    wbAuditoria.Activate
    Sheets("ZHR929").Activate
    Range("G2").PasteSpecial Paste:=xlPasteAll
    Application.CutCopyMode = False
    Workbooks("ZHR929_" & fechaFormato & ".XLSX").Save
    Workbooks("ZHR929_" & fechaFormato & ".XLSX").Close
    
    'Buscarv para la columna de cargo y fabrica
    wbAuditoria.Activate
    Sheets("ZHR929").Activate
    lastRow = ActiveSheet.Cells(ActiveSheet.Rows.Count, "A").End(xlUp).row
    Range("E2:E" & lastRow) = "=VLOOKUP(RC[-1],'CARGOS TRANSACCIONES'!C[-4]:C[-1],4,0)"
    Range("F2:F" & lastRow) = "=VLOOKUP(RC[-2],'CARGOS TRANSACCIONES'!C[-5]:C[-4],2,0)"
    Columns("E:F").Select
    Selection.Copy
    Columns("E").Select
    Selection.PasteSpecial Paste:=xlPasteValues
    wbAuditoria.Save
    
    '---------------------------- HACE LOS BUSCAR V CON LA HOJA DE GLOBAL ----------------------------

    ZHR1252 'Va al sub a descargar los datos globales
    
    'Abre el archivo de global
    Workbooks.Open Ruta_Trimestre & "\" & "CARGO SAP GLOBAL_" & fechaFormato & ".XLSX"
    Workbooks("CARGO SAP GLOBAL_" & fechaFormato & ".XLSX").Activate
    Sheets(1).Activate
    Columns("A").Delete
    Rows("1:4").Delete
    Rows("2").Delete
    ActiveWorkbook.Save
    lastRow = ActiveSheet.Cells(ActiveSheet.Rows.Count, "A").End(xlUp).row
    On Error Resume Next
    ActiveSheet.ShowAllData
    On Error GoTo 0
    
    'Pega
    Range("A2:K" & lastRow).Select
    Selection.Copy
    wbAuditoria.Activate
    Sheets("CARGOS SAP GLOBAL").Activate
    Range("A2").PasteSpecial Paste:=xlPasteAll
    Workbooks("CARGO SAP GLOBAL_" & fechaFormato & ".XLSX").Save
    Workbooks("CARGO SAP GLOBAL_" & fechaFormato & ".XLSX").Close
    
    'Hace los buscarv
    wbAuditoria.Activate
    Sheets("Revisión de Cargos H2R-HR-IT").Activate
    lastRow = ActiveSheet.Cells(ActiveSheet.Rows.Count, "A").End(xlUp).row
    Range("E2:E" & lastRow) = "=VLOOKUP(RC[-4],'CARGOS SAP GLOBAL'!C[-2]:C,3,0)"
    Range("F2:F" & lastRow) = "=VLOOKUP(RC[-5],ZHR929!C[-2]:C[1],4,0)"
    Columns("E:F").Select
    Selection.Copy
    Columns("E").Select
    Selection.PasteSpecial Paste:=xlPasteValues
    
    
    '---------------------------- TRANSACCIONES CRITICAS ----------------------------
    
    'Abre el archivo de global
    Workbooks.Open Ruta_Trimestre & "\" & "TRANSACCIONES CRÍTICAS_" & fechaFormato & ".XLSX"
    Workbooks("TRANSACCIONES CRÍTICAS_" & fechaFormato & ".XLSX").Activate
    Sheets(1).Activate
    Columns("A").Delete
    Rows("1:4").Delete
    Range("A1").Value = "Transaccion"
    Range("B1").Value = "Rol asociado"
    Range("C1").Value = "User"
    Range("D1").Value = "Nombre User"
    Range("E1").Value = "Sociedad"
    Range("F1").Value = "Coordinacion"
    Range("G1").Value = "Cargo"
    Range("H1").Value = "Email"
    Range("I1").Value = "Validez"
    ActiveWorkbook.Save
    lastRow = ActiveSheet.Cells(ActiveSheet.Rows.Count, "A").End(xlUp).row
    On Error Resume Next
    ActiveSheet.ShowAllData
    On Error GoTo 0
    
    'Pega los datos al reporte
    Range("A2:G" & lastRow).Select
    Selection.Copy
    wbAuditoria.Activate
    Sheets("REVISIÓN H2R").Activate
    Range("A2").PasteSpecial Paste:=xlPasteAll
    Workbooks("TRANSACCIONES CRÍTICAS_" & fechaFormato & ".XLSX").Activate
    Sheets(1).Activate
    Range("H2:I" & lastRow).Select
    Selection.Copy
    wbAuditoria.Activate
    Sheets("REVISIÓN H2R").Activate
    Range("J2").PasteSpecial Paste:=xlPasteAll
    Workbooks("TRANSACCIONES CRÍTICAS_" & fechaFormato & ".XLSX").Save
    Workbooks("TRANSACCIONES CRÍTICAS_" & fechaFormato & ".XLSX").Close
    wbAuditoria.Save
    
    'Cambiar formato de fecha
    Sheets("REVISIÓN H2R").Activate
    Dim CelsD As Range
    Dim UltimaFilaD As Long
    UltimaFilaD = ActiveSheet.Cells(Rows.Count, 1).End(xlUp).row
    
    If UltimaFilaD >= 2 Then
        For Each CelsD In ActiveSheet.Range("K2:K" & UltimaFilaD)
            If Len(CelsD.Value) = 10 And Mid(CelsD.Value, 3, 1) = "." And Mid(CelsD.Value, 6, 1) = "." Then
                CelsD.Value = DateSerial(Right(CelsD.Value, 4), Mid(CelsD.Value, 4, 2), Left(CelsD.Value, 2))
                CelsD.NumberFormat = "dd/mm/yyyy"
            End If
        Next CelsD
    End If
    
    'eliminamos todo lo que esté anterior a la fecha en la que realizamos el control,
    'ya que es información que no necesitamos. Se deja solo información con fechas futuras o al infinito.
    
    'Define el primer dia del semestre
    Select Case Trimestre
        Case "1"
            incio_trimestre = Format(DateSerial(año, 1, 1), "dd/mm/yyyy")
        Case "2"
            incio_trimestre = Format(DateSerial(año, 4, 1), "dd/mm/yyyy")
        Case "3"
            incio_trimestre = Format(DateSerial(año, 7, 1), "dd/mm/yyyy")
        Case "4"
            incio_trimestre = Format(DateSerial(año, 10, 1), "dd/mm/yyyy")
    End Select
    
    infinito = "12/31/9999"
    
    Set rng = Sheets("REVISIÓN H2R").Range("K1:K" & Sheets("REVISIÓN H2R").Cells(Sheets("REVISIÓN H2R").Rows.Count, "K").End(xlUp).row)
    For Each cell In rng
        If IsDate(cell.Value) Then
            If CDate(cell.Value) < incio_trimestre Or CDate(cell.Value) <> infinito Then
                Set rowToColor = Sheets("REVISIÓN H2R").Range("A" & cell.row & ":N" & cell.row)
                rowToColor.Interior.Color = RGB(255, 255, 0) ' Color amarillo
            End If
        End If
    Next cell
    
    
    'Completa informacion
    wbAuditoria.Activate
    Sheets("REVISIÓN H2R").Activate
    lastRow = ActiveSheet.Cells(ActiveSheet.Rows.Count, "A").End(xlUp).row
    Range("H2:H" & lastRow) = "=VLOOKUP(RC[-5],'CARGOS TRANSACCIONES'!C1:C4,4,0)"
    Range("I2:I" & lastRow) = "=VLOOKUP(RC[-6],ZHR929!C[-5]:C[-2],4,0)"
    Columns("H:I").Select
    Selection.Copy
    Columns("H").Select
    Selection.PasteSpecial Paste:=xlPasteValues
    wbAuditoria.Save
    
    'Organizar en orden alfabetico
    lastRow = ActiveSheet.Cells(ActiveSheet.Rows.Count, "A").End(xlUp).row
    wbAuditoria.Worksheets("REVISIÓN H2R").AutoFilter.Sort.SortFields.Clear
    wbAuditoria.Worksheets("REVISIÓN H2R").AutoFilter.Sort.SortFields.Add Key _
        :=Range("A1:A" & lastRow), SortOn:=xlSortOnValues, Order:=xlAscending, _
        DataOption:=xlSortNormal
    With wbAuditoria.Worksheets("REVISIÓN H2R").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    wbAuditoria.Save
    
End Sub

