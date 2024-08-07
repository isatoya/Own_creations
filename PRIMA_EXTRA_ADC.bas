Attribute VB_Name = "PRIMA_EXTRA_ADC"
Option Explicit
'variables para todo el proyecto

    Dim MAESTRO, wb_auditoria As Workbook
    Dim Fecha1, Fecha2, inicio_sem, fin_sem, mes, Mes_texto, Año, variante, ruta, Ruta_Año, Ruta_Mes, Ruta_Audi, ruta_maestroA, promedioFormula As String
    Dim LastRow, i As Long
    Dim CelsAG As Object
    Dim rng, row, visibleRange As Range
    Dim cell As Range
    
Sub InicializarVariables()
'Definicion de las variables
    
    '------------ Variables originales del archivo -----------
    mes = ThisWorkbook.Sheets("Reportes").Range("N8").Value
    Mes_texto = ThisWorkbook.Sheets("Reportes").Range("I12").Value
    Año = ThisWorkbook.Sheets("Reportes").Range("I10").Value
    Fecha1 = ThisWorkbook.Sheets("Reportes").Range("I8").Value
    Fecha2 = ThisWorkbook.Sheets("Reportes").Range("M8").Value
    
    '------------ Rutas -----------
    ruta = ThisWorkbook.Path & "\"
    Ruta_Año = ruta & Año
    Ruta_Mes = Ruta_Año & "\" & mes & ". " & Mes_texto
    Ruta_Audi = Ruta_Mes & "\" & "AUDITORIAS DE NOMINA"
    
    
    '------------ Calcular fechas de los semestres-----------
    'Define el primer dia del semestre
    Select Case Mes_texto
        Case "Junio"
            inicio_sem = Format(DateSerial(Año, 1, 1), "dd.mm.yyyy")
        Case "Diciembre"
            inicio_sem = Format(DateSerial(Año, 7, 1), "dd.mm.yyyy")
    End Select
    
    'Define el ultimo dia del semestre
    Select Case Mes_texto
        Case "Junio"
            fin_sem = Format(DateSerial(Año, 6, 30), "dd.mm.yyyy")
        Case "Diciembre"
            fin_sem = Format(DateSerial(Año, 12, 31), "dd.mm.yyyy")
    End Select

End Sub

Sub Ejecutar_AUDI_EXTRA_LEGAL_ADC()

    ' Verificar si hay datos en las celdas I8 y M8
    If ThisWorkbook.Sheets("Reportes").Range("I8").Value = "" Or ThisWorkbook.Sheets("Reportes").Range("M8").Value = "" Then
        MsgBox "Datos incompletos, por favor ingrese los datos antes de ejecutar.", vbExclamation
        Exit Sub
    End If

    ' Llama a cada una de las funciones
    DeactivateStuff
    CrearCarpetas
    Copia_maestro
    SAP_1028
    SAP_BASES
    SAP_LNR
    Auditoria
    ReactivateStuff
    
    MsgBox "Reporte finalizado. Lo podra encontrar en carpeta de auditorias", vbInformation

End Sub

Sub CrearCarpetas()
    
    'Creación y validacion de las carpetas
    InicializarVariables
    '''''''
    ''AÑO''
    '''''''
    Ruta_Año = ruta & Año
    If Dir(Ruta_Año, vbDirectory + vbHidden) = "" Then
        'Comprueba que la carpeta no exista para crear el directorio.
        If Dir(Ruta_Año & vbDirectory + vbHidden) = "" Then MkDir Ruta_Año
    End If
    '''''''
    ''MES''
    '''''''
    Ruta_Mes = Ruta_Año & "\" & mes & ". " & Mes_texto
    If Dir(Ruta_Mes, vbDirectory + vbHidden) = "" Then
        'Comprueba que la carpeta no exista para crear el directorio.
        If Dir(Ruta_Mes & vbDirectory + vbHidden) = "" Then MkDir Ruta_Mes
    End If
    ''''''''''''''''''''
    ''AUDITORIA NÓMINA''
    ''''''''''''''''''''
    Ruta_Audi = Ruta_Mes & "\" & "AUDITORIAS DE NOMINA"
    If Dir(Ruta_Audi, vbDirectory + vbHidden) = "" Then
        'Comprueba que la carpeta no exista para crear el directorio.
        If Dir(Ruta_Audi & vbDirectory + vbHidden) = "" Then MkDir Ruta_Audi
    End If
        
End Sub

Sub Copia_maestro()

    InicializarVariables
    
    ' Solicitar al usuario que abra un archivo
    MsgBox "Por favor selecciene el archivo del Maestro de Activos correspondiente.", vbInformation
    ruta_maestroA = Application.GetOpenFilename("Archivos Excel (*.xls; *.xlsx), *.xls; *.xlsx")
    Application.AskToUpdateLinks = False
    
    ' Abre el reporte seleccionado por el usuario y copiia la primera hoja
    If ruta_maestroA <> "Falso" Then
        
        Set MAESTRO = Workbooks.Open(Filename:=ruta_maestroA, UpdateLinks:=0)
        MAESTRO.Activate
        Worksheets("SALARIAL").Activate
        
        If ActiveSheet.AutoFilterMode Then
            ActiveSheet.AutoFilterMode = False
        End If
        Columns("A:X").Select
        Selection.Copy
        
        'Crea el archivo nuevo
        Workbooks.Add
        ActiveSheet.Range("A1").PasteSpecial Paste:=xlPasteAll
        Application.CutCopyMode = False

        ' Verifica si el archivo ya existe en la ruta y lo elimina
        If Dir(Ruta_Audi & "\" & Año & "." & mes & "." & "AUDITORIA PRIMA EXTRALEGAL ADC" & ".XLSX") <> "" Then
            Kill Ruta_Audi & "\" & Año & "." & mes & "." & "AUDITORIA PRIMA EXTRALEGAL ADC" & ".XLSX"
        End If
        
        'Crea el archivo nuevo
        ActiveSheet.Name = "MAESTRO"
        ActiveWorkbook.SaveAs Ruta_Audi & "\" & Año & "." & mes & "." & "AUDITORIA PRIMA EXTRALEGAL ADC" & ".XLSX"
        ActiveWorkbook.Close SaveChanges:=True
        MAESTRO.Close
        Application.AskToUpdateLinks = True
        
    Else
        MsgBox "Operación cancelada por el usuario.", vbInformation
    End If
End Sub

Sub Auditoria()

    InicializarVariables
 
    'Abre el archivo de la auditoria
    Set wb_auditoria = Workbooks.Open(Ruta_Audi & "\" & Año & "." & mes & "." & "AUDITORIA PRIMA EXTRALEGAL ADC" & ".XLSX")
    wb_auditoria.Activate
    Sheets.Add(After:=Sheets(Sheets.Count)).Name = "AUDITORIA"
    Sheets.Add(After:=Sheets(Sheets.Count)).Name = "LNR"
    Sheets.Add(After:=Sheets(Sheets.Count)).Name = "BASES"
    
    'Pega los datos del cc nom 1028 en la hoja de auditoria
    Workbooks.Open Ruta_Audi & "\" & "1028-" & Mes_texto & ".XLSX"
    Workbooks("1028-" & Mes_texto & ".XLSX").Activate
    Sheets(1).Activate
    Range("A:H").Copy
    wb_auditoria.Activate
    Sheets("AUDITORIA").Activate
    Range("A1").PasteSpecial Paste:=xlPasteAll
    Columns("A:H").AutoFit
    Workbooks("1028-" & Mes_texto & ".XLSX").Save
    Workbooks("1028-" & Mes_texto & ".XLSX").Close
    
    'Organiza primera fila con los titulos y pone los formatos de las celdas
    wb_auditoria.Activate
    Sheets("AUDITORIA").Activate
    Range("I1").Value = "AREA NOMINA"
    Range("J1").Value = "AREA PERSONAL"
    Range("K1").Value = "F. ALTA"
    Range("L1").Value = "F. INICIO"
    Range("M1").Value = "F. FIN"
    Range("N1").Value = "DIAS"
    Range("O1").Value = "LNR"
    Range("P1").Value = "TOTAL DIAS"
    Range("Q1").Value = "CAL. DIAS"
    Range("R1").Value = "DIFERENCIA"
    Range("S1").Value = "BASE"
    Range("T1").Value = "CALCULO MANUAL"
    Range("U1").Value = "DIFERENCIA"
    Range("V1").Value = "POSICION"
    Range("W1").Value = "RELACION LABORAL"
    
    With Range("A1:H1")
        .Interior.Color = RGB(178, 186, 187)
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
    End With
    
    With Range("I1:W1")
        .Interior.Color = RGB(214, 234, 248)
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
    End With
    Columns("K:M").NumberFormat = "dd/mm/yyyy"
    Columns("Q:R").NumberFormat = "0.00"
    Columns("A:W").AutoFit
    
    'Emepieza a realizar las formulas
    LastRow = ActiveSheet.Cells(Rows.Count, 1).End(xlUp).row
    
    Sheets("AUDITORIA").Range("I2:I" & LastRow) = "=VLOOKUP(RC[-8],MAESTRO!R2C1:R1758C6,6,0)"
    
    Sheets("AUDITORIA").Range("J2:J" & LastRow) = "=VLOOKUP(RC[-9],MAESTRO!R1C1:R1758C22,22,0)"
    
    Sheets("AUDITORIA").Range("K2:K" & LastRow) = "=VLOOKUP(RC[-10],MAESTRO!R2C1:R1758C16,16,0)"
    
    Sheets("AUDITORIA").Range("L2:L" & LastRow) = "=IF(RC[-1]<(DATE(" & Año & "," & 1 & ",1)),(DATE(" & Año & "," & 1 & ",1)),RC[-1])"

    Sheets("AUDITORIA").Range("M2:M" & LastRow) = "=DATE(" & Año & "," & 6 & ",30)"
    
    Sheets("AUDITORIA").Range("N2:N" & LastRow) = "=DAYS360(RC[-2],RC[-1])+1"
    
    
    Sheets("AUDITORIA").Range("P2:P" & LastRow) = "=RC[-2]-RC[-1]"
    
    Sheets("AUDITORIA").Range("Q2:Q" & LastRow) = "=RC[-1]*30/180"

    Sheets("AUDITORIA").Range("R2:R" & LastRow) = "=RC[-11]-RC[-1]"

    Sheets("AUDITORIA").Range("T2:T" & LastRow) = "=RC[-4]*RC[-1]/180"

    Sheets("AUDITORIA").Range("U2:U" & LastRow) = "=RC[-13]-RC[-1]"

    Sheets("AUDITORIA").Range("V2:V" & LastRow) = "=VLOOKUP(RC[-21],MAESTRO!R2C1:R1758C24,24,0)"

    Sheets("AUDITORIA").Range("W2:W" & LastRow) = "=VLOOKUP(RC[-22],MAESTRO!R2C1:R1758C5,5,0)"
    Columns("A:W").AutoFit
    
    'Crea la hoja de las bases
    Workbooks.Open Ruta_Audi & "\" & "BASES PRIMA-" & Mes_texto & ".XLSX"
    Workbooks("BASES PRIMA-" & Mes_texto & ".XLSX").Activate
    Sheets(1).Activate
    Range("A:H").Copy
    wb_auditoria.Activate
    Sheets("BASES").Activate
    Range("A1").PasteSpecial Paste:=xlPasteAll
    Columns("A:H").AutoFit
    Workbooks("BASES PRIMA-" & Mes_texto & ".XLSX").Save
    Workbooks("BASES PRIMA-" & Mes_texto & ".XLSX").Close
    wb_auditoria.Activate
    Sheets("BASES").Activate
    
    'Elimina lo que tenga en periodo para 0
    Sheets("BASES").Activate
    LastRow = Cells(Rows.Count, "A").End(xlUp).row
    For i = LastRow To 1 Step -1
        If ActiveSheet.Cells(i, "C").Value = "0" Then
            ActiveSheet.Rows(i).Delete
        End If
    Next i
    
    'Elimina lo que no sea del semestre

    Columns("D:D").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    LastRow = ActiveSheet.Cells(ActiveSheet.Rows.Count, "A").End(xlUp).row
    
    If Mes_texto = "Junio" Then
                Range("D2:D" & LastRow) = "=YEAR(RC[1])"
                Columns("D:D").Select
                Selection.Copy
                Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                    :=False, Transpose:=False
                For i = LastRow To 2 Step -1
                    If Cells(i, "D").Value = Año - 1 Then
                        Rows(i).Delete
                    End If
                Next i
                Columns("D:D").Select
                Selection.Delete
    
    Else
                Range("D2:D" & LastRow) = "=MID(RC[-1],1,4)"
                Columns("D:D").Select
                Selection.Copy
                Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                    :=False, Transpose:=False
                For i = LastRow To 1 Step -1
                    If InStr(1, Año - 1, ActiveSheet.Cells(i, "D").Value) > 0 Then
                        ActiveSheet.Rows(i).Delete
                    End If
                Next i
                Columns("D:D").Select
                Selection.Delete
                
                Range("D2:D" & LastRow) = "=MID(RC[-1],5,2)"
                Columns("D:D").Select
                Selection.Copy
                Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                    :=False, Transpose:=False
                For i = LastRow To 1 Step -1
                    If Val(ActiveSheet.Cells(i, "D").Value) <= 6 Then
                        ActiveSheet.Rows(i).Delete
                    End If
                Next i
                Columns("D:D").Select
                Selection.Delete
    End If
    

    'Tabla dinamica
    'Determina el rango
        Dim ult_Tabla As Long
        ult_Tabla = Cells(Rows.Count, "A").End(xlUp).row
        
        Dim rangoTabla1 As Range
        Set rangoTabla1 = Sheets("BASES").Range("A1:H" & ult_Tabla)
        ActiveSheet.ListObjects.Add(xlSrcRange, rangoTabla1, , xlYes).Name = "Tabla1"
        
        'Crear tabla dinamica
    
        Dim celdaTablaDinamica1 As Range
        Set celdaTablaDinamica1 = Sheets("BASES").Range("L1")
        Dim tablaDinamica1 As PivotTable
        
        'Activa campos y le pone formato tabular
        ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
            rangoTabla1, Version:=6).CreatePivotTable TableDestination:= _
            celdaTablaDinamica1, TableName:="tablaDinamica1", DefaultVersion:=6
            
        Sheets("BASES").Select
        With ActiveSheet.PivotTables("tablaDinamica1").PivotFields("Nº pers.")
        .Orientation = xlRowField
        .Position = 1
        End With
        With ActiveSheet.PivotTables("tablaDinamica1").PivotFields("Apellido Nombre")
            .Orientation = xlRowField
            .Position = 2
        End With
        With ActiveSheet.PivotTables("tablaDinamica1").PivotFields("Fecha pago")
            .Orientation = xlColumnField
            .Position = 1
        End With
        ActiveSheet.PivotTables("tablaDinamica1").PivotFields("Fecha pago").AutoGroup
        ActiveSheet.PivotTables("tablaDinamica1").AddDataField ActiveSheet.PivotTables( _
            "tablaDinamica1").PivotFields("Importe"), "Suma de Importe", xlSum
        ActiveSheet.PivotTables("tablaDinamica1").PivotFields("Nº pers.").Subtotals = _
            Array(False, False, False, False, False, False, False, False, False, False, False, False)
        ActiveSheet.PivotTables("tablaDinamica1").PivotFields("Apellido Nombre"). _
            Subtotals = Array(False, False, False, False, False, False, False, False, False, False, _
            False, False)
        ActiveSheet.PivotTables("tablaDinamica1").PivotFields("Per.para").Subtotals = _
            Array(False, False, False, False, False, False, False, False, False, False, False, False)
        ActiveSheet.PivotTables("tablaDinamica1").PivotFields("Fecha pago").Subtotals _
            = Array(False, False, False, False, False, False, False, False, False, False, False, False _
            )
        ActiveSheet.PivotTables("tablaDinamica1").PivotFields("CC-n.").Subtotals = _
            Array(False, False, False, False, False, False, False, False, False, False, False, False)
        ActiveSheet.PivotTables("tablaDinamica1").PivotFields("Texto expl.CC-nómina"). _
            Subtotals = Array(False, False, False, False, False, False, False, False, False, False, _
            False, False)
        ActiveSheet.PivotTables("tablaDinamica1").PivotFields("Cantidad").Subtotals = _
            Array(False, False, False, False, False, False, False, False, False, False, False, False)
        ActiveSheet.PivotTables("tablaDinamica1").PivotFields("Importe").Subtotals = _
            Array(False, False, False, False, False, False, False, False, False, False, False, False)
        ActiveSheet.PivotTables("tablaDinamica1").RowAxisLayout xlTabularRow
        
        
        'Copia y pega tabla dinamica como valores
        Columns("L:T").Select
        Selection.Copy
        Range("V1").Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
        Application.CutCopyMode = False
        Columns("X:AD").NumberFormat = "$#,##0"
        Columns("V:AD").AutoFit
        
        Columns("A:U").Select
        Selection.Delete
        
    'Pone las formulas de promedio
    LastRow = Cells(Rows.Count, 1).End(xlUp).row
    Sheets("BASES").Range("J4:J" & LastRow) = "=AVERAGE(RC[-7]:RC[-2])"
    Range("J2").Value = "PROMEDIO"
    Columns("J").NumberFormat = "$#,##0"
    Columns("A:J").AutoFit
   
    'Buscarv del proemedio
    Sheets("AUDITORIA").Range("S2:S" & LastRow) = "=VLOOKUP(RC[-18],BASES!C[-18]:C[-9],10,0)"
    Columns("J").NumberFormat = "$#,##0"
    Columns("R:T").AutoFit
    ActiveWorkbook.Save
    
    'Crea la hoja del LRN
    Workbooks.Open Ruta_Audi & "\" & "LNR-" & Mes_texto & ".XLSX"
    Workbooks("LNR-" & Mes_texto & ".XLSX").Activate
    Sheets(1).Activate
    Range("A:H").Copy
    wb_auditoria.Activate
    Sheets("LNR").Activate
    Range("A1").PasteSpecial Paste:=xlPasteAll
    Columns("A:H").AutoFit
    Workbooks("LNR-" & Mes_texto & ".XLSX").Save
    Workbooks("LNR-" & Mes_texto & ".XLSX").Close
    wb_auditoria.Activate
    Sheets("AUDITORIA").Activate
    LastRow = Cells(Rows.Count, 1).End(xlUp).row
    Sheets("AUDITORIA").Range("O2:O" & LastRow) = "=IFERROR(VLOOKUP(RC[-14],LNR!C[-14]:C[-8],7,0),0)"
    Columns("S:T").NumberFormat = "$#,##0"
    Columns("U").NumberFormat = "$#,##0.00"
    Columns("A:W").AutoFit

    
    'Resaltar diferencias
    Sheets("AUDITORIA").Activate
    LastRow = Cells(Rows.Count, 1).End(xlUp).row
    On Error Resume Next
    For Each cell In ActiveSheet.Range("U2:U" & LastRow)
        If IsError(cell.Value) Or (IsNumeric(cell.Value) And cell.Value <> 0) Then
            cell.Interior.Color = RGB(255, 255, 0)
        End If
    Next cell
    On Error GoTo 0
    
    ActiveWorkbook.Save
    ActiveWorkbook.Close
    
    
End Sub


Sub SAP_1028()
InicializarVariables
    
    'Conexion con SAP
    Dim SapGuiAuto As Object
    Dim App As Object
    Dim Connection As Object
    Dim session As Object
    
    Application.DisplayAlerts = False
    Set SapGuiAuto = GetObject("SAPGUI")
    Set App = SapGuiAuto.GetScriptingEngine
    Set Connection = App.Children(0)
    Set session = Connection.Children(0)
    

'-------------------- DESCARGA EL SEGUNDO REPORTE DE LA CWTR : CON CC NOM 1028 ---------------------
    'Vuelve a la pantalla inicial
    session.findById("wnd[0]/tbar[0]/okcd").Text = "/n"
    session.findById("wnd[0]").sendVKey 0

    'Entra a la transacion
    session.findById("wnd[0]/tbar[0]/okcd").Text = "/nPC00_M99_CWTR"
    session.findById("wnd[0]").sendVKey 0
    
    'Busca la variante
    session.findById("wnd[0]/tbar[1]/btn[17]").press
    session.findById("wnd[1]/usr/txtV-LOW").Text = "TCPRIMAEXTADC"
    session.findById("wnd[1]/usr/txtENAME-LOW").Text = "ASANC1K"
    session.findById("wnd[1]/usr/txtV-LOW").caretPosition = 14
    session.findById("wnd[1]/tbar[0]/btn[8]").press
    
    'Cambia fechas
    session.findById("wnd[0]/usr/ctxtBEGD_CAL").Text = inicio_sem
    session.findById("wnd[0]/usr/ctxtENDD_CAL").Text = fin_sem

    'Ejecuta
    session.findById("wnd[0]/tbar[1]/btn[8]").press
    
    'Descarga el archivo
    session.findById("wnd[0]/mbar/menu[0]/menu[1]/menu[2]").Select
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").Select
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").SetFocus
    session.findById("wnd[1]/tbar[0]/btn[0]").press
    session.findById("wnd[1]/usr/ctxtDY_PATH").Text = Ruta_Audi
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = "1028-" & Mes_texto & ".XLS"
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 16
    session.findById("wnd[1]/tbar[0]/btn[11]").press
    
    'Organizar documento
    Application.CutCopyMode = False
    Workbooks.Open Ruta_Audi & "\" & "1028-" & Mes_texto & ".XLS"
    ActiveSheet.Cells.Select
    Selection.Copy
    Workbooks.Add
    ActiveSheet.Paste
    ActiveWorkbook.SaveAs Ruta_Audi & "\" & "1028-" & Mes_texto & ".XLSX"
    ActiveWorkbook.Close SaveChanges:=True
    Workbooks("1028-" & Mes_texto & ".XLS").Close
    Kill Ruta_Audi & "\" & "1028-" & Mes_texto & ".XLS"
    
    'Cambios de formato
    Workbooks.Open Ruta_Audi & "\" & "1028-" & Mes_texto & ".XLSX"
    Workbooks("1028-" & Mes_texto & ".XLSX").Activate
    Rows("1").Delete
    Rows("2").Delete
    Columns("A").Delete
    Columns("A:J").AutoFit
    
    'Cambios de formato para la cantidad
    LastRow = ActiveSheet.Cells(ActiveSheet.Rows.Count, "A").End(xlUp).row
    Columns("G:G").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("G1").Value = "Cantidad"
    If LastRow >= 2 Then
        Range("G2:G" & LastRow).Formula = "=VALUE(SUBSTITUTE(RC[1], CHAR(160),""""))"
    End If
    Columns("G:G").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("H:H").Select
    Selection.Delete

    'Cambios de formato para el importe
    Columns("H:H").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("H1").Value = "Importe"
        If LastRow >= 2 Then
        Range("H2:H" & LastRow).Formula = "=VALUE(SUBSTITUTE(RC[1], CHAR(160),""""))"
    End If
    Columns("H:H").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("I:I").Select
    Selection.Delete
    Columns("H:H").NumberFormat = "$#,##0"
    Columns("A:H").AutoFit

    'Cambiar formato de fecha
    Dim CelsD As Range
    Dim UltimaFilaD As Long
    UltimaFilaD = ActiveSheet.Cells(Rows.Count, 1).End(xlUp).row
    
    If UltimaFilaD >= 2 Then
        For Each CelsD In ActiveSheet.Range("D2:D" & UltimaFilaD)
            If Len(CelsD.Value) = 10 And Mid(CelsD.Value, 3, 1) = "." And Mid(CelsD.Value, 6, 1) = "." Then
                CelsD.Value = DateSerial(Right(CelsD.Value, 4), Mid(CelsD.Value, 4, 2), Left(CelsD.Value, 2))
                CelsD.NumberFormat = "dd/mm/yyyy"
            End If
        Next CelsD
    End If
    ActiveWorkbook.Save
    ActiveWorkbook.Close
    
End Sub

Sub SAP_BASES()
InicializarVariables
    
    'Conexion con SAP
    Dim SapGuiAuto As Object
    Dim App As Object
    Dim Connection As Object
    Dim session As Object
    
    Application.DisplayAlerts = False
    Set SapGuiAuto = GetObject("SAPGUI")
    Set App = SapGuiAuto.GetScriptingEngine
    Set Connection = App.Children(0)
    Set session = Connection.Children(0)
    

'-------------------- DESCARGA EL SEGUNDO REPORTE DE LA CWTR : BASE PRIMA--------------------
    'Vuelve a la pantalla inicial
    session.findById("wnd[0]/tbar[0]/okcd").Text = "/n"
    session.findById("wnd[0]").sendVKey 0

    'Entra a la transacion
    session.findById("wnd[0]/tbar[0]/okcd").Text = "/nPC00_M99_CWTR"
    session.findById("wnd[0]").sendVKey 0
    
    'Busca la variante
    session.findById("wnd[0]/tbar[1]/btn[17]").press
    session.findById("wnd[1]/usr/txtV-LOW").Text = "TC_BASEPRIABS"
    session.findById("wnd[1]/usr/txtENAME-LOW").Text = "ASANC1K"
    session.findById("wnd[1]/usr/txtV-LOW").caretPosition = 14
    session.findById("wnd[1]/tbar[0]/btn[8]").press
    
    'Cambia fechas
    session.findById("wnd[0]/usr/ctxtBEGD_CAL").Text = inicio_sem
    session.findById("wnd[0]/usr/ctxtENDD_CAL").Text = fin_sem

    'Ejecuta
    session.findById("wnd[0]/tbar[1]/btn[8]").press
    
    'Descarga el archivo
    session.findById("wnd[0]/mbar/menu[0]/menu[1]/menu[2]").Select
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").Select
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").SetFocus
    session.findById("wnd[1]/tbar[0]/btn[0]").press
    session.findById("wnd[1]/usr/ctxtDY_PATH").Text = Ruta_Audi
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = "BASES PRIMA-" & Mes_texto & ".XLS"
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 16
    session.findById("wnd[1]/tbar[0]/btn[11]").press
    
    'Organizar documento
    Application.CutCopyMode = False
    Workbooks.Open Ruta_Audi & "\" & "BASES PRIMA-" & Mes_texto & ".XLS"
    ActiveSheet.Cells.Select
    Selection.Copy
    Workbooks.Add
    ActiveSheet.Paste
    ActiveWorkbook.SaveAs Ruta_Audi & "\" & "BASES PRIMA-" & Mes_texto & ".XLSX"
    ActiveWorkbook.Close SaveChanges:=True
    Workbooks("BASES PRIMA-" & Mes_texto & ".XLS").Close
    Kill Ruta_Audi & "\" & "BASES PRIMA-" & Mes_texto & ".XLS"
    
    'Cambios de formato
    Workbooks.Open Ruta_Audi & "\" & "BASES PRIMA-" & Mes_texto & ".XLSX"
    Workbooks("BASES PRIMA-" & Mes_texto & ".XLSX").Activate
    Rows("1").Delete
    Rows("2").Delete
    Columns("A").Delete
    Columns("A:J").AutoFit
    
    'Cambios de formato para la cantidad
    LastRow = ActiveSheet.Cells(ActiveSheet.Rows.Count, "A").End(xlUp).row
    Columns("G:G").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("G1").Value = "Cantidad"
    If LastRow >= 2 Then
        Range("G2:G" & LastRow).Formula = "=VALUE(SUBSTITUTE(RC[1], CHAR(160),""""))"
    End If
    Columns("G:G").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("H:H").Select
    Selection.Delete

    'Cambios de formato para el importe
    Columns("H:H").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("H1").Value = "Importe"
        If LastRow >= 2 Then
        Range("H2:H" & LastRow).Formula = "=VALUE(SUBSTITUTE(RC[1], CHAR(160),""""))"
    End If
    Columns("H:H").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("I:I").Select
    Selection.Delete
    Columns("H:H").NumberFormat = "$#,##0"
    Columns("A:H").AutoFit

    'Cambiar formato de fecha
    Dim CelsD As Range
    Dim UltimaFilaD As Long
    UltimaFilaD = ActiveSheet.Cells(Rows.Count, 1).End(xlUp).row
    
    If UltimaFilaD >= 2 Then
        For Each CelsD In ActiveSheet.Range("D2:D" & UltimaFilaD)
            If Len(CelsD.Value) = 10 And Mid(CelsD.Value, 3, 1) = "." And Mid(CelsD.Value, 6, 1) = "." Then
                CelsD.Value = DateSerial(Right(CelsD.Value, 4), Mid(CelsD.Value, 4, 2), Left(CelsD.Value, 2))
                CelsD.NumberFormat = "dd/mm/yyyy"
            End If
        Next CelsD
    End If
    ActiveWorkbook.Save
    ActiveWorkbook.Close
     
End Sub

Sub SAP_LNR()
InicializarVariables
    
    'Conexion con SAP
    Dim SapGuiAuto As Object
    Dim App As Object
    Dim Connection As Object
    Dim session As Object
    
    Application.DisplayAlerts = False
    Set SapGuiAuto = GetObject("SAPGUI")
    Set App = SapGuiAuto.GetScriptingEngine
    Set Connection = App.Children(0)
    Set session = Connection.Children(0)
    

'-------------------- DESCARGA EL SEGUNDO REPORTE DE LA CWTR : CON CC NOM 1028 ---------------------
    'Vuelve a la pantalla inicial
    session.findById("wnd[0]/tbar[0]/okcd").Text = "/n"
    session.findById("wnd[0]").sendVKey 0

    'Entra a la transacion
    session.findById("wnd[0]/tbar[0]/okcd").Text = "/nPC00_M99_CWTR"
    session.findById("wnd[0]").sendVKey 0
    
    'Busca la variante
    session.findById("wnd[0]/tbar[1]/btn[17]").press
    session.findById("wnd[1]/usr/txtV-LOW").Text = "TC_LNR_LIQUIDA"
    session.findById("wnd[1]/usr/txtENAME-LOW").Text = "ASANC1K"
    session.findById("wnd[1]/usr/txtV-LOW").caretPosition = 14
    session.findById("wnd[1]/tbar[0]/btn[8]").press
    
    'Cambia fechas
    session.findById("wnd[0]/usr/ctxtBEGD_CAL").Text = inicio_sem
    session.findById("wnd[0]/usr/ctxtENDD_CAL").Text = fin_sem

    'Ejecuta
    session.findById("wnd[0]/tbar[1]/btn[8]").press
    
    'Descarga el archivo
    session.findById("wnd[0]/mbar/menu[0]/menu[1]/menu[2]").Select
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").Select
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").SetFocus
    session.findById("wnd[1]/tbar[0]/btn[0]").press
    session.findById("wnd[1]/usr/ctxtDY_PATH").Text = Ruta_Audi
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = "LNR-" & Mes_texto & ".XLS"
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 16
    session.findById("wnd[1]/tbar[0]/btn[11]").press
    
    'Organizar documento
    Application.CutCopyMode = False
    Workbooks.Open Ruta_Audi & "\" & "LNR-" & Mes_texto & ".XLS"
    ActiveSheet.Cells.Select
    Selection.Copy
    Workbooks.Add
    ActiveSheet.Paste
    ActiveWorkbook.SaveAs Ruta_Audi & "\" & "LNR-" & Mes_texto & ".XLSX"
    ActiveWorkbook.Close SaveChanges:=True
    Workbooks("LNR-" & Mes_texto & ".XLS").Close
    Kill Ruta_Audi & "\" & "LNR-" & Mes_texto & ".XLS"
    
    'Cambios de formato
    Workbooks.Open Ruta_Audi & "\" & "LNR-" & Mes_texto & ".XLSX"
    Workbooks("LNR-" & Mes_texto & ".XLSX").Activate
    Rows("1").Delete
    Rows("2").Delete
    Columns("A").Delete
    Columns("A:J").AutoFit
    
    'Cambios de formato para la cantidad
    LastRow = ActiveSheet.Cells(ActiveSheet.Rows.Count, "A").End(xlUp).row
    Columns("G:G").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("G1").Value = "Cantidad"
    If LastRow >= 2 Then
        Range("G2:G" & LastRow).Formula = "=VALUE(SUBSTITUTE(RC[1], CHAR(160),""""))"
    End If
    Columns("G:G").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("H:H").Select
    Selection.Delete

    'Cambios de formato para el importe
    Columns("H:H").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("H1").Value = "Importe"
        If LastRow >= 2 Then
        Range("H2:H" & LastRow).Formula = "=VALUE(SUBSTITUTE(RC[1], CHAR(160),""""))"
    End If
    Columns("H:H").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("I:I").Select
    Selection.Delete
    Columns("H:H").NumberFormat = "$#,##0"
    Columns("A:H").AutoFit

    'Cambiar formato de fecha
    Dim CelsD As Range
    Dim UltimaFilaD As Long
    UltimaFilaD = ActiveSheet.Cells(Rows.Count, 1).End(xlUp).row
    
    If UltimaFilaD >= 2 Then
        For Each CelsD In ActiveSheet.Range("D2:D" & UltimaFilaD)
            If Len(CelsD.Value) = 10 And Mid(CelsD.Value, 3, 1) = "." And Mid(CelsD.Value, 6, 1) = "." Then
                CelsD.Value = DateSerial(Right(CelsD.Value, 4), Mid(CelsD.Value, 4, 2), Left(CelsD.Value, 2))
                CelsD.NumberFormat = "dd/mm/yyyy"
            End If
        Next CelsD
    End If
    ActiveWorkbook.Save
    ActiveWorkbook.Close
    
End Sub
Private Sub DeactivateStuff()
'SUBPROCEDURE EXPLANATION: Deactivates unwanted alerts, messages and pop-up windows.
   Application.IgnoreRemoteRequests = False
   Application.ScreenUpdating = False
   Application.DisplayAlerts = False
   Application.DisplayStatusBar = False
End Sub
Private Sub ReactivateStuff()
'SUBPROCEDURE EXPLANATION: Reactivates normal behavior of alerts, messages and pop-up windows.
   Application.ScreenUpdating = True
   Application.DisplayAlerts = True
   Application.DisplayStatusBar = True
   Application.IgnoreRemoteRequests = False
End Sub



