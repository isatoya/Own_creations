Attribute VB_Name = "PRIMA_EXTRA_ABS"
Option Explicit
'variables para todo el proyecto

    Dim MAESTRO As String
    Dim Fecha1, Fecha2, inicio_tri, fin_tri, mes, Mes_texto, Año, MesTrimestreAnterior, variante As String
    Dim ruta, Ruta_Año, Ruta_Mes, Ruta_Audi, ruta_maestroA As String
    Dim LastRow, i As Long
    Dim promedioFormula As String
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
    'ruta = ThisWorkbook.Sheets("Reportes").Range("H21").Value
    Ruta_Año = ruta & Año
    Ruta_Mes = Ruta_Año & "\" & mes & ". " & Mes_texto
    Ruta_Audi = Ruta_Mes & "\" & "AUDITORIAS DE NOMINA"
    
    
    '------------ Calcular fechas de los trimestres -----------
    ' Definir el primer día del trimestre anterior
    Select Case Mes_texto
        Case "Enero"
            inicio_tri = Format(DateSerial(Año - 1, 10, 1), "dd.mm.yyyy")
        Case "Abril"
            inicio_tri = Format(DateSerial(Año, 1, 1), "dd.mm.yyyy")
        Case "Julio"
            inicio_tri = Format(DateSerial(Año, 4, 1), "dd.mm.yyyy")
        Case "Octubre"
            inicio_tri = Format(DateSerial(Año, 7, 1), "dd.mm.yyyy")
    End Select
    
    ' Definir el último día del trimestre anterior y aplicar formato
    Select Case Mes_texto
        Case "Enero"
            fin_tri = Format(DateSerial(Año - 1, 12, 31), "dd.mm.yyyy")
        Case "Abril"
            fin_tri = Format(DateSerial(Año, 3, 31), "dd.mm.yyyy")
        Case "Julio"
            fin_tri = Format(DateSerial(Año, 6, 30), "dd.mm.yyyy")
        Case "Octubre"
            fin_tri = Format(DateSerial(Año, 9, 30), "dd.mm.yyyy")
    End Select

    '------------ Calcula meses de los trimestres -----------
    Select Case Mes_texto
        Case "Enero"
            MesTrimestreAnterior = "10,11,12"
        Case "Abril"
            MesTrimestreAnterior = "01,02,03"
        Case "Julio"
            MesTrimestreAnterior = "04,05,06"
        Case "Octubre"
            MesTrimestreAnterior = "07,08,09"
    End Select


End Sub

Sub Ejecutar_AUDI_PRIMAEXTRA_ABS()

    ' Verificar si hay datos en las celdas I8 y M8
    If ThisWorkbook.Sheets("Reportes").Range("I8").Value = "" Or ThisWorkbook.Sheets("Reportes").Range("M8").Value = "" Then
        MsgBox "Datos incompletos, por favor ingrese los datos antes de ejecutar.", vbExclamation
        Exit Sub
    End If

    ' Llama a cada una de las funciones
    DeactivateStuff
    CrearCarpetas
    Bases
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

Sub Bases()
InicializarVariables

'-------------------- DESCARGA Y ORGANIZA EL REPORTE DE CWTR --------------------
    
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
    session.findById("wnd[0]/tbar[0]/okcd").Text = "/nPC00_M99_CWTR"
    session.findById("wnd[0]").sendVKey 0
    
    'Busca la variante
    session.findById("wnd[0]/tbar[1]/btn[17]").press
    session.findById("wnd[1]/usr/txtV-LOW").Text = "TC_BASEPRIABS"
    session.findById("wnd[1]/usr/txtENAME-LOW").Text = "ASANC1K"
    session.findById("wnd[1]/usr/txtV-LOW").caretPosition = 14
    session.findById("wnd[1]/tbar[0]/btn[8]").press
    
    'Cambia fechas
    session.findById("wnd[0]/usr/ctxtBEGD_CAL").Text = inicio_tri
    session.findById("wnd[0]/usr/ctxtENDD_CAL").Text = fin_tri

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
    ActiveWorkbook.Save
    
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
    
    'Agregar hojas
    Sheets.Add(After:=Sheets(Sheets.Count)).Name = "BASES"
    Sheets.Add(After:=Sheets(Sheets.Count)).Name = "BASE TRIMESTRE"
    
    'Crea la hoja de Bases
    Sheets(1).Activate
    Range("A:H").Copy
    Sheets("BASES").Activate
    Range("A1").PasteSpecial Paste:=xlPasteAll
    Columns("A:H").AutoFit
    Application.CutCopyMode = False
    
    'Elimina lo que tenga en periodo para 0
    Sheets("BASES").Activate
    LastRow = Cells(Rows.Count, "A").End(xlUp).row
    For i = LastRow To 1 Step -1
        If ActiveSheet.Cells(i, "C").Value = "0" Then
            ActiveSheet.Rows(i).Delete
        End If
    Next i
    
    'Elimina los de meses que no esten en el trimestre
    Columns("D:D").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    LastRow = ActiveSheet.Cells(ActiveSheet.Rows.Count, "A").End(xlUp).row
    Range("D2:D" & LastRow) = "=MID(RC[-1],5,2)"
    Columns("D:D").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    For i = LastRow To 1 Step -1
        If Not InStr(1, MesTrimestreAnterior, ActiveSheet.Cells(i, "D").Value) > 0 Then
            ActiveSheet.Rows(i).Delete
        End If
    Next i
    Columns("D:D").Select
    Selection.Delete
    
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
        Columns("L:P").Select
        Selection.Copy
        Sheets("BASE TRIMESTRE").Activate
        Range("A1").Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
        Application.CutCopyMode = False
        Columns("B:F").NumberFormat = "$#,##0"
        Columns("A:E").AutoFit
        
        
    'Pone las formulas de promedio
    LastRow = Cells(Rows.Count, 1).End(xlUp).row
    Sheets("BASE TRIMESTRE").Range("F4:F" & LastRow) = "=IF(COUNTBLANK(RC[-3]:RC[-1])=0, AVERAGE(RC[-3]:RC[-1]), IF(COUNTBLANK(RC[-3])=1, AVERAGE(RC[-2]:RC[-1]), IF(COUNTBLANK(RC[-2])=1, AVERAGE(RC[-3],RC[-1]), IF(COUNTBLANK(RC[-1])=1, AVERAGE(RC[-3],RC[-2]), IF(AND(COUNTBLANK(RC[-3])=0, COUNTBLANK(RC[-2])=0), AVERAGE(RC[-3],RC[-2]), IF(AND(COUNTBLANK(RC[-3])=0, COUNTBLANK(RC[-1])=0), AVERAGE(RC[-3],RC[-1]), IF(AND(COUNT" & _
        "BLANK(RC[-2])=0, COUNTBLANK(RC[-1])=0), AVERAGE(RC[-3]:RC[-2]), IF(AND(COUNTBLANK(RC[-3])=0, COUNTBLANK(RC[-2])=0, COUNTBLANK(RC[-1])=0), AVERAGE(RC[-3]:RC[-1]), 0))))))))" & _
        ""
    ActiveWorkbook.Save
    
End Sub

Sub Auditoria()
InicializarVariables
Dim MAESTRO As Workbook

'-------------------- DESCARGA Y ORGANIZA EL REPORTE DE CWTR --------------------
    
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
    session.findById("wnd[0]/tbar[0]/okcd").Text = "/nPC00_M99_CWTR"
    session.findById("wnd[0]").sendVKey 0
    
    'Busca la variante
    Select Case Mes_texto
        Case "Enero"
            variante = "TC_PRIMAS_ENER"
        Case "Abril"
            variante = "TC_PRIMA_PERMA"
        Case "Julio"
            variante = "TC_PRIMA_JULAB"
        Case "octubre"
            variante = "TC_PRIM_OCTABS"
        Case Else
            variante = ""
    End Select
    
    session.findById("wnd[0]/tbar[1]/btn[17]").press
    session.findById("wnd[1]/usr/txtV-LOW").Text = variante
    session.findById("wnd[1]/usr/txtENAME-LOW").Text = "ASANC1K"
    session.findById("wnd[1]/usr/txtV-LOW").caretPosition = 14
    session.findById("wnd[1]/tbar[0]/btn[8]").press
    
    'Cambia fechas
    session.findById("wnd[0]/usr/ctxtBEGD_CAL").Text = Fecha1
    session.findById("wnd[0]/usr/ctxtENDD_CAL").Text = Fecha2

    'Ejecuta
    session.findById("wnd[0]/tbar[1]/btn[8]").press
    
    'Descarga el archivo
    session.findById("wnd[0]/mbar/menu[0]/menu[1]/menu[2]").Select
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").Select
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").SetFocus
    session.findById("wnd[1]/tbar[0]/btn[0]").press
    session.findById("wnd[1]/usr/ctxtDY_PATH").Text = Ruta_Audi
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = Año & mes & "." & "AUDITORIA PRIMA-" & Mes_texto & ".XLS"
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 16
    session.findById("wnd[1]/tbar[0]/btn[11]").press
    
    'Organizar documento
    Application.CutCopyMode = False
    Workbooks.Open Ruta_Audi & "\" & Año & mes & "." & "AUDITORIA PRIMA-" & Mes_texto & ".XLS"
    ActiveSheet.Cells.Select
    Selection.Copy
    Workbooks.Add
    ActiveSheet.Paste
    ActiveWorkbook.SaveAs Ruta_Audi & "\" & Año & mes & "." & "AUDITORIA PRIMA-" & Mes_texto & ".XLSX"
    ActiveWorkbook.Close SaveChanges:=True
    Workbooks(Año & mes & "." & "AUDITORIA PRIMA-" & Mes_texto & ".XLS").Close
    Kill Ruta_Audi & "\" & Año & mes & "." & "AUDITORIA PRIMA-" & Mes_texto & ".XLS"

     'Cambios de formato
    Workbooks.Open Ruta_Audi & "\" & Año & mes & "." & "AUDITORIA PRIMA-" & Mes_texto & ".XLSX"
    Workbooks(Año & mes & "." & "AUDITORIA PRIMA-" & Mes_texto & ".XLSX").Activate
    Rows("1").Delete
    Rows("2").Delete
    Columns("A").Delete
    Columns("A:H").AutoFit
    
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
    
    'Crea Hojas
    Workbooks(Año & mes & "." & "AUDITORIA PRIMA-" & Mes_texto & ".XLSX").Activate
    ActiveSheet.Name = "CWTR"
    Sheets.Add(After:=Sheets(Sheets.Count)).Name = "BASES"
    Sheets.Add(After:=Sheets(Sheets.Count)).Name = "MAESTRO"
    Sheets.Add(After:=Sheets(Sheets.Count)).Name = "AUDITORIA"
    
    'Trae las bases y el promedio del otro archivo
    Workbooks("BASES PRIMA-" & Mes_texto & ".XLSX").Activate
    Sheets("BASE TRIMESTRE").Activate
    Columns("A:F").Select
    Selection.Copy
    Workbooks(Año & mes & "." & "AUDITORIA PRIMA-" & Mes_texto & ".XLSX").Activate
    Sheets("BASES").Activate
    Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    Columns("A:F").AutoFit
    Workbooks(Año & mes & "." & "AUDITORIA PRIMA-" & Mes_texto & ".XLSX").Save
    Workbooks("BASES PRIMA-" & Mes_texto & ".XLSX").Save
    Workbooks("BASES PRIMA-" & Mes_texto & ".XLSX").Close
    
    'Va a traer los datos del archivo de maestro de activos
    MsgBox "Por favor abra el maestro de activos correspondiente.", vbInformation
    ruta_maestroA = Application.GetOpenFilename("Archivos Excel (*.xls; *.xlsx), *.xls; *.xlsx")
    Application.AskToUpdateLinks = False
    
    ' Abre el reporte seleccionado
    If ruta_maestroA <> "Falso" Then
            
            Set MAESTRO = Workbooks.Open(Filename:=ruta_maestroA, UpdateLinks:=0)
            MAESTRO.Activate
            Worksheets("SALARIAL").Activate
            On Error Resume Next
            ActiveSheet.ShowAllData
            On Error GoTo 0
            Columns("A:Y").Select
            Selection.Copy
            
            'Trae la columna necesaria
            Workbooks(Año & mes & "." & "AUDITORIA PRIMA-" & Mes_texto & ".XLSX").Activate
            Sheets("MAESTRO").Activate
            ActiveSheet.Range("A1").PasteSpecial Paste:=xlPasteAll
            Columns("A:Y").AutoFit
            Application.CutCopyMode = False
            Workbooks(Año & mes & "." & "AUDITORIA PRIMA-" & Mes_texto & ".XLSX").Save
            MAESTRO.Save
            MAESTRO.Close
        
    End If
    
    'Filtra hoja de maestro
    Workbooks(Año & mes & "." & "AUDITORIA PRIMA-" & Mes_texto & ".XLSX").Activate
    Sheets("MAESTRO").Activate
    Rows("1:1").AutoFilter
    Rows("1:1").AutoFilter Field:=6, Criteria1:="=Ley 50"
    Range("1:1").AutoFilter Field:=7, Criteria1:="=CU"
    Range("1:1").AutoFilter Field:=25, Criteria1:="<>*TRAINEE"
    
    
    'Copia la hoja filtrada
    ActiveSheet.Columns("A:Q").SpecialCells(xlCellTypeVisible).Copy
    Sheets("AUDITORIA").Activate
    Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    With Range("A1:Y1")
        .Interior.Color = RGB(174, 214, 241)
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
    End With
    Range("R1").Value = "F.INICIO"
    Range("S1").Value = "F.FIN"
    Range("T1").Value = "DIAS LAB"
    Range("U1").Value = "PROMEDIO"
    Range("V1").Value = "CALCULO"
    Range("W1").Value = "CWTR"
    Range("X1").Value = "DIFERENCIA"
    Range("Y1").Value = "OBSERVACION"
    
    Columns("K:K").NumberFormat = "$#,##0"
    Columns("M:M").NumberFormat = "$#,##0"
    Columns("N:N").NumberFormat = "$#,##0"
    Columns("Q:Q").NumberFormat = "dd/mm/yyyy"

    
    'Calcula fecha inicio, primero borra lo que no sea del año presente
        
    If Mes_texto = "Enero" Then  'Elimina datos del año anterior

        Columns("Q:Q").Select
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        LastRow = ActiveSheet.Cells(ActiveSheet.Rows.Count, "A").End(xlUp).row
        Range("Q2:Q" & LastRow) = "=YEAR(RC[1])"
        Columns("Q:Q").Select
        Selection.Copy
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
        
        Rows("1:1").AutoFilter
        Range("1:1").AutoFilter Field:=17, Criteria1:=Año
        LastRow = ActiveSheet.Cells(ActiveSheet.Rows.Count, "A").End(xlUp).row
        Rows("2:" & LastRow).Select
        ' Seleccionar solo las filas visibles
        On Error Resume Next
        Set visibleRange = Selection.SpecialCells(xlCellTypeVisible)
        On Error GoTo 0
        
        ' Eliminar las filas visibles si existen
        If Not visibleRange Is Nothing Then
            visibleRange.Delete
        End If
        ActiveSheet.ShowAllData
    
    Else
        
        'extre el numero del mes
        Columns("Q:Q").Select
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        LastRow = ActiveSheet.Cells(ActiveSheet.Rows.Count, "A").End(xlUp).row
        Range("Q2:Q" & LastRow) = "=MONTH(RC[1])"
        Columns("Q:Q").Select
        Selection.Copy
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
        Columns("Q:Q").NumberFormat = "00"
        Range("Q1").Value = "Mes"
        
        'extrae el año
        Columns("Q:Q").Select
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        LastRow = ActiveSheet.Cells(ActiveSheet.Rows.Count, "A").End(xlUp).row
        Range("Q2:Q" & LastRow) = "=YEAR(RC[2])"
        Columns("Q:Q").Select
        Selection.Copy
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
        Range("Q1").Value = "Año"
        
        'Filtra y elimina las celdas que cumplan con es condicion
        Rows("1:1").AutoFilter
        Rows("1:1").AutoFilter Field:=17, Criteria1:=Año
        Range("1:1").AutoFilter Field:=18, Criteria1:=mes
        LastRow = ActiveSheet.Cells(ActiveSheet.Rows.Count, "A").End(xlUp).row
        Rows("2:" & LastRow).Select
    
        ' Seleccionar solo las filas visibles
        On Error Resume Next
        Set visibleRange = Selection.SpecialCells(xlCellTypeVisible)
        On Error GoTo 0
        
        ' Eliminar las filas visibles si existen
        If Not visibleRange Is Nothing Then
            visibleRange.Delete
        End If
        ActiveSheet.ShowAllData
      
        Columns("Q:R").Select
        Selection.Delete
    
    End If
    
    'Calcula la fecha de inicio
    LastRow = ActiveSheet.Cells(ActiveSheet.Rows.Count, "A").End(xlUp).row
    Columns("Q:Q").NumberFormat = "dd/mm/yyyy"
    Columns("R:R").NumberFormat = "dd/mm/yyyy"

    Select Case Mes_texto
        Case "Enero"
            Sheets("AUDITORIA").Range("R2:R" & LastRow) = "=IF(RC[-1]<(DATE(" & Año - 1 & "," & 9 & ",30)),(DATE(" & Año - 1 & "," & 10 & ",01)),RC[-1])"
        Case "Abril"
            Sheets("AUDITORIA").Range("R2:R" & LastRow) = "=IF(RC[-1]>(DATE(" & Año - 1 & "," & 12 & ",31)),RC[-1],(DATE(" & Año & "," & 1 & ",01)))"
        Case "Julio"
            Sheets("AUDITORIA").Range("R2:R" & LastRow) = "=IF(RC[-1]>(DATE(" & Año & "," & 3 & ",31)),RC[-1],(DATE(" & Año & "," & 4 & ",01)))"
        Case "Octubre"
            Sheets("AUDITORIA").Range("R2:R" & LastRow) = "=IF(RC[-1]>(DATE(" & Año & "," & 6 & ",30)),RC[-1],(DATE(" & Año & "," & 7 & ",01)))"
        Case Else
    End Select
    
    'Calcula la fecha de fin
    Columns("S:S").NumberFormat = "dd/mm/yyyy"
    Select Case Mes_texto
        Case "Enero"
            Sheets("AUDITORIA").Range("S2:S" & LastRow) = "=DATE(" & Año - 1 & ",12,31)"
        Case "Abril"
            Sheets("AUDITORIA").Range("S2:S" & LastRow) = "=DATE(" & Año & ",03,31)"
        Case "Julio"
            Sheets("AUDITORIA").Range("S2:S" & LastRow) = "=DATE(" & Año & ",06,30)"
        Case "Octubre"
            Sheets("AUDITORIA").Range("S2:S" & LastRow) = "=DATE(" & Año & ",09,30)"
        Case Else
    End Select
    
    'Calculo de dias laborados
    If Mes_texto = "Abril" Or Mes_texto = "Enero" Then
        Sheets("AUDITORIA").Range("T2:T" & LastRow) = "=DAYS360(RC[-2],RC[-1],0)"
    Else
        Sheets("AUDITORIA").Range("T2:T" & LastRow) = "=DAYS360(RC[-2],RC[-1],0)+1"
    End If
    
    'Buscarv de los promedios
    Columns("U:U").NumberFormat = "$#,##0"
    Sheets("AUDITORIA").Range("U2:U" & LastRow) = "=IFERROR(VLOOKUP(RC[-20],BASES!C[-20]:C[-15],6,0), ""0"")"
    
    'Calculo
    Columns("V:V").NumberFormat = "$#,##0"
    Sheets("AUDITORIA").Range("V2:V" & LastRow) = "=RC[-2]*RC[-1]/90"
    
    'CWTR
    Columns("W:W").NumberFormat = "$#,##0"
    Sheets("AUDITORIA").Range("W2:W" & LastRow) = "=IFERROR(VLOOKUP(RC[-22],CWTR!C[-22]:C[-15],8,0), ""0"")"
    
    'Diferencias
    Columns("X:X").NumberFormat = "0.00"
    Sheets("AUDITORIA").Range("X2:X" & LastRow) = "=RC[-2]-RC[-1]"
    Columns("A:Y").AutoFit
    
    'Pone la observacion de revisar
    On Error Resume Next
    For i = 2 To LastRow
    If IsError(Cells(i, "X").Value) Or Cells(i, "X").Value <> 0 Then
        Cells(i, "Y").Value = "Revisar"
    End If
    Next i
    On Error GoTo 0
    ActiveWorkbook.Save
    
    '---------- Termina de hacer cambios en la hoja de autoria
    '---------- Inicia cambios en la CWTR
    
    'Buscarv
    Sheets("CWTR").Activate
    LastRow = ActiveSheet.Cells(ActiveSheet.Rows.Count, "A").End(xlUp).row
    Range("I1").Value = "MAESTRO"
    Range("I1").Interior.Color = RGB(252, 243, 207)
    Range("I1").Font.Bold = True
    Range("I1").HorizontalAlignment = xlCenter
    Sheets("CWTR").Range("I2:I" & LastRow) = "=VLOOKUP(RC[-8],AUDITORIA!C[-8]:C[13],1,0)"
    Workbooks(Año & mes & "." & "AUDITORIA PRIMA-" & Mes_texto & ".XLSX").Save
    Workbooks(Año & mes & "." & "AUDITORIA PRIMA-" & Mes_texto & ".XLSX").Close
    
    
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



