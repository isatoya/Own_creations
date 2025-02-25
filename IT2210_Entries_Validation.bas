Attribute VB_Name = "IT2210_Entries_Validation"
Option Explicit
'variables para todo el proyecto

Dim wbauditoria, archivoBase As Workbook
Dim hojaBase, hojaOrigen As Worksheet
Dim Fecha1, Fecha2, Fecha3, mes, Mes_Texto, BPA, año, dia, auditoria, ruta, Ruta_pais, Ruta_Año, Ruta_Audi, ruta_Base, texto, ExisteArchivo As String
Dim hoja As Integer
Dim lastRow, lastCol, i As Long
Dim titulos As Variant
Dim rng As Range

Sub Ejecutar_IT221()

' Verificar si hay datos en las celdas I8 y M8
If ThisWorkbook.Sheets("Home Page").Range("I8").Value = "" Or ThisWorkbook.Sheets("Home Page").Range("E18").Value = "" Then
    MsgBox "Incomplete data, please enter the data before executing.", vbExclamation
    Exit Sub
End If

' Llama a cada una de las funciones
InicializarVariables
CrearCarpetas
Desarrollo_Entries_Validation

End Sub

Sub InicializarVariables()
'Definicion de las variables
    
'Fechas
mes = ThisWorkbook.Sheets("Home Page").Range("N8").Text
Mes_Texto = ThisWorkbook.Sheets("Home Page").Range("I12").Value
año = ThisWorkbook.Sheets("Home Page").Range("I10").Value
Fecha1 = ThisWorkbook.Sheets("Home Page").Range("I8").Value
Fecha2 = ThisWorkbook.Sheets("Home Page").Range("M8").Value
dia = ThisWorkbook.Sheets("Home Page").Range("N10").Text
BPA = ThisWorkbook.Sheets("Home Page").Range("E18").Text
auditoria = ThisWorkbook.Sheets("Home Page").Range("M17").Text

'Rutas
ruta = ThisWorkbook.Path & "\"
Ruta_pais = ruta & "YEAR END CA"
Ruta_Audi = Ruta_pais & "\" & "ENTRIES VALIDATION"
Ruta_Año = Ruta_Audi & "\" & año

End Sub
Sub CrearCarpetas()

'PAIS
Ruta_pais = ruta & "YEAR END CA"
If Dir(Ruta_pais, vbDirectory + vbHidden) = "" Then
    If Dir(Ruta_pais & vbDirectory + vbHidden) = "" Then MkDir Ruta_pais
End If

'CARPETA DE AUDITORIAS
Ruta_Audi = Ruta_pais & "\" & "ENTRIES VALIDATION"
If Dir(Ruta_Audi, vbDirectory + vbHidden) = "" Then
    If Dir(Ruta_Audi & vbDirectory + vbHidden) = "" Then MkDir Ruta_Audi
End If

'CARPETA AÑO
Ruta_Año = Ruta_Audi & "\" & año
If Dir(Ruta_Año, vbDirectory + vbHidden) = "" Then
    If Dir(Ruta_Audi & vbDirectory + vbHidden) = "" Then MkDir Ruta_Año
End If
     
End Sub

Sub Desarrollo_Entries_Validation()

'Crear un nuevo libro de Excel y le pone el nombre correspondiente a la auditoria
Set wbauditoria = Workbooks.Add
wbauditoria.SaveAs Ruta_Año & "\" & año & mes & " " & auditoria & ".xlsx"

'Crea las hojas
wbauditoria.Sheets.Add(After:=wbauditoria.Sheets(wbauditoria.Sheets.Count)).Name = "BoxJ entry"
wbauditoria.Sheets.Add(After:=wbauditoria.Sheets(wbauditoria.Sheets.Count)).Name = "CCYR T4 before"
wbauditoria.Sheets.Add(After:=wbauditoria.Sheets(wbauditoria.Sheets.Count)).Name = "PT T4 Before"
wbauditoria.Sheets.Add(After:=wbauditoria.Sheets(wbauditoria.Sheets.Count)).Name = "PT T4 Before-values"
wbauditoria.Sheets.Add(After:=wbauditoria.Sheets(wbauditoria.Sheets.Count)).Name = "CCYR T4 After"
wbauditoria.Sheets.Add(After:=wbauditoria.Sheets(wbauditoria.Sheets.Count)).Name = "PT T4 After"
wbauditoria.Sheets.Add(After:=wbauditoria.Sheets(wbauditoria.Sheets.Count)).Name = "PT T4 After-values"
wbauditoria.Sheets.Add(After:=wbauditoria.Sheets(wbauditoria.Sheets.Count)).Name = "CCYR RL1 before"
wbauditoria.Sheets.Add(After:=wbauditoria.Sheets(wbauditoria.Sheets.Count)).Name = "PT RL1 Before"
wbauditoria.Sheets.Add(After:=wbauditoria.Sheets(wbauditoria.Sheets.Count)).Name = "PT RL1 Before-values"
wbauditoria.Sheets.Add(After:=wbauditoria.Sheets(wbauditoria.Sheets.Count)).Name = "CCYR RL1 After"
wbauditoria.Sheets.Add(After:=wbauditoria.Sheets(wbauditoria.Sheets.Count)).Name = "PT RL1 After"
wbauditoria.Sheets.Add(After:=wbauditoria.Sheets(wbauditoria.Sheets.Count)).Name = "PT RL1 After-values"
wbauditoria.Sheets.Add(After:=wbauditoria.Sheets(wbauditoria.Sheets.Count)).Name = "Validation RL1"
wbauditoria.Sheets.Add(After:=wbauditoria.Sheets(wbauditoria.Sheets.Count)).Name = "Validation T4"
Application.DisplayAlerts = False
Sheets(1).Delete
Application.DisplayAlerts = True


'---------------------------------- ABRE ARCHIVO DE T4 (BEFORE) ----------------------------------
MsgBox "Please select the T4 (BEFORE) database file downloaded from SAP", vbInformation
ruta_Base = Application.GetOpenFilename("Archivos Excel (*.xls; *.xlsx), *.xls; *.xlsx")
Application.AskToUpdateLinks = False

' Verifica si el usuario seleccionó un archivo
If ruta_Base <> "Falso" Then
    'Abre el archivo base
    Set archivoBase = Workbooks.Open(Filename:=ruta_Base, UpdateLinks:=0)
    archivoBase.Activate
    Set hojaOrigen = archivoBase.Sheets(1)
    
    'Filtra solo los TTA
    Sheets(1).Activate
    lastRow = Cells(Rows.Count, "A").End(xlUp).row
    Rows("1:1").AutoFilter
    Rows("1:1").AutoFilter Field:=14, Criteria1:="=#TTA"
    Range("A1:P" & lastRow).SpecialCells(xlCellTypeVisible).Copy
    wbauditoria.Sheets("CCYR T4 before").Cells(1, 1).PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    
    'Guarda el archivo nuevo y cierra el de la base
    wbauditoria.Save
    archivoBase.Close SaveChanges:=False
    
Else
End If

'Poner todos los boxes para que el rango no cambie
lastRow = Sheets("CCYR T4 before").Cells(Sheets("CCYR T4 before").Rows.Count, 1).End(xlUp).row
lastCol = Sheets("CCYR T4 before").Cells(2, Sheets("CCYR T4 before").Columns.Count).End(xlToLeft).Column

ThisWorkbook.Sheets("Anexxes").Range("H2:H25").Copy
wbauditoria.Sheets("CCYR T4 before").Cells(lastRow + 1, "L").PasteSpecial Paste:=xlPasteValues
Application.CutCopyMode = False
ThisWorkbook.Sheets("Anexxes").Range("I2:I25").Copy
wbauditoria.Sheets("CCYR T4 before").Cells(lastRow + 1, "A").PasteSpecial Paste:=xlPasteValues
Application.CutCopyMode = False


'Crea tabla dinamica
Sheets("CCYR T4 before").Activate

    Dim ult_Tabla As Long
    ult_Tabla = Cells(Rows.Count, "A").End(xlUp).row
    Dim rangoTabla1 As Range
    Set rangoTabla1 = Sheets("CCYR T4 before").Range("A1:P" & ult_Tabla)
    ActiveSheet.ListObjects.Add(xlSrcRange, rangoTabla1, , xlYes).Name = "Tabla1"
    
    'Crear tabla dinamica
    Dim celdaTablaDinamica1 As Range
    Set celdaTablaDinamica1 = Sheets("PT T4 Before").Range("A1")
    Dim tablaDinamica1 As PivotTable
    
    'Activa campos y le pone formato tabular
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        rangoTabla1, Version:=6).CreatePivotTable TableDestination:= _
        celdaTablaDinamica1, TableName:="tablaDinamica1", DefaultVersion:=6
        
    Sheets("PT T4 Before").Select
    With ActiveSheet.PivotTables("tablaDinamica1").PivotFields("PERNR")
        .Orientation = xlRowField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("tablaDinamica1").PivotFields("BUSNM")
        .Orientation = xlRowField
        .Position = 2
    End With
    With ActiveSheet.PivotTables("tablaDinamica1").PivotFields("NACHN")
        .Orientation = xlRowField
        .Position = 3
    End With
    With ActiveSheet.PivotTables("tablaDinamica1").PivotFields("VORNA")
        .Orientation = xlRowField
        .Position = 4
    End With
    With ActiveSheet.PivotTables("tablaDinamica1").PivotFields("WRKAR")
        .Orientation = xlRowField
        .Position = 5
    End With
    With ActiveSheet.PivotTables("tablaDinamica1").PivotFields("SLART")
        .Orientation = xlColumnField
        .Position = 1
    End With
    ActiveSheet.PivotTables("tablaDinamica1").AddDataField ActiveSheet.PivotTables( _
        "tablaDinamica1").PivotFields("BETRG"), "Cuenta de BETRG", xlCount
    With ActiveSheet.PivotTables("tablaDinamica1").PivotFields("Cuenta de BETRG")
        .Caption = "Suma de BETRG"
        .Function = xlSum
    End With
    ActiveSheet.PivotTables("tablaDinamica1").PivotFields("PERNR").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("tablaDinamica1").PivotFields("PERSONID").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("tablaDinamica1").PivotFields("FORML").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("tablaDinamica1").PivotFields("INDX1").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("tablaDinamica1").PivotFields("BUSNM").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("tablaDinamica1").PivotFields("WRKAR").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("tablaDinamica1").PivotFields("WTEXT").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("tablaDinamica1").PivotFields("KOSTL").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("tablaDinamica1").PivotFields("KTEXT").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("tablaDinamica1").PivotFields("NACHN").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("tablaDinamica1").PivotFields("VORNA").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("tablaDinamica1").PivotFields("SLART").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("tablaDinamica1").PivotFields("STEXT").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("tablaDinamica1").PivotFields("LGART").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("tablaDinamica1").PivotFields("LTEXT").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("tablaDinamica1").PivotFields("BETRG").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("tablaDinamica1").RowAxisLayout xlTabularRow
    ActiveSheet.PivotTables("tablaDinamica1").RepeatAllLabels xlRepeatLabels
    
    
wbauditoria.Save

'Copia y pega los datos en la hoja de valores
lastRow = Sheets("PT T4 Before").Cells(Sheets("PT T4 Before").Rows.Count, 1).End(xlUp).row
lastCol = Sheets("PT T4 Before").Cells(2, Sheets("PT T4 Before").Columns.Count).End(xlToLeft).Column
Sheets("PT T4 Before").Range(Cells(2, 1), Cells(lastRow - 1, lastCol)).Copy
Sheets("PT T4 Before-values").Range("B1").PasteSpecial Paste:=xlPasteValues
Application.CutCopyMode = False

'Texto a numero
Sheets("PT T4 Before-values").Activate
lastRow = Cells(Rows.Count, "B").End(xlUp).row
For i = 2 To lastRow
    If IsNumeric(Range("b" & i).Value) Then
        Range("B" & i).Value = Val(Range("B" & i).Value)
    End If
Next i

'Formula del concatenar el KEY NUMBER
Range("A1").Value = "KEY NUMBER"
lastRow = Cells(Rows.Count, "B").End(xlUp).row
Sheets("PT T4 Before-values").Range("A2:A" & lastRow) = "=+CONCATENATE(RC[1],RC[2])"
Columns("A:AB").AutoFit

'Elimina los numeros de empleado que son 00000000
lastRow = Cells(Rows.Count, "A").End(xlUp).row
For i = lastRow To 2 Step -1
        If Cells(i, "B").Value = "0" Then
            Rows(i).Delete
        End If
Next i


lastRow = Cells(Rows.Count, "A").End(xlUp).row
For i = lastRow To 2 Step -1
        If Cells(i, "B").Value = "999999" Then
            Rows(i).Delete
        End If
Next i

wbauditoria.Save


'---------------------------------- ABRE ARCHIVO DE T4 (AFTER) ----------------------------------
MsgBox "Please select the T4 (AFTER) database file downloaded from SAP", vbInformation
ruta_Base = Application.GetOpenFilename("Archivos Excel (*.xls; *.xlsx), *.xls; *.xlsx")
Application.AskToUpdateLinks = False

' Verifica si el usuario seleccionó un archivo
If ruta_Base <> "Falso" Then
    'Abre el archivo base
    Set archivoBase = Workbooks.Open(Filename:=ruta_Base, UpdateLinks:=0)
    archivoBase.Activate
    Set hojaOrigen = archivoBase.Sheets(1)
    
    'Filtra solo los TTA
    Sheets(1).Activate
    lastRow = Cells(Rows.Count, "A").End(xlUp).row
    Rows("1:1").AutoFilter
    Rows("1:1").AutoFilter Field:=14, Criteria1:="=#TTA"
    Range("A1:P" & lastRow).SpecialCells(xlCellTypeVisible).Copy
    wbauditoria.Sheets("CCYR T4 After").Cells(1, 1).PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    
    'Guarda el archivo nuevo y cierra el de la base
    wbauditoria.Save
    archivoBase.Close SaveChanges:=False
    
Else
End If

'Poner todos los boxes para que el rango no cambie
lastRow = Sheets("CCYR T4 After").Cells(Sheets("CCYR T4 After").Rows.Count, 1).End(xlUp).row
lastCol = Sheets("CCYR T4 After").Cells(2, Sheets("CCYR T4 After").Columns.Count).End(xlToLeft).Column

ThisWorkbook.Sheets("Anexxes").Range("H2:H25").Copy
wbauditoria.Sheets("CCYR T4 After").Cells(lastRow + 1, "L").PasteSpecial Paste:=xlPasteValues
Application.CutCopyMode = False
ThisWorkbook.Sheets("Anexxes").Range("I2:I25").Copy
wbauditoria.Sheets("CCYR T4 After").Cells(lastRow + 1, "A").PasteSpecial Paste:=xlPasteValues
Application.CutCopyMode = False

'Crea tabla dinamica
Sheets("CCYR T4 After").Activate

    ult_Tabla = Cells(Rows.Count, "A").End(xlUp).row
    Set rangoTabla1 = Sheets("CCYR T4 After").Range("A1:P" & ult_Tabla)
    ActiveSheet.ListObjects.Add(xlSrcRange, rangoTabla1, , xlYes).Name = "Tabla1"
    
    'Crear tabla dinamica
    Set celdaTablaDinamica1 = Sheets("PT T4 After").Range("A1")
    
    'Activa campos y le pone formato tabular
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        rangoTabla1, Version:=6).CreatePivotTable TableDestination:= _
        celdaTablaDinamica1, TableName:="tablaDinamica1", DefaultVersion:=6
        
    Sheets("PT T4 After").Select
        With ActiveSheet.PivotTables("tablaDinamica1").PivotFields("PERNR")
        .Orientation = xlRowField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("tablaDinamica1").PivotFields("BUSNM")
        .Orientation = xlRowField
        .Position = 2
    End With
    With ActiveSheet.PivotTables("tablaDinamica1").PivotFields("NACHN")
        .Orientation = xlRowField
        .Position = 3
    End With
    With ActiveSheet.PivotTables("tablaDinamica1").PivotFields("VORNA")
        .Orientation = xlRowField
        .Position = 4
    End With
    With ActiveSheet.PivotTables("tablaDinamica1").PivotFields("WRKAR")
        .Orientation = xlRowField
        .Position = 5
    End With
    With ActiveSheet.PivotTables("tablaDinamica1").PivotFields("SLART")
        .Orientation = xlColumnField
        .Position = 1
    End With
    ActiveSheet.PivotTables("tablaDinamica1").AddDataField ActiveSheet.PivotTables( _
        "tablaDinamica1").PivotFields("BETRG"), "Cuenta de BETRG", xlCount
    With ActiveSheet.PivotTables("tablaDinamica1").PivotFields("Cuenta de BETRG")
        .Caption = "Suma de BETRG"
        .Function = xlSum
    End With
    ActiveSheet.PivotTables("tablaDinamica1").PivotFields("PERNR").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("tablaDinamica1").PivotFields("PERSONID").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("tablaDinamica1").PivotFields("FORML").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("tablaDinamica1").PivotFields("INDX1").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("tablaDinamica1").PivotFields("BUSNM").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("tablaDinamica1").PivotFields("WRKAR").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("tablaDinamica1").PivotFields("WTEXT").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("tablaDinamica1").PivotFields("KOSTL").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("tablaDinamica1").PivotFields("KTEXT").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("tablaDinamica1").PivotFields("NACHN").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("tablaDinamica1").PivotFields("VORNA").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("tablaDinamica1").PivotFields("SLART").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("tablaDinamica1").PivotFields("STEXT").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("tablaDinamica1").PivotFields("LGART").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("tablaDinamica1").PivotFields("LTEXT").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("tablaDinamica1").PivotFields("BETRG").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("tablaDinamica1").RowAxisLayout xlTabularRow
    ActiveSheet.PivotTables("tablaDinamica1").RepeatAllLabels xlRepeatLabels

wbauditoria.Save

'Copia y pega los datos en la hoja de valores
lastRow = Sheets("PT T4 After").Cells(Sheets("PT T4 After").Rows.Count, 1).End(xlUp).row
lastCol = Sheets("PT T4 After").Cells(2, Sheets("PT T4 After").Columns.Count).End(xlToLeft).Column
Sheets("PT T4 After").Range(Cells(2, 1), Cells(lastRow, lastCol)).Copy
Sheets("PT T4 After-values").Range("B1").PasteSpecial Paste:=xlPasteValues
Application.CutCopyMode = False
wbauditoria.Save

'Texto a numero
Sheets("PT T4 After-values").Activate
lastRow = Cells(Rows.Count, "B").End(xlUp).row
For i = 2 To lastRow
    If IsNumeric(Range("b" & i).Value) Then
        Range("B" & i).Value = Val(Range("B" & i).Value)
    End If
Next i

'Formula del concatenar el KEY NUMBER
Range("A1").Value = "KEY NUMBER"
lastRow = Cells(Rows.Count, "B").End(xlUp).row
Sheets("PT T4 After-values").Range("A2:A" & lastRow) = "=+CONCATENATE(RC[1],RC[2])"
Columns("A:AB").AutoFit

'Elimina los numeros de empleado que son 00000000
lastRow = Cells(Rows.Count, "A").End(xlUp).row
For i = lastRow To 2 Step -1
        If Cells(i, "B").Value = "0" Then
            Rows(i).Delete
        End If
Next i

lastRow = Cells(Rows.Count, "A").End(xlUp).row
For i = lastRow To 2 Step -1
        If Cells(i, "B").Value = "999999" Then
            Rows(i).Delete
        End If
Next i

wbauditoria.Save


'---------------------------------- ABRE ARCHIVO DE RL1 (BEFORE) ----------------------------------
MsgBox "Please select the RL1 (BEFORE) database file downloaded from SAP", vbInformation
ruta_Base = Application.GetOpenFilename("Archivos Excel (*.xls; *.xlsx), *.xls; *.xlsx")
Application.AskToUpdateLinks = False

' Verifica si el usuario seleccionó un archivo
If ruta_Base <> "Falso" Then
    'Abre el archivo base
    Set archivoBase = Workbooks.Open(Filename:=ruta_Base, UpdateLinks:=0)
    archivoBase.Activate
    Set hojaOrigen = archivoBase.Sheets(1)

    'Filtra solo los TTA
    Sheets(1).Activate
    lastRow = Cells(Rows.Count, "A").End(xlUp).row
    Rows("1:1").AutoFilter
    Rows("1:1").AutoFilter Field:=14, Criteria1:="=#TTA"
    Range("A1:P" & lastRow).SpecialCells(xlCellTypeVisible).Copy
    wbauditoria.Sheets("CCYR RL1 before").Cells(1, 1).PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False

    'Guarda el archivo nuevo y cierra el de la base
    wbauditoria.Save
    archivoBase.Close SaveChanges:=False

Else
End If



'Codigo para que cambie lo de las B
Sheets("CCYR RL1 before").Activate
Columns("M:M").Insert Shift:=xlToRight
Range("M1").Value = "BOX NAME"

' Recorrer cada fila desde la 2 hasta la última
lastRow = Cells(Rows.Count, "A").End(xlUp).row
For i = 2 To lastRow
    texto = Sheets("CCYR RL1 before").Cells(i, "L").Value
    Select Case texto
        Case "B01": Sheets("CCYR RL1 before").Cells(i, "M").Value = "Box A"
        Case "B02": Sheets("CCYR RL1 before").Cells(i, "M").Value = "Box B"
        Case "B02S": Sheets("CCYR RL1 before").Cells(i, "M").Value = "Box Bs"
        Case "B03": Sheets("CCYR RL1 before").Cells(i, "M").Value = "Box C"
        Case "B04": Sheets("CCYR RL1 before").Cells(i, "M").Value = "Box D"
        Case "B05": Sheets("CCYR RL1 before").Cells(i, "M").Value = "Box E"
        Case "B06": Sheets("CCYR RL1 before").Cells(i, "M").Value = "Box F"
        Case "B07": Sheets("CCYR RL1 before").Cells(i, "M").Value = "Box G"
        Case "B22": Sheets("CCYR RL1 before").Cells(i, "M").Value = "Box H"
        Case "B23": Sheets("CCYR RL1 before").Cells(i, "M").Value = "Box I"
        Case "B10": Sheets("CCYR RL1 before").Cells(i, "M").Value = "Box J"
        Case "B11": Sheets("CCYR RL1 before").Cells(i, "M").Value = "Box K"
        Case "B12": Sheets("CCYR RL1 before").Cells(i, "M").Value = "Box L"
        Case "B13": Sheets("CCYR RL1 before").Cells(i, "M").Value = "Box M"
        Case "B14": Sheets("CCYR RL1 before").Cells(i, "M").Value = "Box N"
        Case "B15": Sheets("CCYR RL1 before").Cells(i, "M").Value = "Box O"
        Case "B16": Sheets("CCYR RL1 before").Cells(i, "M").Value = "Box P"
        Case "B17": Sheets("CCYR RL1 before").Cells(i, "M").Value = "Box Q"
        Case "B18": Sheets("CCYR RL1 before").Cells(i, "M").Value = "Box R"
        Case "B19": Sheets("CCYR RL1 before").Cells(i, "M").Value = "Box S"
        Case "B20": Sheets("CCYR RL1 before").Cells(i, "M").Value = "Box T"
        Case "B21": Sheets("CCYR RL1 before").Cells(i, "M").Value = "Box U"
        Case "B08": Sheets("CCYR RL1 before").Cells(i, "M").Value = "Box V"
        Case "B09": Sheets("CCYR RL1 before").Cells(i, "M").Value = "Box W"
        Case Else: Sheets("CCYR RL1 before").Cells(i, "M").Value = texto ' Mantener el mismo valor
    End Select
Next i

'Poner todos los boxes para que el rango no cambie
lastRow = Sheets("CCYR RL1 before").Cells(Sheets("CCYR RL1 before").Rows.Count, 1).End(xlUp).row
lastCol = Sheets("CCYR RL1 before").Cells(2, Sheets("CCYR RL1 before").Columns.Count).End(xlToLeft).Column

ThisWorkbook.Sheets("Anexxes").Range("G2:G23").Copy
wbauditoria.Sheets("CCYR RL1 before").Cells(lastRow + 1, "M").PasteSpecial Paste:=xlPasteValues
Application.CutCopyMode = False
ThisWorkbook.Sheets("Anexxes").Range("I2:I23").Copy
wbauditoria.Sheets("CCYR RL1 before").Cells(lastRow + 1, "A").PasteSpecial Paste:=xlPasteValues
Application.CutCopyMode = False

'Crea tabla dinamica
Sheets("CCYR RL1 before").Activate

    ult_Tabla = Cells(Rows.Count, "A").End(xlUp).row
    Set rangoTabla1 = Sheets("CCYR RL1 before").Range("A1:Q" & ult_Tabla)
    ActiveSheet.ListObjects.Add(xlSrcRange, rangoTabla1, , xlYes).Name = "Tabla1"
    
    'Crear tabla dinamica
    Set celdaTablaDinamica1 = Sheets("PT RL1 Before").Range("A1")
    
    'Activa campos y le pone formato tabular
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        rangoTabla1, Version:=6).CreatePivotTable TableDestination:= _
        celdaTablaDinamica1, TableName:="tablaDinamica1", DefaultVersion:=6
        
    Sheets("PT RL1 Before").Select
    With ActiveSheet.PivotTables("tablaDinamica1").PivotFields("PERNR")
        .Orientation = xlRowField
        .Position = 1
    End With
    ActiveWindow.SmallScroll Down:=6
    With ActiveSheet.PivotTables("tablaDinamica1").PivotFields("BUSNM")
        .Orientation = xlRowField
        .Position = 2
    End With
    With ActiveSheet.PivotTables("tablaDinamica1").PivotFields("NACHN")
        .Orientation = xlRowField
        .Position = 3
    End With
    With ActiveSheet.PivotTables("tablaDinamica1").PivotFields("VORNA")
        .Orientation = xlRowField
        .Position = 4
    End With
    With ActiveSheet.PivotTables("tablaDinamica1").PivotFields("WRKAR")
        .Orientation = xlRowField
        .Position = 5
    End With
    With ActiveSheet.PivotTables("tablaDinamica1").PivotFields("BOX NAME")
        .Orientation = xlColumnField
        .Position = 1
    End With
    ActiveSheet.PivotTables("tablaDinamica1").AddDataField ActiveSheet.PivotTables( _
        "tablaDinamica1").PivotFields("BETRG"), "Cuenta de BETRG", xlCount
    With ActiveSheet.PivotTables("tablaDinamica1").PivotFields("Cuenta de BETRG")
        .Caption = "Suma de BETRG"
        .Function = xlSum
    End With
    ActiveSheet.PivotTables("tablaDinamica1").PivotFields("PERNR").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("tablaDinamica1").PivotFields("PERSONID").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("tablaDinamica1").PivotFields("FORML").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("tablaDinamica1").PivotFields("INDX1").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("tablaDinamica1").PivotFields("BUSNM").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("tablaDinamica1").PivotFields("WRKAR").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("tablaDinamica1").PivotFields("WTEXT").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("tablaDinamica1").PivotFields("KOSTL").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("tablaDinamica1").PivotFields("KTEXT").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("tablaDinamica1").PivotFields("NACHN").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("tablaDinamica1").PivotFields("VORNA").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("tablaDinamica1").PivotFields("SLART").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("tablaDinamica1").PivotFields("BOX NAME").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("tablaDinamica1").PivotFields("STEXT").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("tablaDinamica1").PivotFields("LGART").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("tablaDinamica1").PivotFields("LTEXT").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("tablaDinamica1").PivotFields("BETRG").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("tablaDinamica1").RowAxisLayout xlTabularRow
    ActiveSheet.PivotTables("tablaDinamica1").RepeatAllLabels xlRepeatLabels
    

wbauditoria.Save

'Copia y pega los datos en la hoja de valores
lastRow = Sheets("PT RL1 Before").Cells(Sheets("PT RL1 Before").Rows.Count, 1).End(xlUp).row
lastCol = Sheets("PT RL1 Before").Cells(2, Sheets("PT RL1 Before").Columns.Count).End(xlToLeft).Column
Sheets("PT RL1 Before").Range(Cells(2, 1), Cells(lastRow, lastCol)).Copy
Sheets("PT RL1 Before-values").Range("B1").PasteSpecial Paste:=xlPasteValues
Application.CutCopyMode = False
wbauditoria.Save

'Texto a numero
Sheets("PT RL1 Before-values").Activate
lastRow = Cells(Rows.Count, "B").End(xlUp).row
For i = 2 To lastRow
    If IsNumeric(Range("b" & i).Value) Then
        Range("B" & i).Value = Val(Range("B" & i).Value)
    End If
Next i

'Formula del concatenar el KEY NUMBER
Range("A1").Value = "KEY NUMBER"
lastRow = Cells(Rows.Count, "B").End(xlUp).row
Sheets("PT RL1 Before-values").Range("A2:A" & lastRow) = "=+CONCATENATE(RC[1],RC[2])"
Columns("A:AC").AutoFit

'Elimina los numeros de empleado que son 00000000
lastRow = Cells(Rows.Count, "A").End(xlUp).row
For i = lastRow To 2 Step -1
        If Cells(i, "B").Value = "0" Then
            Rows(i).Delete
        End If
Next i

lastRow = Cells(Rows.Count, "A").End(xlUp).row
For i = lastRow To 2 Step -1
        If Cells(i, "B").Value = "999999" Then
            Rows(i).Delete
        End If
Next i

wbauditoria.Save


'---------------------------------- ABRE ARCHIVO DE RL1 (AFTER) ----------------------------------
MsgBox "Please select the RL1 (AFTER) database file downloaded from SAP", vbInformation
ruta_Base = Application.GetOpenFilename("Archivos Excel (*.xls; *.xlsx), *.xls; *.xlsx")
Application.AskToUpdateLinks = False

' Verifica si el usuario seleccionó un archivo
If ruta_Base <> "Falso" Then
    'Abre el archivo base
    Set archivoBase = Workbooks.Open(Filename:=ruta_Base, UpdateLinks:=0)
    archivoBase.Activate
    Set hojaOrigen = archivoBase.Sheets(1)

    'Filtra solo los TTA
    Sheets(1).Activate
    lastRow = Cells(Rows.Count, "A").End(xlUp).row
    Rows("1:1").AutoFilter
    Rows("1:1").AutoFilter Field:=14, Criteria1:="=#TTA"
    Range("A1:P" & lastRow).SpecialCells(xlCellTypeVisible).Copy
    wbauditoria.Sheets("CCYR RL1 After").Cells(1, 1).PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False

    'Guarda el archivo nuevo y cierra el de la base
    wbauditoria.Save
    archivoBase.Close SaveChanges:=False

Else
End If

'Codigo para que cambie lo de las B
Sheets("CCYR RL1 After").Activate
Columns("M:M").Insert Shift:=xlToRight
Range("M1").Value = "BOX NAME"

' Recorrer cada fila desde la 2 hasta la última
lastRow = Cells(Rows.Count, "A").End(xlUp).row
For i = 2 To lastRow
    texto = Sheets("CCYR RL1 After").Cells(i, "L").Value
    Select Case texto
        Case "B01": Sheets("CCYR RL1 After").Cells(i, "M").Value = "Box A"
        Case "B02": Sheets("CCYR RL1 After").Cells(i, "M").Value = "Box B"
        Case "B02S": Sheets("CCYR RL1 After").Cells(i, "M").Value = "Box Bs"
        Case "B03": Sheets("CCYR RL1 After").Cells(i, "M").Value = "Box C"
        Case "B04": Sheets("CCYR RL1 After").Cells(i, "M").Value = "Box D"
        Case "B05": Sheets("CCYR RL1 After").Cells(i, "M").Value = "Box E"
        Case "B06": Sheets("CCYR RL1 After").Cells(i, "M").Value = "Box F"
        Case "B07": Sheets("CCYR RL1 After").Cells(i, "M").Value = "Box G"
        Case "B22": Sheets("CCYR RL1 After").Cells(i, "M").Value = "Box H"
        Case "B23": Sheets("CCYR RL1 After").Cells(i, "M").Value = "Box I"
        Case "B10": Sheets("CCYR RL1 After").Cells(i, "M").Value = "Box J"
        Case "B11": Sheets("CCYR RL1 After").Cells(i, "M").Value = "Box K"
        Case "B12": Sheets("CCYR RL1 After").Cells(i, "M").Value = "Box L"
        Case "B13": Sheets("CCYR RL1 After").Cells(i, "M").Value = "Box M"
        Case "B14": Sheets("CCYR RL1 After").Cells(i, "M").Value = "Box N"
        Case "B15": Sheets("CCYR RL1 After").Cells(i, "M").Value = "Box O"
        Case "B16": Sheets("CCYR RL1 After").Cells(i, "M").Value = "Box P"
        Case "B17": Sheets("CCYR RL1 After").Cells(i, "M").Value = "Box Q"
        Case "B18": Sheets("CCYR RL1 After").Cells(i, "M").Value = "Box R"
        Case "B19": Sheets("CCYR RL1 After").Cells(i, "M").Value = "Box S"
        Case "B20": Sheets("CCYR RL1 After").Cells(i, "M").Value = "Box T"
        Case "B21": Sheets("CCYR RL1 After").Cells(i, "M").Value = "Box U"
        Case "B08": Sheets("CCYR RL1 After").Cells(i, "M").Value = "Box V"
        Case "B09": Sheets("CCYR RL1 After").Cells(i, "M").Value = "Box W"
        Case Else: Sheets("CCYR RL1 After").Cells(i, "M").Value = texto ' Mantener el mismo valor
    End Select
Next i

'Poner todos los boxes para que el rango no cambie
lastRow = Sheets("CCYR RL1 After").Cells(Sheets("CCYR RL1 After").Rows.Count, 1).End(xlUp).row
lastCol = Sheets("CCYR RL1 After").Cells(2, Sheets("CCYR RL1 After").Columns.Count).End(xlToLeft).Column

ThisWorkbook.Sheets("Anexxes").Range("G2:G23").Copy
wbauditoria.Sheets("CCYR RL1 After").Cells(lastRow + 1, "M").PasteSpecial Paste:=xlPasteValues
Application.CutCopyMode = False
ThisWorkbook.Sheets("Anexxes").Range("I2:I23").Copy
wbauditoria.Sheets("CCYR RL1 After").Cells(lastRow + 1, "A").PasteSpecial Paste:=xlPasteValues
Application.CutCopyMode = False

'Crea tabla dinamica
Sheets("CCYR RL1 After").Activate

    ult_Tabla = Cells(Rows.Count, "A").End(xlUp).row
    Set rangoTabla1 = Sheets("CCYR RL1 After").Range("A1:Q" & ult_Tabla)
    ActiveSheet.ListObjects.Add(xlSrcRange, rangoTabla1, , xlYes).Name = "Tabla1"
    
    'Crear tabla dinamica
    Set celdaTablaDinamica1 = Sheets("PT RL1 After").Range("A1")
    
    'Activa campos y le pone formato tabular
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        rangoTabla1, Version:=6).CreatePivotTable TableDestination:= _
        celdaTablaDinamica1, TableName:="tablaDinamica1", DefaultVersion:=6
        
    Sheets("PT RL1 After").Select
   With ActiveSheet.PivotTables("tablaDinamica1").PivotFields("PERNR")
        .Orientation = xlRowField
        .Position = 1
    End With
    ActiveWindow.SmallScroll Down:=6
    With ActiveSheet.PivotTables("tablaDinamica1").PivotFields("BUSNM")
        .Orientation = xlRowField
        .Position = 2
    End With
    With ActiveSheet.PivotTables("tablaDinamica1").PivotFields("NACHN")
        .Orientation = xlRowField
        .Position = 3
    End With
    With ActiveSheet.PivotTables("tablaDinamica1").PivotFields("VORNA")
        .Orientation = xlRowField
        .Position = 4
    End With
    With ActiveSheet.PivotTables("tablaDinamica1").PivotFields("WRKAR")
        .Orientation = xlRowField
        .Position = 5
    End With
    With ActiveSheet.PivotTables("tablaDinamica1").PivotFields("BOX NAME")
        .Orientation = xlColumnField
        .Position = 1
    End With
    ActiveSheet.PivotTables("tablaDinamica1").AddDataField ActiveSheet.PivotTables( _
        "tablaDinamica1").PivotFields("BETRG"), "Cuenta de BETRG", xlCount
    With ActiveSheet.PivotTables("tablaDinamica1").PivotFields("Cuenta de BETRG")
        .Caption = "Suma de BETRG"
        .Function = xlSum
    End With
    ActiveSheet.PivotTables("tablaDinamica1").PivotFields("PERNR").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("tablaDinamica1").PivotFields("PERSONID").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("tablaDinamica1").PivotFields("FORML").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("tablaDinamica1").PivotFields("INDX1").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("tablaDinamica1").PivotFields("BUSNM").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("tablaDinamica1").PivotFields("WRKAR").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("tablaDinamica1").PivotFields("WTEXT").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("tablaDinamica1").PivotFields("KOSTL").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("tablaDinamica1").PivotFields("KTEXT").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("tablaDinamica1").PivotFields("NACHN").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("tablaDinamica1").PivotFields("VORNA").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("tablaDinamica1").PivotFields("SLART").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("tablaDinamica1").PivotFields("BOX NAME").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("tablaDinamica1").PivotFields("STEXT").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("tablaDinamica1").PivotFields("LGART").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("tablaDinamica1").PivotFields("LTEXT").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("tablaDinamica1").PivotFields("BETRG").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("tablaDinamica1").RowAxisLayout xlTabularRow
    ActiveSheet.PivotTables("tablaDinamica1").RepeatAllLabels xlRepeatLabels
    

wbauditoria.Save

'Copia y pega los datos en la hoja de valores
lastRow = Sheets("PT RL1 After").Cells(Sheets("PT RL1 After").Rows.Count, 1).End(xlUp).row
lastCol = Sheets("PT RL1 After").Cells(2, Sheets("PT RL1 After").Columns.Count).End(xlToLeft).Column
Sheets("PT RL1 After").Range(Cells(2, 1), Cells(lastRow, lastCol)).Copy
Sheets("PT RL1 After-values").Range("B1").PasteSpecial Paste:=xlPasteValues
Application.CutCopyMode = False
wbauditoria.Save

'Texto a numero
Sheets("PT RL1 After-values").Activate
lastRow = Cells(Rows.Count, "B").End(xlUp).row
For i = 2 To lastRow
    If IsNumeric(Range("b" & i).Value) Then
        Range("B" & i).Value = Val(Range("B" & i).Value)
    End If
Next i

'Formula del concatenar el KEY NUMBER
Range("A1").Value = "KEY NUMBER"
lastRow = Cells(Rows.Count, "B").End(xlUp).row
Sheets("PT RL1 After-values").Range("A2:A" & lastRow) = "=+CONCATENATE(RC[1],RC[2])"
Columns("A:AC").AutoFit

'Elimina los numeros de empleado que son 00000000
lastRow = Cells(Rows.Count, "A").End(xlUp).row
For i = lastRow To 2 Step -1
        If Cells(i, "B").Value = "0" Then
            Rows(i).Delete
        End If
Next i

lastRow = Cells(Rows.Count, "A").End(xlUp).row
For i = lastRow To 2 Step -1
        If Cells(i, "B").Value = "999999" Then
            Rows(i).Delete
        End If
Next i

'Oculta hojas que ya no son necesarias
wbauditoria.Sheets("CCYR T4 before").Visible = xlSheetHidden
wbauditoria.Sheets("CCYR T4 After").Visible = xlSheetHidden
wbauditoria.Sheets("CCYR RL1 before").Visible = xlSheetHidden
wbauditoria.Sheets("CCYR RL1 After").Visible = xlSheetHidden
wbauditoria.Save

'---------------------------------- ELIMINAR LAS COLUMNAS QUE NO SON NECESARIAS ----------------------------------
'La eliminacion depende de cual auditoria se esta ejecutando, solo va a dejar las que necesita

'Hoja PT T4 Before-values
Sheets("PT T4 Before-values").Activate
Select Case auditoria
 
     Case "Quebec Health Tax entries"
        Union(Columns("G:I"), Columns("L:P"), Columns("R:AF")).Delete
         
     Case "Auto Taxable benefit entries"
        Union(Columns("L:P"), Columns("T:Y"), Columns("Z:AF")).Delete
         
     Case "Taxable Benefits entries (PSA - Benefits)"
        Union(Columns("L:P"), Columns("T:Z"), Columns("AB:AF")).Delete
         
     Case "Pension Adjustment entries"
        Union(Columns("G:U"), Columns("W:AF")).Delete
        
     Case "CPP QPP Boxes validation"
        Union(Columns("G:G"), Columns("L:P"), Columns("T:AF")).Delete
         
     Case "EI boxes validation"
        Union(Columns("G:I"), Columns("K:K"), Columns("N:O"), Columns("Q:AF")).Delete
         
 End Select

'PT T4 After-values
Sheets("PT T4 After-values").Activate
Select Case auditoria
 
     Case "Quebec Health Tax entries"
        Union(Columns("G:I"), Columns("L:P"), Columns("R:AF")).Delete
         
     Case "Auto Taxable benefit entries"
        Union(Columns("L:P"), Columns("T:Y"), Columns("Z:AF")).Delete
         
     Case "Taxable Benefits entries (PSA - Benefits)"
        Union(Columns("L:P"), Columns("T:Z"), Columns("AB:AF")).Delete
         
     Case "Pension Adjustment entries"
        Union(Columns("G:U"), Columns("W:AF")).Delete
        
     Case "CPP QPP Boxes validation"
        Union(Columns("G:G"), Columns("L:P"), Columns("T:AF")).Delete
         
     Case "EI boxes validation"
        Union(Columns("G:I"), Columns("K:K"), Columns("N:O"), Columns("Q:AF")).Delete
         
 End Select
 
'Hoja PT RL1 Before-values
Sheets("PT RL1 Before-values").Activate
Select Case auditoria
 
     Case "Quebec Health Tax entries"
        Union(Columns("K:N"), Columns("P:Q"), Columns("S:AC")).Delete
         
     Case "Auto Taxable benefit entries"
        Union(Columns("K:N"), Columns("P:U"), Columns("W:AC")).Delete
         
     Case "Taxable Benefits entries (PSA - Benefits)"
        Union(Columns("K:N"), Columns("P:R"), Columns("T:AC")).Delete
         
     Case "Pension Adjustment entries"
        Application.DisplayAlerts = False
        Sheets(1).Activate
        wbauditoria.Sheets("CCYR RL1 before").Delete
        wbauditoria.Sheets("PT RL1 Before").Delete
        wbauditoria.Sheets("PT RL1 Before-values").Delete
        wbauditoria.Sheets("Validation RL1").Delete
        Application.DisplayAlerts = True
        
     Case "CPP QPP Boxes validation"
        Union(Columns("G:G"), Columns("K:N"), Columns("P:R"), Columns("T:X"), Columns("AB:AC")).Delete
         
     Case "EI boxes validation"
        Union(Columns("G:J"), Columns("L:AC")).Delete
         
 End Select
 
'Hoja PT RL1 After-values
Sheets("PT RL1 After-values").Activate
Select Case auditoria
 
     Case "Quebec Health Tax entries"
        Union(Columns("K:N"), Columns("P:Q"), Columns("S:AC")).Delete
         
     Case "Auto Taxable benefit entries"
        Union(Columns("K:N"), Columns("P:U"), Columns("W:AC")).Delete
         
     Case "Taxable Benefits entries (PSA - Benefits)"
        Union(Columns("K:N"), Columns("P:R"), Columns("T:AC")).Delete
         
     Case "Pension Adjustment entries"
        Application.DisplayAlerts = False
        Sheets(1).Activate
        wbauditoria.Sheets("CCYR RL1 After").Delete
        wbauditoria.Sheets("PT RL1 After").Delete
        wbauditoria.Sheets("PT RL1 After-values").Delete
        Application.DisplayAlerts = True
        
     Case "CPP QPP Boxes validation"
        Union(Columns("G:G"), Columns("K:N"), Columns("P:R"), Columns("T:X"), Columns("AB:AC")).Delete
         
     Case "EI boxes validation"
        Union(Columns("G:J"), Columns("L:AC")).Delete
         
 End Select
 wbauditoria.Save


'---------------------------------- CREA PLANTILLA DE LAS HOJAS DE LAS VALIDACIONES ----------------------------------
'Va a cada modulo y crea la auditoria

Select Case auditoria
 
     Case "Quebec Health Tax entries"
        Call Quebec_HT_Entry
       
     Case "Auto Taxable benefit entries"
        Call Auto_TB_Entries
         
     Case "Taxable Benefits entries (PSA - Benefits)"
        Call Taxable_B_Entries
         
     Case "Pension Adjustment entries"
        Call Pension_Entries
        
     Case "CPP QPP Boxes validation"
        Call CPP_QPP_validation
         
     Case "EI boxes validation"
        Call EI_boxes_validation
         
 End Select


End Sub

Sub Quebec_HT_Entry()

     '------ Crea la hoja de validacion del RL1
        
        'Copia los datos de la hoja de before del RL1
        Sheets("Validation RL1").Activate
        lastRow = Sheets("PT RL1 Before-values").Cells(Sheets("PT RL1 Before-values").Rows.Count, 3).End(xlUp).row
        lastCol = Sheets("PT RL1 Before-values").Cells(1, Sheets("PT RL1 Before-values").Columns.Count).End(xlToLeft).Column
        Set rng = wbauditoria.Sheets("PT RL1 Before-values").Range(wbauditoria.Sheets("PT RL1 Before-values").Cells(1, 1), wbauditoria.Sheets("PT RL1 Before-values").Cells(lastRow, lastCol))
        wbauditoria.Sheets("Validation RL1").Range("A2").Resize(rng.Rows.Count, rng.Columns.Count).Value = rng.Value
        
        'Calcula ultima fila y columna de la hoja de validacion
        Columns("A:A").Insert Shift:=xlToRight
        lastRow = Sheets("Validation RL1").Cells(Sheets("Validation RL1").Rows.Count, 3).End(xlUp).row
        lastCol = Sheets("Validation RL1").Cells(2, Sheets("Validation RL1").Columns.Count).End(xlToLeft).Column
        
        'Coloca los titulos en la fila 2
        ThisWorkbook.Sheets("Anexxes").Range("L2:AS2").Copy
        wbauditoria.Sheets("Validation RL1").Range("A2").PasteSpecial Paste:=xlPasteValues
        Application.CutCopyMode = False
        Columns("A:AS").AutoFit
        
        'Coloca titulos en celdas combinadas
        Range("H1:M1").Merge
        Range("H1").Value = "BEFORE ENTRIES"
        Range("N1:S1").Merge
        Range("N1").Value = "AFTER ENTRIES"
        Range("T1:Y1").Merge
        Range("T1").Value = "DIFFERENCES"
        Range("AA1:AH1").Merge
        Range("AA1").Value = "VALIDATION"
        
        With Range("A2:G2")
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
        
        With Range("H1:M2")
            .Font.Bold = True
            .Interior.Color = RGB(209, 242, 235)
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
        
        With Range("N1:S2")
            .Font.Bold = True
            .Interior.Color = RGB(250, 229, 211)
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
        
        With Range("T1:Y2")
            .Font.Bold = True
            .Interior.Color = RGB(250, 219, 216)
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
        
        With Range("Z1:Z2")
            .Font.Bold = True
            .Interior.Color = RGB(212, 230, 241)
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
        
        With Range("AA1:AH2")
            .Font.Bold = True
            .Interior.Color = RGB(232, 218, 239)
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
        
        Columns("A:AH").AutoFit
        
        'Formula de la columa A
        lastRow = Sheets("Validation RL1").Cells(Sheets("Validation RL1").Rows.Count, 3).End(xlUp).row
        lastCol = Sheets("Validation RL1").Cells(2, Sheets("Validation RL1").Columns.Count).End(xlToLeft).Column
        Sheets("Validation RL1").Range("A3:A" & lastRow) = "=COUNTIF('PT RL1 After-values'!C[1],'Validation RL1'!RC[2])"
        
        'Coloca el resto de las formulas
        lastRow = Sheets("Validation RL1").Cells(Sheets("Validation RL1").Rows.Count, 3).End(xlUp).row
        Sheets("Validation RL1").Range("N3:N" & lastRow) = "=VLOOKUP(RC[-12],'PT RL1 After-values'!C1:C12,7,0)"
        Sheets("Validation RL1").Range("O3:O" & lastRow) = "=VLOOKUP(RC[-13],'PT RL1 After-values'!C1:C12,8,0)"
        Sheets("Validation RL1").Range("P3:P" & lastRow) = "=VLOOKUP(RC[-14],'PT RL1 After-values'!C1:C12,9,0)"
        Sheets("Validation RL1").Range("Q3:Q" & lastRow) = "=VLOOKUP(RC[-15],'PT RL1 After-values'!C1:C12,10,0)"
        Sheets("Validation RL1").Range("R3:R" & lastRow) = "=VLOOKUP(RC[-16],'PT RL1 After-values'!C1:C12,11,0)"
        Sheets("Validation RL1").Range("S3:S" & lastRow) = "=VLOOKUP(RC[-17],'PT RL1 After-values'!C1:C12,12,0)"
        Sheets("Validation RL1").Range("T3:T" & lastRow) = "=RC[-6]-RC[-12]"
        Sheets("Validation RL1").Range("U3:U" & lastRow) = "=RC[-6]-RC[-12]"
        Sheets("Validation RL1").Range("V3:V" & lastRow) = "=RC[-6]-RC[-12]"
        Sheets("Validation RL1").Range("W3:W" & lastRow) = "=RC[-6]-RC[-12]"
        Sheets("Validation RL1").Range("X3:X" & lastRow) = "=RC[-6]-RC[-12]"
        Sheets("Validation RL1").Range("Y3:Y" & lastRow) = "=RC[-6]-RC[-12]"
        Sheets("Validation RL1").Range("Z3:Z" & lastRow) = "=VLOOKUP(RC[-23],'BoxJ entry'!C1:C4,4,0)"
        Sheets("Validation RL1").Range("AA3:AA" & lastRow) = "=ROUND(IF(RC[-7]=0,0,RC[-7]-RC[-1]),2)"
        Sheets("Validation RL1").Range("AB3:AB" & lastRow) = "=ROUND(IF(RC[-7]=0,0,RC[-7]-RC[-2]),2)"
        Sheets("Validation RL1").Range("AC3:AC" & lastRow) = "=ROUND(IF(RC[-7]=0,0,RC[-7]-RC[-3]),2)"
        Sheets("Validation RL1").Range("AD3:AD" & lastRow) = "=IF((RC[-9]+RC[-8])=RC[-7],0,RC[-7])"
        Sheets("Validation RL1").Range("AE3:AE" & lastRow) = "=ROUND(IF(RC[-7]=0,0,RC[-7]-RC[-5]),2)"
        Sheets("Validation RL1").Range("AF3:AF" & lastRow) = "=ROUND(IF(RC[-7]=0,0,RC[-7]-RC[-6]),2)"
        Sheets("Validation RL1").Range("AG3:AG" & lastRow) = "=VLOOKUP(RC[-30],'BoxJ entry'!C1:C2,2,0)"
        Sheets("Validation RL1").Range("AH3:AH" & lastRow) = "=RC[-1]-RC[-15]"
        Columns("A:AH").AutoFit

        
        '------ Crea la hoja de validacion del T4
        
        'Copia los datos de la hoja de before del T4
        Sheets("Validation T4").Activate
        lastRow = Sheets("PT T4 Before-values").Cells(Sheets("PT T4 Before-values").Rows.Count, 1).End(xlUp).row
        lastCol = Sheets("PT T4 Before-values").Cells(1, Sheets("PT T4 Before-values").Columns.Count).End(xlToLeft).Column
        Set rng = wbauditoria.Sheets("PT T4 Before-values").Range(wbauditoria.Sheets("PT T4 Before-values").Cells(1, 1), wbauditoria.Sheets("PT T4 Before-values").Cells(lastRow, lastCol))
        wbauditoria.Sheets("Validation T4").Range("A2").Resize(rng.Rows.Count, rng.Columns.Count).Value = rng.Value
        
        'Calcula ultima fila y columna de la hoja de validacion
        Columns("A:A").Insert Shift:=xlToRight
        lastRow = Sheets("Validation T4").Cells(Sheets("Validation T4").Rows.Count, 3).End(xlUp).row
        lastCol = Sheets("Validation T4").Cells(2, Sheets("Validation T4").Columns.Count).End(xlToLeft).Column
        
        'Coloca los titulos en la fila 2
        ThisWorkbook.Sheets("Anexxes").Range("L1:AL1").Copy
        wbauditoria.Sheets("Validation T4").Range("A2").PasteSpecial Paste:=xlPasteValues
        Application.CutCopyMode = False
        Columns("A:AC").AutoFit
        
        'Coloca titulos en celdas combinadas
        Range("H1:J1").Merge
        Range("H1").Value = "BEFORE ENTRIES"
        Range("K1:M1").Merge
        Range("K1").Value = "AFTER ENTRIES"
        Range("N1:P1").Merge
        Range("N1").Value = "DIFFERENCES"
        Range("R1:T1").Merge
        Range("R1").Value = "VALIDATION"
        Range("U1:AA1").Merge
        Range("U1").Value = "RL1 VALIDATION"
        
        With Range("A2:G2")
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
        
        With Range("H1:J2")
            .Font.Bold = True
            .Interior.Color = RGB(209, 242, 235)
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
        
        With Range("K1:M2")
            .Font.Bold = True
            .Interior.Color = RGB(250, 229, 211)
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
        
        With Range("N1:P2")
            .Font.Bold = True
            .Interior.Color = RGB(250, 219, 216)
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
        
        With Range("Q1:Q2")
            .Font.Bold = True
            .Interior.Color = RGB(212, 230, 241)
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
        
        With Range("R1:T2")
            .Font.Bold = True
            .Interior.Color = RGB(232, 218, 239)
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
        
        With Range("U1:AA2")
            .Font.Bold = True
            .Interior.Color = RGB(250, 229, 211)
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
        Columns("A:AC").AutoFit
        
        'Formula de la columa A
        lastRow = Sheets("Validation T4").Cells(Sheets("Validation T4").Rows.Count, 3).End(xlUp).row
        lastCol = Sheets("Validation T4").Cells(2, Sheets("Validation T4").Columns.Count).End(xlToLeft).Column
        Sheets("Validation T4").Range("A3:A" & lastRow) = "=COUNTIF('PT T4 After-values'!C[1],'Validation T4'!RC[2])"
        
        'Coloca el resto de las formulas
        lastRow = Sheets("Validation T4").Cells(Sheets("Validation T4").Rows.Count, 3).End(xlUp).row
        Sheets("Validation T4").Range("K3:K" & lastRow) = "=VLOOKUP(RC[-9],'PT T4 After-values'!C1:C9,7,0)"
        Sheets("Validation T4").Range("L3:L" & lastRow) = "=VLOOKUP(RC[-10],'PT T4 After-values'!C1:C9,8,0)"
        Sheets("Validation T4").Range("M3:M" & lastRow) = "=VLOOKUP(RC[-11],'PT T4 After-values'!C1:C9,9,0)"
        Sheets("Validation T4").Range("N3:N" & lastRow) = "=RC[-3]-RC[-6]"
        Sheets("Validation T4").Range("O3:O" & lastRow) = "=RC[-3]-RC[-6]"
        Sheets("Validation T4").Range("P3:P" & lastRow) = "=RC[-3]-RC[-6]"
        Sheets("Validation T4").Range("Q3:Q" & lastRow) = "=VLOOKUP(RC[-14],'BoxJ entry'!C1:C4,4,0)"
        Sheets("Validation T4").Range("R3:R" & lastRow) = "=ROUND(IF(RC[-4]=0,0,RC[-4]-RC17),2)"
        Sheets("Validation T4").Range("S3:S" & lastRow) = "=ROUND(IF(RC[-4]=0,0,RC[-4]-RC17),2)"
        Sheets("Validation T4").Range("T3:T" & lastRow) = "=ROUND(IF(RC[-4]=0,0,RC[-4]-RC17),2)"
        Sheets("Validation T4").Range("U3:U" & lastRow) = "=VLOOKUP(RC[-18],'Validation RL1'!C[-18]:C[-2],13,0)"
        Sheets("Validation T4").Range("V3:V" & lastRow) = "=VLOOKUP(RC[-19],'Validation RL1'!C3:C19,14,0)"
        Sheets("Validation T4").Range("W3:W" & lastRow) = "=RC[-2]-RC[-12]"
        Sheets("Validation T4").Range("X3:X" & lastRow) = "=RC[-2]-RC[-12]"
        Sheets("Validation T4").Range("Y3:Y" & lastRow) = "=VLOOKUP(RC[-22],'Validation RL1'!C3:C19,15,0)"
        Sheets("Validation T4").Range("Z3:Z" & lastRow) = "=(RC[-5]+RC[-4])-RC[-1]"
        Sheets("Validation T4").Range("AA3:AA" & lastRow) = "=RC[-2]-RC[-6]"
        Columns("A:AC").AutoFit
        
        'Ordena hojas
        Sheets("BoxJ entry").Move Before:=Sheets(1)
        Sheets("Validation RL1").Move After:=Sheets("BoxJ entry")
        Sheets("Validation T4").Move After:=Sheets("Validation RL1")
        Sheets("Validation RL1").Tab.Color = RGB(250, 219, 216)
        Sheets("Validation T4").Tab.Color = RGB(250, 229, 211)
        
        'Organiza la hoja de BoxJ entry
        Sheets("BoxJ entry").Activate
        Range("A1").Value = "SAP ID"
        Range("B1").Value = "CORRECT BOX J"
        Range("C1").Value = "BOX J TOTAL"
        Range("D1").Value = "SAP DIFFRENCE (ENTRY)"
         With Range("A1:D1")
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
            .Interior.Color = RGB(235, 237, 239)
        End With
        Columns("A:D").AutoFit
        
        'Finaliza
        wbauditoria.Save
        wbauditoria.Close
        MsgBox "The process has been completed. Please enter the document and complete the data in the 'BoxJ entry' sheet.", vbInformation

End Sub

Sub Auto_TB_Entries()

'------ Crea la hoja de validacion del RL1
        
        'Copia los datos de la hoja de before del RL1
        Sheets("Validation RL1").Activate
        lastRow = Sheets("PT RL1 Before-values").Cells(Sheets("PT RL1 Before-values").Rows.Count, 3).End(xlUp).row
        lastCol = Sheets("PT RL1 Before-values").Cells(1, Sheets("PT RL1 Before-values").Columns.Count).End(xlToLeft).Column
        Set rng = wbauditoria.Sheets("PT RL1 Before-values").Range(wbauditoria.Sheets("PT RL1 Before-values").Cells(1, 1), wbauditoria.Sheets("PT RL1 Before-values").Cells(lastRow, lastCol))
        wbauditoria.Sheets("Validation RL1").Range("A2").Resize(rng.Rows.Count, rng.Columns.Count).Value = rng.Value

        'Calcula ultima fila y columna de la hoja de validacion
        Columns("A:A").Insert Shift:=xlToRight
        lastRow = Sheets("Validation RL1").Cells(Sheets("Validation RL1").Rows.Count, 3).End(xlUp).row
        lastCol = Sheets("Validation RL1").Cells(2, Sheets("Validation RL1").Columns.Count).End(xlToLeft).Column

        'Coloca los titulos en la fila 2
        ThisWorkbook.Sheets("Anexxes").Range("L4:AQ4").Copy
        wbauditoria.Sheets("Validation RL1").Range("A2").PasteSpecial Paste:=xlPasteValues
        Application.CutCopyMode = False
        Columns("A:AQ").AutoFit

        'Coloca titulos en celdas combinadas
        Range("H1:M1").Merge
        Range("H1").Value = "BEFORE ENTRIES"
        Range("N1:S1").Merge
        Range("N1").Value = "AFTER ENTRIES"
        Range("T1:Y1").Merge
        Range("T1").Value = "DIFFERENCES"
'        Range("AA1:AF1").Merge
'        Range("AA1").Value = "VALIDATION"

        With Range("A2:G2")
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With

        With Range("H1:M2")
            .Font.Bold = True
            .Interior.Color = RGB(209, 242, 235)
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With

        With Range("N1:S2")
            .Font.Bold = True
            .Interior.Color = RGB(250, 229, 211)
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With

        With Range("T1:Y2")
            .Font.Bold = True
            .Interior.Color = RGB(250, 219, 216)
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With

        With Range("Z1:Z2")
            .Font.Bold = True
            .Interior.Color = RGB(212, 230, 241)
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With

'        With Range("AA1:AH2")
'            .Font.Bold = True
'            .Interior.Color = RGB(232, 218, 239)
'            .HorizontalAlignment = xlCenter
'            .VerticalAlignment = xlCenter
'        End With
'
'        Columns("A:AH").AutoFit
'
        'Formula de la columa A
        lastRow = Sheets("Validation RL1").Cells(Sheets("Validation RL1").Rows.Count, 3).End(xlUp).row
        lastCol = Sheets("Validation RL1").Cells(2, Sheets("Validation RL1").Columns.Count).End(xlToLeft).Column
        Sheets("Validation RL1").Range("A3:A" & lastRow) = "=COUNTIF('PT RL1 After-values'!C[1],'Validation RL1'!RC[2])"
'
'        'Coloca el resto de las formulas
'        lastRow = Sheets("Validation RL1").Cells(Sheets("Validation RL1").Rows.Count, 3).End(xlUp).row
'        Sheets("Validation RL1").Range("N3:N" & lastRow) = "=VLOOKUP(RC[-12],'PT RL1 After-values'!C1:C12,7,0)"
'        Sheets("Validation RL1").Range("O3:O" & lastRow) = "=VLOOKUP(RC[-13],'PT RL1 After-values'!C1:C12,8,0)"
'        Sheets("Validation RL1").Range("P3:P" & lastRow) = "=VLOOKUP(RC[-14],'PT RL1 After-values'!C1:C12,9,0)"
'        Sheets("Validation RL1").Range("Q3:Q" & lastRow) = "=VLOOKUP(RC[-15],'PT RL1 After-values'!C1:C12,10,0)"
'        Sheets("Validation RL1").Range("R3:R" & lastRow) = "=VLOOKUP(RC[-16],'PT RL1 After-values'!C1:C12,11,0)"
'        Sheets("Validation RL1").Range("S3:S" & lastRow) = "=VLOOKUP(RC[-17],'PT RL1 After-values'!C1:C12,12,0)"
'        Sheets("Validation RL1").Range("T3:T" & lastRow) = "=RC[-6]-RC[-12]"
'        Sheets("Validation RL1").Range("U3:U" & lastRow) = "=RC[-6]-RC[-12]"
'        Sheets("Validation RL1").Range("V3:V" & lastRow) = "=RC[-6]-RC[-12]"
'        Sheets("Validation RL1").Range("W3:W" & lastRow) = "=RC[-6]-RC[-12]"
'        Sheets("Validation RL1").Range("X3:X" & lastRow) = "=RC[-6]-RC[-12]"
'        Sheets("Validation RL1").Range("Y3:Y" & lastRow) = "=RC[-6]-RC[-12]"
'        Sheets("Validation RL1").Range("Z3:Z" & lastRow) = "=VLOOKUP(RC[-23],'BoxJ entry'!C1:C4,4,0)"
'        Sheets("Validation RL1").Range("AA3:AA" & lastRow) = "=ROUND(IF(RC[-7]=0,0,RC[-7]-RC[-1]),2)"
'        Sheets("Validation RL1").Range("AB3:AB" & lastRow) = "=ROUND(IF(RC[-7]=0,0,RC[-7]-RC[-2]),2)"
'        Sheets("Validation RL1").Range("AC3:AC" & lastRow) = "=ROUND(IF(RC[-7]=0,0,RC[-7]-RC[-3]),2)"
'        Sheets("Validation RL1").Range("AD3:AD" & lastRow) = "=IF((RC[-9]+RC[-8])=RC[-7],0,RC[-7])"
'        Sheets("Validation RL1").Range("AE3:AE" & lastRow) = "=ROUND(IF(RC[-7]=0,0,RC[-7]-RC[-5]),2)"
'        Sheets("Validation RL1").Range("AF3:AF" & lastRow) = "=ROUND(IF(RC[-7]=0,0,RC[-7]-RC[-6]),2)"
'        Sheets("Validation RL1").Range("AG3:AG" & lastRow) = "=VLOOKUP(RC[-30],'BoxJ entry'!C1:C2,2,0)"
'        Sheets("Validation RL1").Range("AH3:AH" & lastRow) = "=RC[-1]-RC[-15]"
'        Columns("A:AH").AutoFit


        '------ Crea la hoja de validacion del T4

        'Copia los datos de la hoja de before del T4
        Sheets("Validation T4").Activate
        lastRow = Sheets("PT T4 Before-values").Cells(Sheets("PT T4 Before-values").Rows.Count, 1).End(xlUp).row
        lastCol = Sheets("PT T4 Before-values").Cells(1, Sheets("PT T4 Before-values").Columns.Count).End(xlToLeft).Column
        Set rng = wbauditoria.Sheets("PT T4 Before-values").Range(wbauditoria.Sheets("PT T4 Before-values").Cells(1, 1), wbauditoria.Sheets("PT T4 Before-values").Cells(lastRow, lastCol))
        wbauditoria.Sheets("Validation T4").Range("A2").Resize(rng.Rows.Count, rng.Columns.Count).Value = rng.Value

        'Calcula ultima fila y columna de la hoja de validacion
        Columns("A:A").Insert Shift:=xlToRight
        lastRow = Sheets("Validation T4").Cells(Sheets("Validation T4").Rows.Count, 3).End(xlUp).row
        lastCol = Sheets("Validation T4").Cells(2, Sheets("Validation T4").Columns.Count).End(xlToLeft).Column

        'Coloca los titulos en la fila 2
        ThisWorkbook.Sheets("Anexxes").Range("L3:AU3").Copy
        wbauditoria.Sheets("Validation T4").Range("A2").PasteSpecial Paste:=xlPasteValues
        Application.CutCopyMode = False
        Columns("A:AU").AutoFit

        'Coloca titulos en celdas combinadas
        Range("H1:N1").Merge
        Range("H1").Value = "BEFORE ENTRIES"
        Range("O1:U1").Merge
        Range("O1").Value = "AFTER ENTRIES"
        Range("V1:AB1").Merge
        Range("V1").Value = "DIFFERENCES"
        Range("R1:T1").Merge
'        Range("R1").Value = "VALIDATION"
'        Range("U1:AA1").Merge
'        Range("U1").Value = "RL1 VALIDATION"

        With Range("A2:G2")
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With

        With Range("H1:N2")
            .Font.Bold = True
            .Interior.Color = RGB(209, 242, 235)
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With

        With Range("O1:U2")
            .Font.Bold = True
            .Interior.Color = RGB(250, 229, 211)
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With

        With Range("V1:AB2")
            .Font.Bold = True
            .Interior.Color = RGB(250, 219, 216)
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With

        With Range("AC1:AC2")
            .Font.Bold = True
            .Interior.Color = RGB(212, 230, 241)
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With

'        With Range("R1:T2")
'            .Font.Bold = True
'            .Interior.Color = RGB(232, 218, 239)
'            .HorizontalAlignment = xlCenter
'            .VerticalAlignment = xlCenter
'        End With
'
'        With Range("U1:AA2")
'            .Font.Bold = True
'            .Interior.Color = RGB(250, 229, 211)
'            .HorizontalAlignment = xlCenter
'            .VerticalAlignment = xlCenter
'        End With
        Columns("A:AC").AutoFit

        'Formula de la columa A
        lastRow = Sheets("Validation T4").Cells(Sheets("Validation T4").Rows.Count, 3).End(xlUp).row
        lastCol = Sheets("Validation T4").Cells(2, Sheets("Validation T4").Columns.Count).End(xlToLeft).Column
        Sheets("Validation T4").Range("A3:A" & lastRow) = "=COUNTIF('PT T4 After-values'!C[1],'Validation T4'!RC[2])"
'
'        'Coloca el resto de las formulas
'        lastRow = Sheets("Validation T4").Cells(Sheets("Validation T4").Rows.Count, 3).End(xlUp).row
'        Sheets("Validation T4").Range("K3:K" & lastRow) = "=VLOOKUP(RC[-9],'PT T4 After-values'!C1:C9,7,0)"
'        Sheets("Validation T4").Range("L3:L" & lastRow) = "=VLOOKUP(RC[-10],'PT T4 After-values'!C1:C9,8,0)"
'        Sheets("Validation T4").Range("M3:M" & lastRow) = "=VLOOKUP(RC[-11],'PT T4 After-values'!C1:C9,9,0)"
'        Sheets("Validation T4").Range("N3:N" & lastRow) = "=RC[-3]-RC[-6]"
'        Sheets("Validation T4").Range("O3:O" & lastRow) = "=RC[-3]-RC[-6]"
'        Sheets("Validation T4").Range("P3:P" & lastRow) = "=RC[-3]-RC[-6]"
'        Sheets("Validation T4").Range("Q3:Q" & lastRow) = "=VLOOKUP(RC[-14],'BoxJ entry'!C1:C4,4,0)"
'        Sheets("Validation T4").Range("R3:R" & lastRow) = "=ROUND(IF(RC[-4]=0,0,RC[-4]-RC17),2)"
'        Sheets("Validation T4").Range("S3:S" & lastRow) = "=ROUND(IF(RC[-4]=0,0,RC[-4]-RC17),2)"
'        Sheets("Validation T4").Range("T3:T" & lastRow) = "=ROUND(IF(RC[-4]=0,0,RC[-4]-RC17),2)"
'        Sheets("Validation T4").Range("U3:U" & lastRow) = "=VLOOKUP(RC[-18],'Validation RL1'!C[-18]:C[-2],13,0)"
'        Sheets("Validation T4").Range("V3:V" & lastRow) = "=VLOOKUP(RC[-19],'Validation RL1'!C3:C19,14,0)"
'        Sheets("Validation T4").Range("W3:W" & lastRow) = "=RC[-2]-RC[-12]"
'        Sheets("Validation T4").Range("X3:X" & lastRow) = "=RC[-2]-RC[-12]"
'        Sheets("Validation T4").Range("Y3:Y" & lastRow) = "=VLOOKUP(RC[-22],'Validation RL1'!C3:C19,15,0)"
'        Sheets("Validation T4").Range("Z3:Z" & lastRow) = "=(RC[-5]+RC[-4])-RC[-1]"
'        Sheets("Validation T4").Range("AA3:AA" & lastRow) = "=RC[-2]-RC[-6]"
'        Columns("A:AC").AutoFit

        'Ordena hojas
        Sheets("BoxJ entry").Move Before:=Sheets(1)
        Sheets("Validation RL1").Move After:=Sheets("BoxJ entry")
        Sheets("Validation T4").Move After:=Sheets("Validation RL1")
        Sheets("Validation RL1").Tab.Color = RGB(250, 219, 216)
        Sheets("Validation T4").Tab.Color = RGB(250, 229, 211)

        'Organiza la hoja de BoxJ entry
        Sheets("BoxJ entry").Activate
        Range("A1").Value = "SAP ID"
        Range("B1").Value = "CORRECT BOX J"
        Range("C1").Value = "BOX J TOTAL"
        Range("D1").Value = "SAP DIFFRENCE (ENTRY)"
         With Range("A1:D1")
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
            .Interior.Color = RGB(235, 237, 239)
        End With
        Columns("A:D").AutoFit

        'Finaliza
        wbauditoria.Save
        wbauditoria.Close
        MsgBox "The process has been completed. Please enter the document and complete the data in the 'BoxJ entry' sheet.", vbInformation

End Sub

Sub Taxable_B_Entries()

'------ Crea la hoja de validacion del RL1
        
        'Copia los datos de la hoja de before del RL1
        Sheets("Validation RL1").Activate
        lastRow = Sheets("PT RL1 Before-values").Cells(Sheets("PT RL1 Before-values").Rows.Count, 3).End(xlUp).row
        lastCol = Sheets("PT RL1 Before-values").Cells(1, Sheets("PT RL1 Before-values").Columns.Count).End(xlToLeft).Column
        Set rng = wbauditoria.Sheets("PT RL1 Before-values").Range(wbauditoria.Sheets("PT RL1 Before-values").Cells(1, 1), wbauditoria.Sheets("PT RL1 Before-values").Cells(lastRow, lastCol))
        wbauditoria.Sheets("Validation RL1").Range("A2").Resize(rng.Rows.Count, rng.Columns.Count).Value = rng.Value

        'Calcula ultima fila y columna de la hoja de validacion
        Columns("A:A").Insert Shift:=xlToRight
        lastRow = Sheets("Validation RL1").Cells(Sheets("Validation RL1").Rows.Count, 3).End(xlUp).row
        lastCol = Sheets("Validation RL1").Cells(2, Sheets("Validation RL1").Columns.Count).End(xlToLeft).Column

        'Coloca los titulos en la fila 2
        ThisWorkbook.Sheets("Anexxes").Range("L6:AK6").Copy
        wbauditoria.Sheets("Validation RL1").Range("A2").PasteSpecial Paste:=xlPasteValues
        Application.CutCopyMode = False
        Columns("A:AQ").AutoFit

        'Coloca titulos en celdas combinadas
        Range("H1:M1").Merge
        Range("H1").Value = "BEFORE ENTRIES"
        Range("N1:S1").Merge
        Range("N1").Value = "AFTER ENTRIES"
        Range("T1:Y1").Merge
        Range("T1").Value = "DIFFERENCES"
'        Range("AA1:AF1").Merge
'        Range("AA1").Value = "VALIDATION"

        With Range("A2:G2")
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With

        With Range("H1:M2")
            .Font.Bold = True
            .Interior.Color = RGB(209, 242, 235)
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With

        With Range("N1:S2")
            .Font.Bold = True
            .Interior.Color = RGB(250, 229, 211)
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With

        With Range("T1:Y2")
            .Font.Bold = True
            .Interior.Color = RGB(250, 219, 216)
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With

        With Range("Z1:Z2")
            .Font.Bold = True
            .Interior.Color = RGB(212, 230, 241)
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With

'        With Range("AA1:AH2")
'            .Font.Bold = True
'            .Interior.Color = RGB(232, 218, 239)
'            .HorizontalAlignment = xlCenter
'            .VerticalAlignment = xlCenter
'        End With
'
'        Columns("A:AH").AutoFit
'
        'Formula de la columa A
        lastRow = Sheets("Validation RL1").Cells(Sheets("Validation RL1").Rows.Count, 3).End(xlUp).row
        lastCol = Sheets("Validation RL1").Cells(2, Sheets("Validation RL1").Columns.Count).End(xlToLeft).Column
        Sheets("Validation RL1").Range("A3:A" & lastRow) = "=COUNTIF('PT RL1 After-values'!C[1],'Validation RL1'!RC[2])"
'
'        'Coloca el resto de las formulas
'        lastRow = Sheets("Validation RL1").Cells(Sheets("Validation RL1").Rows.Count, 3).End(xlUp).row
'        Sheets("Validation RL1").Range("N3:N" & lastRow) = "=VLOOKUP(RC[-12],'PT RL1 After-values'!C1:C12,7,0)"
'        Sheets("Validation RL1").Range("O3:O" & lastRow) = "=VLOOKUP(RC[-13],'PT RL1 After-values'!C1:C12,8,0)"
'        Sheets("Validation RL1").Range("P3:P" & lastRow) = "=VLOOKUP(RC[-14],'PT RL1 After-values'!C1:C12,9,0)"
'        Sheets("Validation RL1").Range("Q3:Q" & lastRow) = "=VLOOKUP(RC[-15],'PT RL1 After-values'!C1:C12,10,0)"
'        Sheets("Validation RL1").Range("R3:R" & lastRow) = "=VLOOKUP(RC[-16],'PT RL1 After-values'!C1:C12,11,0)"
'        Sheets("Validation RL1").Range("S3:S" & lastRow) = "=VLOOKUP(RC[-17],'PT RL1 After-values'!C1:C12,12,0)"
'        Sheets("Validation RL1").Range("T3:T" & lastRow) = "=RC[-6]-RC[-12]"
'        Sheets("Validation RL1").Range("U3:U" & lastRow) = "=RC[-6]-RC[-12]"
'        Sheets("Validation RL1").Range("V3:V" & lastRow) = "=RC[-6]-RC[-12]"
'        Sheets("Validation RL1").Range("W3:W" & lastRow) = "=RC[-6]-RC[-12]"
'        Sheets("Validation RL1").Range("X3:X" & lastRow) = "=RC[-6]-RC[-12]"
'        Sheets("Validation RL1").Range("Y3:Y" & lastRow) = "=RC[-6]-RC[-12]"
'        Sheets("Validation RL1").Range("Z3:Z" & lastRow) = "=VLOOKUP(RC[-23],'BoxJ entry'!C1:C4,4,0)"
'        Sheets("Validation RL1").Range("AA3:AA" & lastRow) = "=ROUND(IF(RC[-7]=0,0,RC[-7]-RC[-1]),2)"
'        Sheets("Validation RL1").Range("AB3:AB" & lastRow) = "=ROUND(IF(RC[-7]=0,0,RC[-7]-RC[-2]),2)"
'        Sheets("Validation RL1").Range("AC3:AC" & lastRow) = "=ROUND(IF(RC[-7]=0,0,RC[-7]-RC[-3]),2)"
'        Sheets("Validation RL1").Range("AD3:AD" & lastRow) = "=IF((RC[-9]+RC[-8])=RC[-7],0,RC[-7])"
'        Sheets("Validation RL1").Range("AE3:AE" & lastRow) = "=ROUND(IF(RC[-7]=0,0,RC[-7]-RC[-5]),2)"
'        Sheets("Validation RL1").Range("AF3:AF" & lastRow) = "=ROUND(IF(RC[-7]=0,0,RC[-7]-RC[-6]),2)"
'        Sheets("Validation RL1").Range("AG3:AG" & lastRow) = "=VLOOKUP(RC[-30],'BoxJ entry'!C1:C2,2,0)"
'        Sheets("Validation RL1").Range("AH3:AH" & lastRow) = "=RC[-1]-RC[-15]"
'        Columns("A:AH").AutoFit


        '------ Crea la hoja de validacion del T4

        'Copia los datos de la hoja de before del T4
        Sheets("Validation T4").Activate
        lastRow = Sheets("PT T4 Before-values").Cells(Sheets("PT T4 Before-values").Rows.Count, 1).End(xlUp).row
        lastCol = Sheets("PT T4 Before-values").Cells(1, Sheets("PT T4 Before-values").Columns.Count).End(xlToLeft).Column
        Set rng = wbauditoria.Sheets("PT T4 Before-values").Range(wbauditoria.Sheets("PT T4 Before-values").Cells(1, 1), wbauditoria.Sheets("PT T4 Before-values").Cells(lastRow, lastCol))
        wbauditoria.Sheets("Validation T4").Range("A2").Resize(rng.Rows.Count, rng.Columns.Count).Value = rng.Value

        'Calcula ultima fila y columna de la hoja de validacion
        Columns("A:A").Insert Shift:=xlToRight
        lastRow = Sheets("Validation T4").Cells(Sheets("Validation T4").Rows.Count, 3).End(xlUp).row
        lastCol = Sheets("Validation T4").Cells(2, Sheets("Validation T4").Columns.Count).End(xlToLeft).Column

        'Coloca los titulos en la fila 2
        ThisWorkbook.Sheets("Anexxes").Range("L5:AN5").Copy
        wbauditoria.Sheets("Validation T4").Range("A2").PasteSpecial Paste:=xlPasteValues
        Application.CutCopyMode = False
        Columns("A:AU").AutoFit

        'Coloca titulos en celdas combinadas
        Range("H1:N1").Merge
        Range("H1").Value = "BEFORE ENTRIES"
        Range("O1:U1").Merge
        Range("O1").Value = "AFTER ENTRIES"
        Range("V1:AB1").Merge
        Range("V1").Value = "DIFFERENCES"
        Range("R1:T1").Merge
'        Range("R1").Value = "VALIDATION"
'        Range("U1:AA1").Merge
'        Range("U1").Value = "RL1 VALIDATION"

        With Range("A2:G2")
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With

        With Range("H1:N2")
            .Font.Bold = True
            .Interior.Color = RGB(209, 242, 235)
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With

        With Range("O1:U2")
            .Font.Bold = True
            .Interior.Color = RGB(250, 229, 211)
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With

        With Range("V1:AB2")
            .Font.Bold = True
            .Interior.Color = RGB(250, 219, 216)
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With

        With Range("AC1:AC2")
            .Font.Bold = True
            .Interior.Color = RGB(212, 230, 241)
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With

'        With Range("R1:T2")
'            .Font.Bold = True
'            .Interior.Color = RGB(232, 218, 239)
'            .HorizontalAlignment = xlCenter
'            .VerticalAlignment = xlCenter
'        End With
'
'        With Range("U1:AA2")
'            .Font.Bold = True
'            .Interior.Color = RGB(250, 229, 211)
'            .HorizontalAlignment = xlCenter
'            .VerticalAlignment = xlCenter
'        End With
        Columns("A:AC").AutoFit

        'Formula de la columa A
        lastRow = Sheets("Validation T4").Cells(Sheets("Validation T4").Rows.Count, 3).End(xlUp).row
        lastCol = Sheets("Validation T4").Cells(2, Sheets("Validation T4").Columns.Count).End(xlToLeft).Column
        Sheets("Validation T4").Range("A3:A" & lastRow) = "=COUNTIF('PT T4 After-values'!C[1],'Validation T4'!RC[2])"
'
'        'Coloca el resto de las formulas
'        lastRow = Sheets("Validation T4").Cells(Sheets("Validation T4").Rows.Count, 3).End(xlUp).row
'        Sheets("Validation T4").Range("K3:K" & lastRow) = "=VLOOKUP(RC[-9],'PT T4 After-values'!C1:C9,7,0)"
'        Sheets("Validation T4").Range("L3:L" & lastRow) = "=VLOOKUP(RC[-10],'PT T4 After-values'!C1:C9,8,0)"
'        Sheets("Validation T4").Range("M3:M" & lastRow) = "=VLOOKUP(RC[-11],'PT T4 After-values'!C1:C9,9,0)"
'        Sheets("Validation T4").Range("N3:N" & lastRow) = "=RC[-3]-RC[-6]"
'        Sheets("Validation T4").Range("O3:O" & lastRow) = "=RC[-3]-RC[-6]"
'        Sheets("Validation T4").Range("P3:P" & lastRow) = "=RC[-3]-RC[-6]"
'        Sheets("Validation T4").Range("Q3:Q" & lastRow) = "=VLOOKUP(RC[-14],'BoxJ entry'!C1:C4,4,0)"
'        Sheets("Validation T4").Range("R3:R" & lastRow) = "=ROUND(IF(RC[-4]=0,0,RC[-4]-RC17),2)"
'        Sheets("Validation T4").Range("S3:S" & lastRow) = "=ROUND(IF(RC[-4]=0,0,RC[-4]-RC17),2)"
'        Sheets("Validation T4").Range("T3:T" & lastRow) = "=ROUND(IF(RC[-4]=0,0,RC[-4]-RC17),2)"
'        Sheets("Validation T4").Range("U3:U" & lastRow) = "=VLOOKUP(RC[-18],'Validation RL1'!C[-18]:C[-2],13,0)"
'        Sheets("Validation T4").Range("V3:V" & lastRow) = "=VLOOKUP(RC[-19],'Validation RL1'!C3:C19,14,0)"
'        Sheets("Validation T4").Range("W3:W" & lastRow) = "=RC[-2]-RC[-12]"
'        Sheets("Validation T4").Range("X3:X" & lastRow) = "=RC[-2]-RC[-12]"
'        Sheets("Validation T4").Range("Y3:Y" & lastRow) = "=VLOOKUP(RC[-22],'Validation RL1'!C3:C19,15,0)"
'        Sheets("Validation T4").Range("Z3:Z" & lastRow) = "=(RC[-5]+RC[-4])-RC[-1]"
'        Sheets("Validation T4").Range("AA3:AA" & lastRow) = "=RC[-2]-RC[-6]"
'        Columns("A:AC").AutoFit

        'Ordena hojas
        Sheets("BoxJ entry").Move Before:=Sheets(1)
        Sheets("Validation RL1").Move After:=Sheets("BoxJ entry")
        Sheets("Validation T4").Move After:=Sheets("Validation RL1")
        Sheets("Validation RL1").Tab.Color = RGB(250, 219, 216)
        Sheets("Validation T4").Tab.Color = RGB(250, 229, 211)

        'Organiza la hoja de BoxJ entry
        Sheets("BoxJ entry").Activate
        Range("A1").Value = "SAP ID"
        Range("B1").Value = "CORRECT BOX J"
        Range("C1").Value = "BOX J TOTAL"
        Range("D1").Value = "SAP DIFFRENCE (ENTRY)"
         With Range("A1:D1")
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
            .Interior.Color = RGB(235, 237, 239)
        End With
        Columns("A:D").AutoFit

        'Finaliza
        wbauditoria.Save
        wbauditoria.Close
        MsgBox "The process has been completed. Please enter the document and complete the data in the 'BoxJ entry' sheet.", vbInformation

End Sub

Sub Pension_Entries()

        '------ Crea la hoja de validacion del T4

        'Copia los datos de la hoja de before del T4
        Sheets("Validation T4").Activate
        lastRow = Sheets("PT T4 Before-values").Cells(Sheets("PT T4 Before-values").Rows.Count, 1).End(xlUp).row
        lastCol = Sheets("PT T4 Before-values").Cells(1, Sheets("PT T4 Before-values").Columns.Count).End(xlToLeft).Column
        Set rng = wbauditoria.Sheets("PT T4 Before-values").Range(wbauditoria.Sheets("PT T4 Before-values").Cells(1, 1), wbauditoria.Sheets("PT T4 Before-values").Cells(lastRow, lastCol))
        wbauditoria.Sheets("Validation T4").Range("A2").Resize(rng.Rows.Count, rng.Columns.Count).Value = rng.Value

        'Calcula ultima fila y columna de la hoja de validacion
        Columns("A:A").Insert Shift:=xlToRight
        lastRow = Sheets("Validation T4").Cells(Sheets("Validation T4").Rows.Count, 3).End(xlUp).row
        lastCol = Sheets("Validation T4").Cells(2, Sheets("Validation T4").Columns.Count).End(xlToLeft).Column

        'Coloca los titulos en la fila 2
        ThisWorkbook.Sheets("Anexxes").Range("L7:V7").Copy
        wbauditoria.Sheets("Validation T4").Range("A2").PasteSpecial Paste:=xlPasteValues
        Application.CutCopyMode = False
        Columns("A:K").AutoFit

        'Coloca titulos en celdas combinadas
        Range("H1").Value = "BEFORE ENTRIES"
        Range("I1").Value = "AFTER ENTRIES"
        Range("J1").Value = "DIFFERENCES"
'        Range("R1").Value = "VALIDATION"
'        Range("U1:AA1").Merge
'        Range("U1").Value = "RL1 VALIDATION"

        With Range("A2:G2")
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With

        With Range("H1:H2")
            .Font.Bold = True
            .Interior.Color = RGB(209, 242, 235)
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With

        With Range("I1:I2")
            .Font.Bold = True
            .Interior.Color = RGB(250, 229, 211)
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With

        With Range("J1:J2")
            .Font.Bold = True
            .Interior.Color = RGB(250, 219, 216)
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With

        With Range("K1:K2")
            .Font.Bold = True
            .Interior.Color = RGB(212, 230, 241)
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With

'        With Range("R1:T2")
'            .Font.Bold = True
'            .Interior.Color = RGB(232, 218, 239)
'            .HorizontalAlignment = xlCenter
'            .VerticalAlignment = xlCenter
'        End With
'
'        With Range("U1:AA2")
'            .Font.Bold = True
'            .Interior.Color = RGB(250, 229, 211)
'            .HorizontalAlignment = xlCenter
'            .VerticalAlignment = xlCenter
'        End With
        Columns("A:AC").AutoFit

        'Formula de la columa A
        lastRow = Sheets("Validation T4").Cells(Sheets("Validation T4").Rows.Count, 3).End(xlUp).row
        lastCol = Sheets("Validation T4").Cells(2, Sheets("Validation T4").Columns.Count).End(xlToLeft).Column
        Sheets("Validation T4").Range("A3:A" & lastRow) = "=COUNTIF('PT T4 After-values'!C[1],'Validation T4'!RC[2])"
'
'        'Coloca el resto de las formulas
'        lastRow = Sheets("Validation T4").Cells(Sheets("Validation T4").Rows.Count, 3).End(xlUp).row
'        Sheets("Validation T4").Range("K3:K" & lastRow) = "=VLOOKUP(RC[-9],'PT T4 After-values'!C1:C9,7,0)"
'        Sheets("Validation T4").Range("L3:L" & lastRow) = "=VLOOKUP(RC[-10],'PT T4 After-values'!C1:C9,8,0)"
'        Sheets("Validation T4").Range("M3:M" & lastRow) = "=VLOOKUP(RC[-11],'PT T4 After-values'!C1:C9,9,0)"
'        Sheets("Validation T4").Range("N3:N" & lastRow) = "=RC[-3]-RC[-6]"
'        Sheets("Validation T4").Range("O3:O" & lastRow) = "=RC[-3]-RC[-6]"
'        Sheets("Validation T4").Range("P3:P" & lastRow) = "=RC[-3]-RC[-6]"
'        Sheets("Validation T4").Range("Q3:Q" & lastRow) = "=VLOOKUP(RC[-14],'BoxJ entry'!C1:C4,4,0)"
'        Sheets("Validation T4").Range("R3:R" & lastRow) = "=ROUND(IF(RC[-4]=0,0,RC[-4]-RC17),2)"
'        Sheets("Validation T4").Range("S3:S" & lastRow) = "=ROUND(IF(RC[-4]=0,0,RC[-4]-RC17),2)"
'        Sheets("Validation T4").Range("T3:T" & lastRow) = "=ROUND(IF(RC[-4]=0,0,RC[-4]-RC17),2)"
'        Sheets("Validation T4").Range("U3:U" & lastRow) = "=VLOOKUP(RC[-18],'Validation RL1'!C[-18]:C[-2],13,0)"
'        Sheets("Validation T4").Range("V3:V" & lastRow) = "=VLOOKUP(RC[-19],'Validation RL1'!C3:C19,14,0)"
'        Sheets("Validation T4").Range("W3:W" & lastRow) = "=RC[-2]-RC[-12]"
'        Sheets("Validation T4").Range("X3:X" & lastRow) = "=RC[-2]-RC[-12]"
'        Sheets("Validation T4").Range("Y3:Y" & lastRow) = "=VLOOKUP(RC[-22],'Validation RL1'!C3:C19,15,0)"
'        Sheets("Validation T4").Range("Z3:Z" & lastRow) = "=(RC[-5]+RC[-4])-RC[-1]"
'        Sheets("Validation T4").Range("AA3:AA" & lastRow) = "=RC[-2]-RC[-6]"
'        Columns("A:AC").AutoFit

        'Ordena hojas
        Sheets("BoxJ entry").Move Before:=Sheets(1)
        Sheets("Validation T4").Move After:=Sheets("BoxJ entry")
        Sheets("Validation T4").Tab.Color = RGB(250, 229, 211)

        'Organiza la hoja de BoxJ entry
        Sheets("BoxJ entry").Activate
        Range("A1").Value = "SAP ID"
        Range("B1").Value = "CORRECT BOX J"
        Range("C1").Value = "BOX J TOTAL"
        Range("D1").Value = "SAP DIFFRENCE (ENTRY)"
         With Range("A1:D1")
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
            .Interior.Color = RGB(235, 237, 239)
        End With
        Columns("A:D").AutoFit

        'Finaliza
        wbauditoria.Save
        wbauditoria.Close
        MsgBox "The process has been completed. Please enter the document and complete the data in the 'BoxJ entry' sheet.", vbInformation

End Sub

Sub CPP_QPP_validation()

'------ Crea la hoja de validacion del RL1
        
        'Copia los datos de la hoja de before del RL1
        Sheets("Validation RL1").Activate
        lastRow = Sheets("PT RL1 Before-values").Cells(Sheets("PT RL1 Before-values").Rows.Count, 3).End(xlUp).row
        lastCol = Sheets("PT RL1 Before-values").Cells(1, Sheets("PT RL1 Before-values").Columns.Count).End(xlToLeft).Column
        Set rng = wbauditoria.Sheets("PT RL1 Before-values").Range(wbauditoria.Sheets("PT RL1 Before-values").Cells(1, 1), wbauditoria.Sheets("PT RL1 Before-values").Cells(lastRow, lastCol))
        wbauditoria.Sheets("Validation RL1").Range("A2").Resize(rng.Rows.Count, rng.Columns.Count).Value = rng.Value

        'Calcula ultima fila y columna de la hoja de validacion
        Columns("A:A").Insert Shift:=xlToRight
        lastRow = Sheets("Validation RL1").Cells(Sheets("Validation RL1").Rows.Count, 3).End(xlUp).row
        lastCol = Sheets("Validation RL1").Cells(2, Sheets("Validation RL1").Columns.Count).End(xlToLeft).Column

        'Coloca los titulos en la fila 2
        ThisWorkbook.Sheets("Anexxes").Range("L9:AQ9").Copy
        wbauditoria.Sheets("Validation RL1").Range("A2").PasteSpecial Paste:=xlPasteValues
        Application.CutCopyMode = False
        Columns("A:AQ").AutoFit

        'Coloca titulos en celdas combinadas
        Range("H1:O1").Merge
        Range("H1").Value = "BEFORE ENTRIES"
        Range("P1:W1").Merge
        Range("P1").Value = "AFTER ENTRIES"
        Range("X1:AE1").Merge
        Range("X1").Value = "DIFFERENCES"
'        Range("AA1:AF1").Merge
'        Range("AA1").Value = "VALIDATION"

        With Range("A2:G2")
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With

        With Range("H1:O2")
            .Font.Bold = True
            .Interior.Color = RGB(209, 242, 235)
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With

        With Range("P1:W2")
            .Font.Bold = True
            .Interior.Color = RGB(250, 229, 211)
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With

        With Range("X1:AE2")
            .Font.Bold = True
            .Interior.Color = RGB(250, 219, 216)
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With

        With Range("AF1:AF2")
            .Font.Bold = True
            .Interior.Color = RGB(212, 230, 241)
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With

'        With Range("AA1:AH2")
'            .Font.Bold = True
'            .Interior.Color = RGB(232, 218, 239)
'            .HorizontalAlignment = xlCenter
'            .VerticalAlignment = xlCenter
'        End With
'
'        Columns("A:AH").AutoFit
'
        'Formula de la columa A
        lastRow = Sheets("Validation RL1").Cells(Sheets("Validation RL1").Rows.Count, 3).End(xlUp).row
        lastCol = Sheets("Validation RL1").Cells(2, Sheets("Validation RL1").Columns.Count).End(xlToLeft).Column
        Sheets("Validation RL1").Range("A3:A" & lastRow) = "=COUNTIF('PT RL1 After-values'!C[1],'Validation RL1'!RC[2])"
'
'        'Coloca el resto de las formulas
'        lastRow = Sheets("Validation RL1").Cells(Sheets("Validation RL1").Rows.Count, 3).End(xlUp).row
'        Sheets("Validation RL1").Range("N3:N" & lastRow) = "=VLOOKUP(RC[-12],'PT RL1 After-values'!C1:C12,7,0)"
'        Sheets("Validation RL1").Range("O3:O" & lastRow) = "=VLOOKUP(RC[-13],'PT RL1 After-values'!C1:C12,8,0)"
'        Sheets("Validation RL1").Range("P3:P" & lastRow) = "=VLOOKUP(RC[-14],'PT RL1 After-values'!C1:C12,9,0)"
'        Sheets("Validation RL1").Range("Q3:Q" & lastRow) = "=VLOOKUP(RC[-15],'PT RL1 After-values'!C1:C12,10,0)"
'        Sheets("Validation RL1").Range("R3:R" & lastRow) = "=VLOOKUP(RC[-16],'PT RL1 After-values'!C1:C12,11,0)"
'        Sheets("Validation RL1").Range("S3:S" & lastRow) = "=VLOOKUP(RC[-17],'PT RL1 After-values'!C1:C12,12,0)"
'        Sheets("Validation RL1").Range("T3:T" & lastRow) = "=RC[-6]-RC[-12]"
'        Sheets("Validation RL1").Range("U3:U" & lastRow) = "=RC[-6]-RC[-12]"
'        Sheets("Validation RL1").Range("V3:V" & lastRow) = "=RC[-6]-RC[-12]"
'        Sheets("Validation RL1").Range("W3:W" & lastRow) = "=RC[-6]-RC[-12]"
'        Sheets("Validation RL1").Range("X3:X" & lastRow) = "=RC[-6]-RC[-12]"
'        Sheets("Validation RL1").Range("Y3:Y" & lastRow) = "=RC[-6]-RC[-12]"
'        Sheets("Validation RL1").Range("Z3:Z" & lastRow) = "=VLOOKUP(RC[-23],'BoxJ entry'!C1:C4,4,0)"
'        Sheets("Validation RL1").Range("AA3:AA" & lastRow) = "=ROUND(IF(RC[-7]=0,0,RC[-7]-RC[-1]),2)"
'        Sheets("Validation RL1").Range("AB3:AB" & lastRow) = "=ROUND(IF(RC[-7]=0,0,RC[-7]-RC[-2]),2)"
'        Sheets("Validation RL1").Range("AC3:AC" & lastRow) = "=ROUND(IF(RC[-7]=0,0,RC[-7]-RC[-3]),2)"
'        Sheets("Validation RL1").Range("AD3:AD" & lastRow) = "=IF((RC[-9]+RC[-8])=RC[-7],0,RC[-7])"
'        Sheets("Validation RL1").Range("AE3:AE" & lastRow) = "=ROUND(IF(RC[-7]=0,0,RC[-7]-RC[-5]),2)"
'        Sheets("Validation RL1").Range("AF3:AF" & lastRow) = "=ROUND(IF(RC[-7]=0,0,RC[-7]-RC[-6]),2)"
'        Sheets("Validation RL1").Range("AG3:AG" & lastRow) = "=VLOOKUP(RC[-30],'BoxJ entry'!C1:C2,2,0)"
'        Sheets("Validation RL1").Range("AH3:AH" & lastRow) = "=RC[-1]-RC[-15]"
'        Columns("A:AH").AutoFit


        '------ Crea la hoja de validacion del T4

        'Copia los datos de la hoja de before del T4
        Sheets("Validation T4").Activate
        lastRow = Sheets("PT T4 Before-values").Cells(Sheets("PT T4 Before-values").Rows.Count, 1).End(xlUp).row
        lastCol = Sheets("PT T4 Before-values").Cells(1, Sheets("PT T4 Before-values").Columns.Count).End(xlToLeft).Column
        Set rng = wbauditoria.Sheets("PT T4 Before-values").Range(wbauditoria.Sheets("PT T4 Before-values").Cells(1, 1), wbauditoria.Sheets("PT T4 Before-values").Cells(lastRow, lastCol))
        wbauditoria.Sheets("Validation T4").Range("A2").Resize(rng.Rows.Count, rng.Columns.Count).Value = rng.Value

        'Calcula ultima fila y columna de la hoja de validacion
        Columns("A:A").Insert Shift:=xlToRight
        lastRow = Sheets("Validation T4").Cells(Sheets("Validation T4").Rows.Count, 3).End(xlUp).row
        lastCol = Sheets("Validation T4").Cells(2, Sheets("Validation T4").Columns.Count).End(xlToLeft).Column

        'Coloca los titulos en la fila 2
        ThisWorkbook.Sheets("Anexxes").Range("L8:AH8").Copy
        wbauditoria.Sheets("Validation T4").Range("A2").PasteSpecial Paste:=xlPasteValues
        Application.CutCopyMode = False
        Columns("A:AU").AutoFit

        'Coloca titulos en celdas combinadas
        Range("H1:L1").Merge
        Range("H1").Value = "BEFORE ENTRIES"
        Range("M1:Q1").Merge
        Range("M1").Value = "AFTER ENTRIES"
        Range("R1:V1").Merge
        Range("R1").Value = "DIFFERENCES"
'        Range("R1:T1").Merge
'        Range("R1").Value = "VALIDATION"
'        Range("U1:AA1").Merge
'        Range("U1").Value = "RL1 VALIDATION"

        With Range("A2:G2")
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With

        With Range("H1:L2")
            .Font.Bold = True
            .Interior.Color = RGB(209, 242, 235)
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With

        With Range("M1:Q2")
            .Font.Bold = True
            .Interior.Color = RGB(250, 229, 211)
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With

        With Range("R1:V2")
            .Font.Bold = True
            .Interior.Color = RGB(250, 219, 216)
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With

        With Range("W1:W2")
            .Font.Bold = True
            .Interior.Color = RGB(212, 230, 241)
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With

'        With Range("R1:T2")
'            .Font.Bold = True
'            .Interior.Color = RGB(232, 218, 239)
'            .HorizontalAlignment = xlCenter
'            .VerticalAlignment = xlCenter
'        End With
'
'        With Range("U1:AA2")
'            .Font.Bold = True
'            .Interior.Color = RGB(250, 229, 211)
'            .HorizontalAlignment = xlCenter
'            .VerticalAlignment = xlCenter
'        End With
        Columns("A:AC").AutoFit

        'Formula de la columa A
        lastRow = Sheets("Validation T4").Cells(Sheets("Validation T4").Rows.Count, 3).End(xlUp).row
        lastCol = Sheets("Validation T4").Cells(2, Sheets("Validation T4").Columns.Count).End(xlToLeft).Column
        Sheets("Validation T4").Range("A3:A" & lastRow) = "=COUNTIF('PT T4 After-values'!C[1],'Validation T4'!RC[2])"
'
'        'Coloca el resto de las formulas
'        lastRow = Sheets("Validation T4").Cells(Sheets("Validation T4").Rows.Count, 3).End(xlUp).row
'        Sheets("Validation T4").Range("K3:K" & lastRow) = "=VLOOKUP(RC[-9],'PT T4 After-values'!C1:C9,7,0)"
'        Sheets("Validation T4").Range("L3:L" & lastRow) = "=VLOOKUP(RC[-10],'PT T4 After-values'!C1:C9,8,0)"
'        Sheets("Validation T4").Range("M3:M" & lastRow) = "=VLOOKUP(RC[-11],'PT T4 After-values'!C1:C9,9,0)"
'        Sheets("Validation T4").Range("N3:N" & lastRow) = "=RC[-3]-RC[-6]"
'        Sheets("Validation T4").Range("O3:O" & lastRow) = "=RC[-3]-RC[-6]"
'        Sheets("Validation T4").Range("P3:P" & lastRow) = "=RC[-3]-RC[-6]"
'        Sheets("Validation T4").Range("Q3:Q" & lastRow) = "=VLOOKUP(RC[-14],'BoxJ entry'!C1:C4,4,0)"
'        Sheets("Validation T4").Range("R3:R" & lastRow) = "=ROUND(IF(RC[-4]=0,0,RC[-4]-RC17),2)"
'        Sheets("Validation T4").Range("S3:S" & lastRow) = "=ROUND(IF(RC[-4]=0,0,RC[-4]-RC17),2)"
'        Sheets("Validation T4").Range("T3:T" & lastRow) = "=ROUND(IF(RC[-4]=0,0,RC[-4]-RC17),2)"
'        Sheets("Validation T4").Range("U3:U" & lastRow) = "=VLOOKUP(RC[-18],'Validation RL1'!C[-18]:C[-2],13,0)"
'        Sheets("Validation T4").Range("V3:V" & lastRow) = "=VLOOKUP(RC[-19],'Validation RL1'!C3:C19,14,0)"
'        Sheets("Validation T4").Range("W3:W" & lastRow) = "=RC[-2]-RC[-12]"
'        Sheets("Validation T4").Range("X3:X" & lastRow) = "=RC[-2]-RC[-12]"
'        Sheets("Validation T4").Range("Y3:Y" & lastRow) = "=VLOOKUP(RC[-22],'Validation RL1'!C3:C19,15,0)"
'        Sheets("Validation T4").Range("Z3:Z" & lastRow) = "=(RC[-5]+RC[-4])-RC[-1]"
'        Sheets("Validation T4").Range("AA3:AA" & lastRow) = "=RC[-2]-RC[-6]"
'        Columns("A:AC").AutoFit

        'Ordena hojas
        Sheets("BoxJ entry").Move Before:=Sheets(1)
        Sheets("Validation RL1").Move After:=Sheets("BoxJ entry")
        Sheets("Validation T4").Move After:=Sheets("Validation RL1")
        Sheets("Validation RL1").Tab.Color = RGB(250, 219, 216)
        Sheets("Validation T4").Tab.Color = RGB(250, 229, 211)

        'Organiza la hoja de BoxJ entry
        Sheets("BoxJ entry").Activate
        Range("A1").Value = "SAP ID"
        Range("B1").Value = "CORRECT BOX J"
        Range("C1").Value = "BOX J TOTAL"
        Range("D1").Value = "SAP DIFFRENCE (ENTRY)"
         With Range("A1:D1")
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
            .Interior.Color = RGB(235, 237, 239)
        End With
        Columns("A:D").AutoFit

        'Finaliza
        wbauditoria.Save
        wbauditoria.Close
        MsgBox "The process has been completed. Please enter the document and complete the data in the 'BoxJ entry' sheet.", vbInformation

End Sub


Sub EI_boxes_validation()

'------ Crea la hoja de validacion del RL1
        
        'Copia los datos de la hoja de before del RL1
        Sheets("Validation RL1").Activate
        lastRow = Sheets("PT RL1 Before-values").Cells(Sheets("PT RL1 Before-values").Rows.Count, 3).End(xlUp).row
        lastCol = Sheets("PT RL1 Before-values").Cells(1, Sheets("PT RL1 Before-values").Columns.Count).End(xlToLeft).Column
        Set rng = wbauditoria.Sheets("PT RL1 Before-values").Range(wbauditoria.Sheets("PT RL1 Before-values").Cells(1, 1), wbauditoria.Sheets("PT RL1 Before-values").Cells(lastRow, lastCol))
        wbauditoria.Sheets("Validation RL1").Range("A2").Resize(rng.Rows.Count, rng.Columns.Count).Value = rng.Value

        'Calcula ultima fila y columna de la hoja de validacion
        Columns("A:A").Insert Shift:=xlToRight
        lastRow = Sheets("Validation RL1").Cells(Sheets("Validation RL1").Rows.Count, 3).End(xlUp).row
        lastCol = Sheets("Validation RL1").Cells(2, Sheets("Validation RL1").Columns.Count).End(xlToLeft).Column

        'Coloca los titulos en la fila 2
        ThisWorkbook.Sheets("Anexxes").Range("L11:V11").Copy
        wbauditoria.Sheets("Validation RL1").Range("A2").PasteSpecial Paste:=xlPasteValues
        Application.CutCopyMode = False
        Columns("A:AQ").AutoFit

        'Coloca titulos en celdas combinadas
        Range("H1").Value = "BEFORE ENTRIES"
        Range("I1").Value = "AFTER ENTRIES"
        Range("J1").Value = "DIFFERENCES"
'        Range("AA1:AF1").Merge
'        Range("AA1").Value = "VALIDATION"

        With Range("A2:G2")
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With

        With Range("H1:H2")
            .Font.Bold = True
            .Interior.Color = RGB(209, 242, 235)
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With

        With Range("I1:I2")
            .Font.Bold = True
            .Interior.Color = RGB(250, 229, 211)
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With

        With Range("J1:J2")
            .Font.Bold = True
            .Interior.Color = RGB(250, 219, 216)
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With

        With Range("K1:K2")
            .Font.Bold = True
            .Interior.Color = RGB(212, 230, 241)
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With

'        With Range("AA1:AH2")
'            .Font.Bold = True
'            .Interior.Color = RGB(232, 218, 239)
'            .HorizontalAlignment = xlCenter
'            .VerticalAlignment = xlCenter
'        End With
'
'        Columns("A:AH").AutoFit
'
        'Formula de la columa A
        lastRow = Sheets("Validation RL1").Cells(Sheets("Validation RL1").Rows.Count, 3).End(xlUp).row
        lastCol = Sheets("Validation RL1").Cells(2, Sheets("Validation RL1").Columns.Count).End(xlToLeft).Column
        Sheets("Validation RL1").Range("A3:A" & lastRow) = "=COUNTIF('PT RL1 After-values'!C[1],'Validation RL1'!RC[2])"
'
'        'Coloca el resto de las formulas
'        lastRow = Sheets("Validation RL1").Cells(Sheets("Validation RL1").Rows.Count, 3).End(xlUp).row
'        Sheets("Validation RL1").Range("N3:N" & lastRow) = "=VLOOKUP(RC[-12],'PT RL1 After-values'!C1:C12,7,0)"
'        Sheets("Validation RL1").Range("O3:O" & lastRow) = "=VLOOKUP(RC[-13],'PT RL1 After-values'!C1:C12,8,0)"
'        Sheets("Validation RL1").Range("P3:P" & lastRow) = "=VLOOKUP(RC[-14],'PT RL1 After-values'!C1:C12,9,0)"
'        Sheets("Validation RL1").Range("Q3:Q" & lastRow) = "=VLOOKUP(RC[-15],'PT RL1 After-values'!C1:C12,10,0)"
'        Sheets("Validation RL1").Range("R3:R" & lastRow) = "=VLOOKUP(RC[-16],'PT RL1 After-values'!C1:C12,11,0)"
'        Sheets("Validation RL1").Range("S3:S" & lastRow) = "=VLOOKUP(RC[-17],'PT RL1 After-values'!C1:C12,12,0)"
'        Sheets("Validation RL1").Range("T3:T" & lastRow) = "=RC[-6]-RC[-12]"
'        Sheets("Validation RL1").Range("U3:U" & lastRow) = "=RC[-6]-RC[-12]"
'        Sheets("Validation RL1").Range("V3:V" & lastRow) = "=RC[-6]-RC[-12]"
'        Sheets("Validation RL1").Range("W3:W" & lastRow) = "=RC[-6]-RC[-12]"
'        Sheets("Validation RL1").Range("X3:X" & lastRow) = "=RC[-6]-RC[-12]"
'        Sheets("Validation RL1").Range("Y3:Y" & lastRow) = "=RC[-6]-RC[-12]"
'        Sheets("Validation RL1").Range("Z3:Z" & lastRow) = "=VLOOKUP(RC[-23],'BoxJ entry'!C1:C4,4,0)"
'        Sheets("Validation RL1").Range("AA3:AA" & lastRow) = "=ROUND(IF(RC[-7]=0,0,RC[-7]-RC[-1]),2)"
'        Sheets("Validation RL1").Range("AB3:AB" & lastRow) = "=ROUND(IF(RC[-7]=0,0,RC[-7]-RC[-2]),2)"
'        Sheets("Validation RL1").Range("AC3:AC" & lastRow) = "=ROUND(IF(RC[-7]=0,0,RC[-7]-RC[-3]),2)"
'        Sheets("Validation RL1").Range("AD3:AD" & lastRow) = "=IF((RC[-9]+RC[-8])=RC[-7],0,RC[-7])"
'        Sheets("Validation RL1").Range("AE3:AE" & lastRow) = "=ROUND(IF(RC[-7]=0,0,RC[-7]-RC[-5]),2)"
'        Sheets("Validation RL1").Range("AF3:AF" & lastRow) = "=ROUND(IF(RC[-7]=0,0,RC[-7]-RC[-6]),2)"
'        Sheets("Validation RL1").Range("AG3:AG" & lastRow) = "=VLOOKUP(RC[-30],'BoxJ entry'!C1:C2,2,0)"
        Sheets("Validation RL1").Range("AH3:AH" & lastRow) = "=RC[-1]-RC[-15]"
        Columns("A:AH").AutoFit


        '------ Crea la hoja de validacion del T4

        'Copia los datos de la hoja de before del T4
        Sheets("Validation T4").Activate
        lastRow = Sheets("PT T4 Before-values").Cells(Sheets("PT T4 Before-values").Rows.Count, 1).End(xlUp).row
        lastCol = Sheets("PT T4 Before-values").Cells(1, Sheets("PT T4 Before-values").Columns.Count).End(xlToLeft).Column
        Set rng = wbauditoria.Sheets("PT T4 Before-values").Range(wbauditoria.Sheets("PT T4 Before-values").Cells(1, 1), wbauditoria.Sheets("PT T4 Before-values").Cells(lastRow, lastCol))
        wbauditoria.Sheets("Validation T4").Range("A2").Resize(rng.Rows.Count, rng.Columns.Count).Value = rng.Value

        'Calcula ultima fila y columna de la hoja de validacion
        Columns("A:A").Insert Shift:=xlToRight
        lastRow = Sheets("Validation T4").Cells(Sheets("Validation T4").Rows.Count, 3).End(xlUp).row
        lastCol = Sheets("Validation T4").Cells(2, Sheets("Validation T4").Columns.Count).End(xlToLeft).Column

        'Coloca los titulos en la fila 2
        ThisWorkbook.Sheets("Anexxes").Range("L10:AE10").Copy
        wbauditoria.Sheets("Validation T4").Range("A2").PasteSpecial Paste:=xlPasteValues
        Application.CutCopyMode = False
        Columns("A:AU").AutoFit

        'Coloca titulos en celdas combinadas
        Range("H1:K1").Merge
        Range("H1").Value = "BEFORE ENTRIES"
        Range("L1:O1").Merge
        Range("L1").Value = "AFTER ENTRIES"
        Range("P1:S1").Merge
        Range("P1").Value = "DIFFERENCES"
'        Range("R1:T1").Merge
'        Range("R1").Value = "VALIDATION"
'        Range("U1:AA1").Merge
'        Range("U1").Value = "RL1 VALIDATION"

        With Range("A2:G2")
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With

        With Range("H1:K2")
            .Font.Bold = True
            .Interior.Color = RGB(209, 242, 235)
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With

        With Range("L1:O2")
            .Font.Bold = True
            .Interior.Color = RGB(250, 229, 211)
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With

        With Range("P1:S2")
            .Font.Bold = True
            .Interior.Color = RGB(250, 219, 216)
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With

        With Range("T1:T2")
            .Font.Bold = True
            .Interior.Color = RGB(212, 230, 241)
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With

'        With Range("R1:T2")
'            .Font.Bold = True
'            .Interior.Color = RGB(232, 218, 239)
'            .HorizontalAlignment = xlCenter
'            .VerticalAlignment = xlCenter
'        End With
'
'        With Range("U1:AA2")
'            .Font.Bold = True
'            .Interior.Color = RGB(250, 229, 211)
'            .HorizontalAlignment = xlCenter
'            .VerticalAlignment = xlCenter
'        End With
        Columns("A:AC").AutoFit

        'Formula de la columa A
        lastRow = Sheets("Validation T4").Cells(Sheets("Validation T4").Rows.Count, 3).End(xlUp).row
        lastCol = Sheets("Validation T4").Cells(2, Sheets("Validation T4").Columns.Count).End(xlToLeft).Column
        Sheets("Validation T4").Range("A3:A" & lastRow) = "=COUNTIF('PT T4 After-values'!C[1],'Validation T4'!RC[2])"
'
'        'Coloca el resto de las formulas
'        lastRow = Sheets("Validation T4").Cells(Sheets("Validation T4").Rows.Count, 3).End(xlUp).row
'        Sheets("Validation T4").Range("K3:K" & lastRow) = "=VLOOKUP(RC[-9],'PT T4 After-values'!C1:C9,7,0)"
'        Sheets("Validation T4").Range("L3:L" & lastRow) = "=VLOOKUP(RC[-10],'PT T4 After-values'!C1:C9,8,0)"
'        Sheets("Validation T4").Range("M3:M" & lastRow) = "=VLOOKUP(RC[-11],'PT T4 After-values'!C1:C9,9,0)"
'        Sheets("Validation T4").Range("N3:N" & lastRow) = "=RC[-3]-RC[-6]"
'        Sheets("Validation T4").Range("O3:O" & lastRow) = "=RC[-3]-RC[-6]"
'        Sheets("Validation T4").Range("P3:P" & lastRow) = "=RC[-3]-RC[-6]"
'        Sheets("Validation T4").Range("Q3:Q" & lastRow) = "=VLOOKUP(RC[-14],'BoxJ entry'!C1:C4,4,0)"
'        Sheets("Validation T4").Range("R3:R" & lastRow) = "=ROUND(IF(RC[-4]=0,0,RC[-4]-RC17),2)"
'        Sheets("Validation T4").Range("S3:S" & lastRow) = "=ROUND(IF(RC[-4]=0,0,RC[-4]-RC17),2)"
'        Sheets("Validation T4").Range("T3:T" & lastRow) = "=ROUND(IF(RC[-4]=0,0,RC[-4]-RC17),2)"
'        Sheets("Validation T4").Range("U3:U" & lastRow) = "=VLOOKUP(RC[-18],'Validation RL1'!C[-18]:C[-2],13,0)"
'        Sheets("Validation T4").Range("V3:V" & lastRow) = "=VLOOKUP(RC[-19],'Validation RL1'!C3:C19,14,0)"
'        Sheets("Validation T4").Range("W3:W" & lastRow) = "=RC[-2]-RC[-12]"
'        Sheets("Validation T4").Range("X3:X" & lastRow) = "=RC[-2]-RC[-12]"
'        Sheets("Validation T4").Range("Y3:Y" & lastRow) = "=VLOOKUP(RC[-22],'Validation RL1'!C3:C19,15,0)"
'        Sheets("Validation T4").Range("Z3:Z" & lastRow) = "=(RC[-5]+RC[-4])-RC[-1]"
'        Sheets("Validation T4").Range("AA3:AA" & lastRow) = "=RC[-2]-RC[-6]"
'        Columns("A:AC").AutoFit

        'Ordena hojas
        Sheets("BoxJ entry").Move Before:=Sheets(1)
        Sheets("Validation RL1").Move After:=Sheets("BoxJ entry")
        Sheets("Validation T4").Move After:=Sheets("Validation RL1")
        Sheets("Validation RL1").Tab.Color = RGB(250, 219, 216)
        Sheets("Validation T4").Tab.Color = RGB(250, 229, 211)

        'Organiza la hoja de BoxJ entry
        Sheets("BoxJ entry").Activate
        Range("A1").Value = "SAP ID"
        Range("B1").Value = "CORRECT BOX J"
        Range("C1").Value = "BOX J TOTAL"
        Range("D1").Value = "SAP DIFFRENCE (ENTRY)"
         With Range("A1:D1")
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
            .Interior.Color = RGB(235, 237, 239)
        End With
        Columns("A:D").AutoFit

        'Finaliza
        wbauditoria.Save
        wbauditoria.Close
        MsgBox "The process has been completed. Please enter the document and complete the data in the 'BoxJ entry' sheet.", vbInformation

End Sub
