Attribute VB_Name = "YE_COMBINADA"
Option Explicit
'variables para todo el proyecto

Dim BASE, wbPlantillaRL1vsT4, wbBasesPlantilla, wbBases As Workbook
Dim ws As Worksheet
Dim Fecha1, Fecha2, Fecha3, mes, Mes_Texto, BPA, año, dia, ruta, Ruta_pais, Ruta_Año, Ruta_Audi, ruta_Base, texto, ExisteArchivo As String
Dim hoja As Integer
Dim lastRow, lastCol, i As Long
Dim titles As Variant
    
Sub Ejecutar_YE_COMBINADA()

' Verificar si hay datos en las celdas I8 y M8
If ThisWorkbook.Sheets("Home Page").Range("I8").Value = "" Or ThisWorkbook.Sheets("Home Page").Range("E18").Value = "" Then
    MsgBox "Incomplete data, please enter the data before executing.", vbExclamation
    Exit Sub
End If

' Llama a cada una de las funciones
InicializarVariables
CrearCarpetas
Desarrollo_COMBINADA

MsgBox "RL1 VS T4 audit completed. Please access the document to perform the relevant reviews.", vbInformation

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

'Rutas
ruta = ThisWorkbook.Path & "\"
Ruta_pais = ruta & "YEAR END CA"
Ruta_Audi = Ruta_pais & "\" & "BOXES AUDITS"
Ruta_Año = Ruta_Audi & "\" & año

End Sub
Sub CrearCarpetas()

'PAIS
Ruta_pais = ruta & "YEAR END CA"
If Dir(Ruta_pais, vbDirectory + vbHidden) = "" Then
    If Dir(Ruta_pais & vbDirectory + vbHidden) = "" Then MkDir Ruta_pais
End If

'CARPETA DE AUDITORIAS
Ruta_Audi = Ruta_pais & "\" & "BOXES AUDITS"
If Dir(Ruta_Audi, vbDirectory + vbHidden) = "" Then
    If Dir(Ruta_Audi & vbDirectory + vbHidden) = "" Then MkDir Ruta_Audi
End If

'CARPETA AÑO
Ruta_Año = Ruta_Audi & "\" & año
If Dir(Ruta_Año, vbDirectory + vbHidden) = "" Then
    If Dir(Ruta_Audi & vbDirectory + vbHidden) = "" Then MkDir Ruta_Año
End If
     
End Sub

Sub Desarrollo_COMBINADA()

'-------------------------- CREAR ARCHIVO DE LAS BASES --------------------------
Application.AskToUpdateLinks = False
    
' Abre el reporte de las bases
If ruta_Base <> "Falso" Then
    
    'Verifica si el archivo de las bases existe o no
    ExisteArchivo = Dir(Ruta_Año & "\" & año & mes & " Archivo Bases Boxes Validation" & ".xlsx")

    If ExisteArchivo = "" Then 'Si no existe el archvio de las bases en la carpeta
    MsgBox "First run the RL1 and T4 macro audit before running this one", vbExclamation
    Exit Sub

    Else 'Si el documento de las bases ya existe, entonces lo abre y crea las dos nuevas tablas dinamicas
        Set wbBases = Workbooks.Open(Ruta_Año & "\" & año & mes & " Archivo Bases Boxes Validation" & ".xlsx")
        wbBases.Activate
        
        'Crea tabla dinamica que solo use los datos de los empleados de QC
            Sheets("BASE T4").Activate
            Dim ult_Tabla As Long
            ult_Tabla = Cells(Rows.Count, "A").End(xlUp).row
            Dim rangoTabla1 As Range
            Set rangoTabla1 = Sheets("BASE T4").Range("A1:P" & ult_Tabla)
            'ActiveSheet.ListObjects.Add(xlSrcRange, rangoTabla1, , xlYes).Name = "Tabla1"
            
            'Crear tabla dinamica
            Dim celdaTablaDinamica1 As Range
            Set celdaTablaDinamica1 = Sheets("TD T4 WITHOUT BN").Range("A1")
            Dim tablaDinamica1 As PivotTable
            
            'Activa campos y le pone formato tabular
            ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
                rangoTabla1, Version:=6).CreatePivotTable TableDestination:= _
                celdaTablaDinamica1, TableName:="tablaDinamica1", DefaultVersion:=6
            
            Sheets("TD T4 WITHOUT BN").Select
            With ActiveSheet.PivotTables("tablaDinamica1").PivotFields("PERNR")
                .Orientation = xlRowField
                .Position = 1
            End With
            With ActiveSheet.PivotTables("tablaDinamica1").PivotFields("WRKAR")
                .Orientation = xlRowField
                .Position = 2
            End With
            With ActiveSheet.PivotTables("tablaDinamica1").PivotFields("SLART")
                .Orientation = xlColumnField
                .Position = 1
            End With
            ActiveSheet.PivotTables("tablaDinamica1").AddDataField ActiveSheet.PivotTables( _
                "tablaDinamica1").PivotFields("BETRG"), "Suma de BETRG", xlSum
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
            ActiveSheet.PivotTables("tablaDinamica1").RepeatAllLabels xlRepeatLabels
            ActiveSheet.PivotTables("tablaDinamica1").RowAxisLayout xlTabularRow
            wbBases.Save
            
              
            'Crea tabla dinamica que solo use los datos de los empleados de QC
            Sheets("BASE T4").Activate
            ult_Tabla = Cells(Rows.Count, "A").End(xlUp).row
            Set rangoTabla1 = Sheets("BASE T4").Range("A1:P" & ult_Tabla)
            'ActiveSheet.ListObjects.Add(xlSrcRange, rangoTabla1, , xlYes).Name = "Tabla1"
            
            'Crear tabla dinamica
            Dim celdaTablaDinamica2 As Range
            Set celdaTablaDinamica2 = Sheets("TD T4 ONLY QC").Range("A1")
            Dim tablaDinamica2 As PivotTable
            
            'Activa campos y le pone formato tabular
            ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
                rangoTabla1, Version:=6).CreatePivotTable TableDestination:= _
                celdaTablaDinamica2, TableName:="tablaDinamica2", DefaultVersion:=6
            
            Sheets("TD T4 ONLY QC").Select
            With ActiveSheet.PivotTables("tablaDinamica2").PivotFields("PERNR")
                .Orientation = xlRowField
                .Position = 1
            End With
            With ActiveSheet.PivotTables("tablaDinamica2").PivotFields("WRKAR")
                .Orientation = xlRowField
                .Position = 2
            End With
            With ActiveSheet.PivotTables("tablaDinamica2").PivotFields("SLART")
                .Orientation = xlColumnField
                .Position = 1
            End With
            ActiveSheet.PivotTables("tablaDinamica2").AddDataField ActiveSheet.PivotTables( _
                "tablaDinamica2").PivotFields("BETRG"), "Suma de BETRG", xlSum
            ActiveSheet.PivotTables("tablaDinamica2").PivotFields("PERNR").Subtotals = _
                Array(False, False, False, False, False, False, False, False, False, False, False, False)
            ActiveSheet.PivotTables("tablaDinamica2").PivotFields("PERSONID").Subtotals = _
                Array(False, False, False, False, False, False, False, False, False, False, False, False)
            ActiveSheet.PivotTables("tablaDinamica2").PivotFields("FORML").Subtotals = _
                Array(False, False, False, False, False, False, False, False, False, False, False, False)
            ActiveSheet.PivotTables("tablaDinamica2").PivotFields("INDX1").Subtotals = _
                Array(False, False, False, False, False, False, False, False, False, False, False, False)
            ActiveSheet.PivotTables("tablaDinamica2").PivotFields("BUSNM").Subtotals = _
                Array(False, False, False, False, False, False, False, False, False, False, False, False)
            ActiveSheet.PivotTables("tablaDinamica2").PivotFields("WRKAR").Subtotals = _
                Array(False, False, False, False, False, False, False, False, False, False, False, False)
            ActiveSheet.PivotTables("tablaDinamica2").PivotFields("WTEXT").Subtotals = _
                Array(False, False, False, False, False, False, False, False, False, False, False, False)
            ActiveSheet.PivotTables("tablaDinamica2").PivotFields("KOSTL").Subtotals = _
                Array(False, False, False, False, False, False, False, False, False, False, False, False)
            ActiveSheet.PivotTables("tablaDinamica2").PivotFields("KTEXT").Subtotals = _
                Array(False, False, False, False, False, False, False, False, False, False, False, False)
            ActiveSheet.PivotTables("tablaDinamica2").PivotFields("NACHN").Subtotals = _
                Array(False, False, False, False, False, False, False, False, False, False, False, False)
            ActiveSheet.PivotTables("tablaDinamica2").PivotFields("VORNA").Subtotals = _
                Array(False, False, False, False, False, False, False, False, False, False, False, False)
            ActiveSheet.PivotTables("tablaDinamica2").PivotFields("SLART").Subtotals = _
                Array(False, False, False, False, False, False, False, False, False, False, False, False)
            ActiveSheet.PivotTables("tablaDinamica2").PivotFields("STEXT").Subtotals = _
                Array(False, False, False, False, False, False, False, False, False, False, False, False)
            ActiveSheet.PivotTables("tablaDinamica2").PivotFields("LGART").Subtotals = _
                Array(False, False, False, False, False, False, False, False, False, False, False, False)
            ActiveSheet.PivotTables("tablaDinamica2").PivotFields("LTEXT").Subtotals = _
                Array(False, False, False, False, False, False, False, False, False, False, False, False)
            ActiveSheet.PivotTables("tablaDinamica2").PivotFields("BETRG").Subtotals = _
                Array(False, False, False, False, False, False, False, False, False, False, False, False)
            ActiveSheet.PivotTables("tablaDinamica2").RepeatAllLabels xlRepeatLabels
            ActiveSheet.PivotTables("tablaDinamica2").RowAxisLayout xlTabularRow
            wbBases.Save
            
'            With ActiveSheet.PivotTables("tablaDinamica2").PivotFields("WRKAR")
'                .PivotItems("").Visible = False
'                .PivotItems("AB").Visible = False
'                .PivotItems("BC").Visible = False
'                .PivotItems("MB").Visible = False
'                .PivotItems("NB").Visible = False
'                .PivotItems("NS").Visible = False
'                .PivotItems("ON").Visible = False
'                .PivotItems("SK").Visible = False
'                .PivotItems("US").Visible = False
'                .PivotItems("ZZ").Visible = False
'            End With

    End If
    
End If


'-------------------------- CREAR ARCHIVO DE LA PLANTILLA DEL RL1 vs T4 --------------------------

'Hace copia de la plantilla del RL1 vs T4
On Error Resume Next
Kill Ruta_Año & "\" & año & mes & " T4 vs RL1 Audits" & ".xlsx"
On Error GoTo 0

Set wbPlantillaRL1vsT4 = Workbooks.Open(ruta & "T4 vs RL1 Audits.xlsx")
wbPlantillaRL1vsT4.Activate
ActiveWorkbook.SaveCopyAs Filename:=Ruta_Año & "\" & año & mes & " T4 vs RL1 Audits" & ".xlsx"
wbPlantillaRL1vsT4.Close False

'Abre los archivos correspondientes
Set wbPlantillaRL1vsT4 = Workbooks.Open(Ruta_Año & "\" & año & mes & " T4 vs RL1 Audits" & ".xlsx")

'Pasa la tabla dinamica como valores al documento del T4
wbBases.Activate
wbBases.Sheets("TD T4").Activate
lastRow = Cells(Rows.Count, "A").End(xlUp).row
Range("A2:AE" & lastRow - 1).Select
Range("A2:AE" & lastRow - 1).Copy
wbPlantillaRL1vsT4.Activate
Sheets("TD T4").Activate
Range("C1").Select
Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False
Application.CutCopyMode = False
With Range("A1:AE1")
        .Font.Bold = True
        .Interior.Color = RGB(9, 61, 147)
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
End With

'Texto a numero
lastRow = Cells(Rows.Count, "C").End(xlUp).row
For i = 2 To lastRow
    If IsNumeric(Range("c" & i).Value) Then
        Range("c" & i).Value = Val(Range("c" & i).Value)
    End If
Next i

'Formula del concatenar el KEY NUMBER
Range("B1").Value = "KEY NUMBER"
lastRow = Cells(Rows.Count, "C").End(xlUp).row
Sheets("TD T4").Range("B2:B" & lastRow) = "=+CONCATENATE(RC[1],RC[2],RC[3])"
Columns("A:AE").AutoFit

'Elimina los numeros de empleado que son 00000000
lastRow = Cells(Rows.Count, "C").End(xlUp).row
For i = lastRow To 2 Step -1
        If Cells(i, "C").Value = "0" Then
            Rows(i).Delete
        End If
Next i

lastRow = Cells(Rows.Count, "C").End(xlUp).row
For i = lastRow To 2 Step -1
        If Cells(i, "C").Value = "999999" Then
            Rows(i).Delete
        End If
Next i

'Formula del concatenar el KEY NUMBER #2
Range("A1").Value = "KEY NUMBER #2"
lastRow = Cells(Rows.Count, "C").End(xlUp).row
Sheets("TD T4").Range("A2:A" & lastRow) = "=CONCATENATE(RC[2],RC[4])"
Columns("A:AD").AutoFit

'Pasa la tabla dinamica como valores al documento del RL1
wbBases.Activate
wbBases.Sheets("TD RL1").Activate
lastRow = Cells(Rows.Count, "A").End(xlUp).row
Range("A2:AE" & lastRow - 1).Select
Range("A2:AE" & lastRow - 1).Copy
wbPlantillaRL1vsT4.Activate
Sheets("TD RL1").Activate
Range("C1").Select
Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False
Application.CutCopyMode = False
With Range("A1:AE1")
        .Font.Bold = True
        .Interior.Color = RGB(9, 61, 147)
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
End With

'Texto a numero
lastRow = Cells(Rows.Count, "C").End(xlUp).row
For i = 2 To lastRow
    If IsNumeric(Range("C" & i).Value) Then
        Range("C" & i).Value = Val(Range("C" & i).Value)
    End If
Next i

'Formula del concatenar el KEY NUMBER
Range("B1").Value = "KEY NUMBER"
lastRow = Cells(Rows.Count, "C").End(xlUp).row
Sheets("TD RL1").Range("B2:B" & lastRow) = "=+CONCATENATE(RC[1],RC[2],RC[3])"
Columns("A:AE").AutoFit

'Elimina los numeros de empleado que son 00000000
lastRow = Cells(Rows.Count, "C").End(xlUp).row
For i = lastRow To 2 Step -1
        If Cells(i, "C").Value = "0" Then
            Rows(i).Delete
        End If
Next i

lastRow = Cells(Rows.Count, "C").End(xlUp).row
For i = lastRow To 2 Step -1
        If Cells(i, "C").Value = "999999" Then
            Rows(i).Delete
        End If
Next i

'Formula del concatenar el KEY NUMBER
Range("A1").Value = "KEY NUMBER #2"
lastRow = Cells(Rows.Count, "C").End(xlUp).row
Sheets("TD RL1").Range("A2:A" & lastRow) = "=CONCATENATE(RC[2],RC[4])"
Columns("A:AD").AutoFit


'Crea la hoja sin el Bussines number
wbBases.Activate
wbBases.Sheets("TD T4 WITHOUT BN").Activate
lastRow = Cells(Rows.Count, "A").End(xlUp).row
Range("A2:AE" & lastRow - 1).Select
Range("A2:AE" & lastRow - 1).Copy
wbPlantillaRL1vsT4.Activate
Sheets("TD T4 WITHOUT BN").Activate
Range("B1").Select
Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False
Application.CutCopyMode = False
With Range("A1:AE1")
        .Font.Bold = True
        .Interior.Color = RGB(9, 61, 147)
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
End With

'Texto a numero
lastRow = Cells(Rows.Count, "B").End(xlUp).row
For i = 2 To lastRow
    If IsNumeric(Range("b" & i).Value) Then
        Range("B" & i).Value = Val(Range("B" & i).Value)
    End If
Next i

'Formula del concatenar el KEY NUMBER #2
Range("A1").Value = "KEY NUMBER #2"
lastRow = Cells(Rows.Count, "C").End(xlUp).row
Sheets("TD T4 WITHOUT BN").Range("A2:A" & lastRow) = "=CONCATENATE(RC[1],RC[2])"
Columns("A:AE").AutoFit

'Elimina los numeros de empleado que son 00000000
lastRow = Cells(Rows.Count, "A").End(xlUp).row
For i = lastRow To 2 Step -1
        If Cells(i, "B").Value = "0" Then
            Rows(i).Delete
        End If
Next i

lastRow = Cells(Rows.Count, "B").End(xlUp).row
For i = lastRow To 2 Step -1
        If Cells(i, "B").Value = "999999" Then
            Rows(i).Delete
        End If
Next i

'Crea la hoja de solo QC
wbBases.Activate
wbBases.Sheets("TD T4 ONLY QC").Activate
lastRow = Cells(Rows.Count, "A").End(xlUp).row
Range("A2:AE" & lastRow - 1).Select
Range("A2:AE" & lastRow - 1).Copy
wbPlantillaRL1vsT4.Activate
Sheets("TD T4 ONLY QC").Activate
Range("B1").Select
Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False
Application.CutCopyMode = False
With Range("A1:AE1")
        .Font.Bold = True
        .Interior.Color = RGB(9, 61, 147)
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
End With

'Texto a numero
lastRow = Cells(Rows.Count, "C").End(xlUp).row
For i = 2 To lastRow
    If IsNumeric(Range("B" & i).Value) Then
        Range("B" & i).Value = Val(Range("B" & i).Value)
    End If
Next i

'Elimina los numeros de empleado que son 00000000
lastRow = Cells(Rows.Count, "C").End(xlUp).row
For i = lastRow To 2 Step -1
        If Cells(i, "B").Value = "0" Then
            Rows(i).Delete
        End If
Next i

lastRow = Cells(Rows.Count, "C").End(xlUp).row
For i = lastRow To 2 Step -1
        If Cells(i, "B").Value = "999999" Then
            Rows(i).Delete
        End If
Next i

'Formula del concatenar el KEY NUMBER #2
Range("A1").Value = "KEY NUMBER #2"
lastRow = Cells(Rows.Count, "C").End(xlUp).row
Sheets("TD T4 ONLY QC").Range("A2:A" & lastRow) = "=CONCATENATE(RC[1],RC[2])"
Columns("A:AD").AutoFit

'Deja solo los que son QC
Sheets.Add.Name = "Hoja extra"
Sheets("TD T4 ONLY QC").Activate
lastRow = Cells(Rows.Count, "A").End(xlUp).row
Rows("1:1").AutoFilter
Rows("1:1").AutoFilter Field:=3, Criteria1:="=QC"
Range("A1:AE" & lastRow).SpecialCells(xlCellTypeVisible).Copy
Sheets("Hoja extra").Activate
Range("A1").PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
Application.CutCopyMode = False
Application.DisplayAlerts = False
Sheets("TD T4 ONLY QC").Delete
Application.DisplayAlerts = True
Sheets("Hoja extra").Name = "TD T4 ONLY QC"
Columns("A:AD").AutoFit

'Cierra documento de las bases
wbBases.Activate
wbBases.Save
wbBases.Close
wbPlantillaRL1vsT4.Activate

'-------------------------- Ciclo para pegar las 4 columnas principales en cada una de las hojas--------------------------
'NOTA:


For hoja = 6 To 26

    'Seleccina los datos que va a copiar
    wbPlantillaRL1vsT4.Sheets("TD RL1").Activate
    lastRow = Cells(Rows.Count, "B").End(xlUp).row
    Range("C1:E" & lastRow).Select
    Range("C1:E" & lastRow).Copy
    
    'Pega datos en cada hoja de los reportes
    Set ws = wbPlantillaRL1vsT4.Sheets(hoja)
    ws.Activate
    Range("A3").Select
    Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    
    'Seleccina los datos que va a copiar
    wbPlantillaRL1vsT4.Sheets("TD RL1").Activate
    lastRow = Cells(Rows.Count, "B").End(xlUp).row
    Range("B1:B" & lastRow).Select
    Range("B1:B" & lastRow).Copy
    
    'Pega datos en cada hoja de los reportes
    Set ws = wbPlantillaRL1vsT4.Sheets(hoja)
    ws.Activate
    Range("D3").Select
    Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    
    'Seleccina los datos que va a copiar
    wbPlantillaRL1vsT4.Sheets("TD RL1").Activate
    lastRow = Cells(Rows.Count, "B").End(xlUp).row
    Range("A1:A" & lastRow).Select
    Range("A1:A" & lastRow).Copy
    
    'Pega datos en cada hoja de los reportes
    Set ws = wbPlantillaRL1vsT4.Sheets(hoja)
    ws.Activate
    Range("E3").Select
    Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    
    'Formato de color de los campos de la tabla
    With Range("A3:E3")
        .Font.Bold = True
        .Interior.Color = RGB(9, 61, 147)
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
    End With
    Columns("A:E").AutoFit
    
Next hoja
wbPlantillaRL1vsT4.Save

'-------------------------- REPARTE LAS COLUMNAS DE LOS BOXES PARA LAS AUDIRORIAS --------------------------

Dim wsTDRL1, wsDestino As Worksheet
Dim col As Range
Dim destCol As Long
Dim Posicion_hojas, criterios As Variant

'Setear la hoja de donde toma las columnas
Set wsTDRL1 = wbPlantillaRL1vsT4.Sheets("TD RL1")

'Definir las hojas y los criterios, el orden de los criterios esta segun el orden de las hojas en el archivo
'Nota:
    'Hoja 6 - BOX A - BOX J =BOX 14
    'Hoja 7 - QC EMPLOYEES , IF BOX 16>0, SHOULD HAVE F B1 LINE 96 (CPP CONTRIBUTION)
    'Hoja 8 - QC EMPLOYEES WHO ALSO WORKED IN NON QC PROVINCES, BOX FB1 (CPP CONTRIBUTIONS)= BOX 27 NON QC PROVINCE
    'Hoja 9 - QC EMPLOYEES WHO ALSO WORKED IN NON QC PROVINCES, BOX FB2 (CPP SECOND CONTRIBUTIONS)= BOX16A NON QC PROVINCE
    'Hoja 10 - QC EMPLOYEES WHO ALSO WORKED IN NON QC PROVINCES, BOX FG2 (CPP PENSIONABLE EARNINGS)= BOX 26 NON QC PROVINCE
    'Hoja 11 - BOX B = BOX BS = (BOX 17 + BOX 17A)
    'Hoja 12 - BOX BB= BOX16A
    'Hoja 13 - BOX G=BOX 26
    'Hoja 14 - BOX C=BOX 18
    'Hoja 15 - BOX H = BOX 55
    'Hoja 16 - BOX I = BOX 56
    'Hoja 17 - BOX D =BOX 20
    'Hoja 18 - BOX F=BOX 44
    'Hoja 19 - BOX O -Code RJ = BOX 66+BOX 67
    'Hoja 20 - BOX L=BOX40
    'Hoja 21 - BOX W =BOX 34
    'Hoja 22 - BOX 30 = BOX V
    'Hoja 23 - BOX 46= BOX N
    'Hoja 24 - F235=BOX 85
    'Hoja 25 - NO SIN FIELD IN 000 AND THE SIN SHOULD HAVE 9 DIGITS - (DATA BASE IS CCYR + SQ01 IT002)
    'Hoja 26 - VALIDATE IF SECOND T4S ISSUED FOR EMPLOYEES ARE OKAY


Posicion_hojas = Array(6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24) 'Las hojas en las que va a poner las columas

'Criterios para cada hoja. IMPORTANE: cada renglon corresponde al criterio de la hoja, empezando desde la hoja 6
criterios = Array( _
    Array("Box A", "Box J"), _
    Array("FB1"), _
    Array("FB1"), _
    Array("FB2"), _
    Array("Box G", "FG2"), _
    Array("Box B", "Box BB", "Box Bs"), _
    Array("Box B", "Box BB"), _
    Array("Box G"), _
    Array("Box C"), _
    Array("Box H"), _
    Array("Box I"), _
    Array("Box D"), _
    Array("Box F"), _
    Array("Box O"), _
    Array("Box L"), _
    Array("Box W"), _
    Array("Box V"), _
    Array("Box N"), _
    Array("F235") _
)

'Iterar sobre las hojas
For i = LBound(Posicion_hojas) To UBound(Posicion_hojas)

    'Asignar la hoja de destino
    Set wsDestino = wbPlantillaRL1vsT4.Sheets(Posicion_hojas(i))
    lastRow = wsDestino.Cells(wsDestino.Rows.Count, 1).End(xlUp).row
    
    'Indica que inicie a pegar los datos desde la columna E
    destCol = 6
    
    'Realiza la busqueda
    For Each col In wsTDRL1.Rows(1).Cells
        
        If Not IsError(Application.Match(col.Value, criterios(i), 0)) Then
            wsTDRL1.Range(col, wsTDRL1.Cells(lastRow, col.Column)).Copy
            wsDestino.Cells(3, destCol).PasteSpecial Paste:=xlPasteValues
            destCol = destCol + 1
        End If
    Next col

    Application.CutCopyMode = False
    
Next i

'Formatos de las celdas de los titulos
For hoja = 6 To 26

    Set ws = wbPlantillaRL1vsT4.Sheets(hoja)
    ws.Activate
    lastCol = Cells(3, Columns.Count).End(xlToLeft).Column
    With Range(Cells(3, 6), Cells(3, lastCol))
        .Interior.Color = RGB(243, 156, 18)
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
    End With
    Columns("A:I").AutoFit
    
Next hoja
wbPlantillaRL1vsT4.Save

'-------------------------- CREA LAS FORMULAS DE LAS AUDITORIAS --------------------------

'Hoja 6 - BOX A - BOX J =BOX 14
With Sheets(6)
    .Activate
    lastCol = .Cells(3, .Columns.Count).End(xlToLeft).Column
    lastRow = .Cells(.Rows.Count, 1).End(xlUp).row
    
    ' Títulos y configuración
    titles = Array("BOX 14", "AUDIT", "COMMENTS")
    
    For i = 0 To UBound(titles)
        With .Cells(3, lastCol + 1 + i)
            .Value = titles(i)
            If InStr(1, titles(i), "AUDIT", vbTextCompare) > 0 Then
                .Interior.Color = RGB(146, 208, 80)
            Else
                .Interior.Color = RGB(9, 61, 147)
            End If
            .Font.Color = RGB(255, 255, 255)
            .HorizontalAlignment = xlCenter
            .Font.Bold = True
        End With
    Next i
    
    ' Fórmula
    .Range(.Cells(4, lastCol + 1), .Cells(lastRow, lastCol + 1)).FormulaR1C1 = "=VLOOKUP(RC[-3],'TD T4'!C1:C6,6,0)"
    .Range(.Cells(4, lastCol + 2), .Cells(lastRow, lastCol + 2)).FormulaR1C1 = "=RC[-3]-RC[-2]-RC[-1]"
    .Columns("A:M").AutoFit
End With

'Hoja 7 - QC EMPLOYEES , IF BOX 16>0, SHOULD HAVE F B1 LINE 96 (CPP CONTRIBUTION)

With Sheets(7)
    .Activate
    lastCol = .Cells(3, .Columns.Count).End(xlToLeft).Column
    lastRow = .Cells(.Rows.Count, 1).End(xlUp).row
    
    ' Títulos y configuración
    titles = Array("BOX 16", "AUDIT", "COMMENTS")
    
    For i = 0 To UBound(titles)
        With .Cells(3, lastCol + 1 + i)
            .Value = titles(i)
            If InStr(1, titles(i), "AUDIT", vbTextCompare) > 0 Then
                .Interior.Color = RGB(146, 208, 80)
            Else
                .Interior.Color = RGB(9, 61, 147)
            End If
            .Font.Color = RGB(255, 255, 255)
            .HorizontalAlignment = xlCenter
            .Font.Bold = True
        End With
    Next i
    
    ' Fórmula
    .Range(.Cells(4, lastCol + 1), .Cells(lastRow, lastCol + 1)).FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-2],'TD T4'!C1:C8,7,0),"" "")"
    .Range(.Cells(4, lastCol + 2), .Cells(lastRow, lastCol + 2)).FormulaR1C1 = "=IFERROR(RC[-2]-RC[-1],"" "")"
    .Columns("A:M").AutoFit
End With

'Hoja 8 - QC EMPLOYEES WHO ALSO WORKED IN NON QC PROVINCES, BOX FB1 (CPP CONTRIBUTIONS)= BOX 27 NON QC PROVINCE

With Sheets(8)
    .Activate
    lastCol = .Cells(3, .Columns.Count).End(xlToLeft).Column
    lastRow = .Cells(.Rows.Count, 1).End(xlUp).row
    
    ' Títulos y configuración
    titles = Array("BOX 27", "AUDIT", "COMMENTS")
    
    For i = 0 To UBound(titles)
        With .Cells(3, lastCol + 1 + i)
            .Value = titles(i)
            If InStr(1, titles(i), "AUDIT", vbTextCompare) > 0 Then
                .Interior.Color = RGB(146, 208, 80)
            Else
                .Interior.Color = RGB(9, 61, 147)
            End If
            .Font.Color = RGB(255, 255, 255)
            .HorizontalAlignment = xlCenter
            .Font.Bold = True
        End With
    Next i
    
    ' Fórmula
    .Range(.Cells(4, lastCol + 1), .Cells(lastRow, lastCol + 1)).FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-6],'TD T4 WITHOUT BN'!C2:C15,14,0),"" "")"
    .Range(.Cells(4, lastCol + 2), .Cells(lastRow, lastCol + 2)).FormulaR1C1 = "=IFERROR(RC[-2]-RC[-1],"" "")"
    .Columns("A:M").AutoFit
End With

'Hoja 9 - QC EMPLOYEES WHO ALSO WORKED IN NON QC PROVINCES, BOX FB2 (CPP SECOND CONTRIBUTIONS)= BOX16A NON QC PROVINCE
'-----------------No funciona bien porque FB02 todavia no existe
With Sheets(9)
    .Activate
    lastCol = .Cells(3, .Columns.Count).End(xlToLeft).Column
    lastRow = .Cells(.Rows.Count, 1).End(xlUp).row
    
    ' Títulos y configuración
    titles = Array("BOX 16A", "AUDIT", "COMMENTS")
    
    For i = 0 To UBound(titles)
        With .Cells(3, lastCol + 1 + i)
            .Value = titles(i)
            If InStr(1, titles(i), "AUDIT", vbTextCompare) > 0 Then
                .Interior.Color = RGB(146, 208, 80)
            Else
                .Interior.Color = RGB(9, 61, 147)
            End If
            .Font.Color = RGB(255, 255, 255)
            .HorizontalAlignment = xlCenter
            .Font.Bold = True
        End With
    Next i
    
    ' Fórmula
    .Range(.Cells(4, lastCol + 1), .Cells(lastRow, lastCol + 1)).FormulaR1C1 = "=VLOOKUP(RC[-6],'TD T4 WITHOUT BN'!C2:C6,5,0)"
    .Range(.Cells(4, lastCol + 2), .Cells(lastRow, lastCol + 2)).FormulaR1C1 = "=RC[-2]-RC[-1]"
    .Columns("A:M").AutoFit
End With

'Hoja 10 - QC EMPLOYEES WHO ALSO WORKED IN NON QC PROVINCES, BOX FG2 (CPP PENSIONABLE EARNINGS)= BOX 26 NON QC PROVINCE

With Sheets(10)
    .Activate
    lastCol = .Cells(3, .Columns.Count).End(xlToLeft).Column
    lastRow = .Cells(.Rows.Count, 1).End(xlUp).row
    
    ' Títulos y configuración
    titles = Array("BOX 26", "AUDIT 1", "AUDIT 2", "COMMENTS")
    
    For i = 0 To UBound(titles)
        With .Cells(3, lastCol + 1 + i)
            .Value = titles(i)
            If InStr(1, titles(i), "AUDIT", vbTextCompare) > 0 Then
                .Interior.Color = RGB(146, 208, 80)
            Else
                .Interior.Color = RGB(9, 61, 147)
            End If
            .Font.Color = RGB(255, 255, 255)
            .HorizontalAlignment = xlCenter
            .Font.Bold = True
        End With
    Next i
    
    ' Fórmula
    .Range(.Cells(4, lastCol + 1), .Cells(lastRow, lastCol + 1)).FormulaR1C1 = "=VLOOKUP(RC[-7],'TD T4 WITHOUT BN'!C2:C14,13,0)"
    .Range(.Cells(4, lastCol + 2), .Cells(lastRow, lastCol + 2)).FormulaR1C1 = "=IF(RC[-2]-RC[-1]<0,"" "",RC[-2]-RC[-1])"
    .Range(.Cells(4, lastCol + 3), .Cells(lastRow, lastCol + 3)).FormulaR1C1 = "=IF((RC[-4]+RC[-3]-RC[-2])>0,(RC[-4]+RC[-3]-RC[-2]),"" "")"
    .Columns("A:M").AutoFit
End With

'Hoja 11 - BOX B = BOX BS = (BOX 17 + BOX 17A)

With Sheets(11)
    .Activate
    lastCol = .Cells(3, .Columns.Count).End(xlToLeft).Column
    lastRow = .Cells(.Rows.Count, 1).End(xlUp).row
    
    ' Títulos y configuración
    titles = Array("BOX 17", "BOX 17A", "BOX 17+BOX 17A", "AUDIT", "COMMENTS")
    
    For i = 0 To UBound(titles)
        With .Cells(3, lastCol + 1 + i)
            .Value = titles(i)
            If InStr(1, titles(i), "AUDIT", vbTextCompare) > 0 Then
                .Interior.Color = RGB(146, 208, 80)
            Else
                .Interior.Color = RGB(9, 61, 147)
            End If
            .Font.Color = RGB(255, 255, 255)
            .HorizontalAlignment = xlCenter
            .Font.Bold = True
        End With
    Next i
    
    ' Fórmula
    .Range(.Cells(4, lastCol + 1), .Cells(lastRow, lastCol + 1)).FormulaR1C1 = "=VLOOKUP(RC[-4],'TD T4 ONLY QC'!C1:C8,7,0)"
    .Range(.Cells(4, lastCol + 2), .Cells(lastRow, lastCol + 2)).FormulaR1C1 = "=VLOOKUP(RC[-5],'TD T4 ONLY QC'!C1:C8,8,0)"
    .Range(.Cells(4, lastCol + 3), .Cells(lastRow, lastCol + 3)).FormulaR1C1 = "=RC[-2]+RC[-1]"
    .Range(.Cells(4, lastCol + 4), .Cells(lastRow, lastCol + 4)).FormulaR1C1 = "=IF(AND((RC[-6]+RC[-5])=RC[-4],RC[-4]=RC[-1]),""OK"",""REVIEW"")"
    .Columns("A:M").AutoFit
End With

'Hoja 12 - BOX BB=  BOX17A & BOXB =BOX17º

With Sheets(12)
    .Activate
    lastCol = .Cells(3, .Columns.Count).End(xlToLeft).Column
    lastRow = .Cells(.Rows.Count, 1).End(xlUp).row
    
    ' Títulos y configuración
    titles = Array("BOX 17", "BOX17A", "AUDIT 1", "AUDIT 2", "COMMENTS")
    
    For i = 0 To UBound(titles)
        With .Cells(3, lastCol + 1 + i)
            .Value = titles(i)
            If InStr(1, titles(i), "AUDIT", vbTextCompare) > 0 Then
                .Interior.Color = RGB(146, 208, 80)
            Else
                .Interior.Color = RGB(9, 61, 147)
            End If
            .Font.Color = RGB(255, 255, 255)
            .HorizontalAlignment = xlCenter
            .Font.Bold = True
        End With
    Next i
    
    ' Fórmula
    .Range(.Cells(4, lastCol + 1), .Cells(lastRow, lastCol + 1)).FormulaR1C1 = "=VLOOKUP(RC[-3],'TD T4 ONLY QC'!C1:C8,7,0)"
    .Range(.Cells(4, lastCol + 2), .Cells(lastRow, lastCol + 2)).FormulaR1C1 = "=VLOOKUP(RC[-4],'TD T4 ONLY QC'!C1:C8,8,0)"
    .Range(.Cells(4, lastCol + 3), .Cells(lastRow, lastCol + 3)).FormulaR1C1 = "=IF(RC[-3]=RC[-1],""OK"",""""""REVIEW"")"
    .Range(.Cells(4, lastCol + 4), .Cells(lastRow, lastCol + 4)).FormulaR1C1 = "=IF(RC[-5]=RC[-3],""OK"",""REVIEW"")"
    .Columns("A:M").AutoFit

End With

'Hoja 13 - BOX G=BOX 26
With Sheets(13)
    .Activate
    lastCol = .Cells(3, .Columns.Count).End(xlToLeft).Column
    lastRow = .Cells(.Rows.Count, 1).End(xlUp).row
    
    ' Títulos y configuración
    titles = Array("BOX 26", "AUDIT", "COMMENTS")
    
    For i = 0 To UBound(titles)
        With .Cells(3, lastCol + 1 + i)
            .Value = titles(i)
            If InStr(1, titles(i), "AUDIT", vbTextCompare) > 0 Then
                .Interior.Color = RGB(146, 208, 80)
            Else
                .Interior.Color = RGB(9, 61, 147)
            End If
            .Font.Color = RGB(255, 255, 255)
            .HorizontalAlignment = xlCenter
            .Font.Bold = True
        End With
    Next i
    
    ' Fórmula
    .Range(.Cells(4, lastCol + 1), .Cells(lastRow, lastCol + 1)).FormulaR1C1 = "=VLOOKUP(RC[-2],'TD T4 ONLY QC'!C1:C14,14,0)"
    .Range(.Cells(4, lastCol + 2), .Cells(lastRow, lastCol + 2)).FormulaR1C1 = "=RC[-1]-RC[-2]"
    .Columns("A:M").AutoFit

End With

'Hoja 14 - BOX C=BOX 18
With Sheets(14)
    .Activate
    lastCol = .Cells(3, .Columns.Count).End(xlToLeft).Column
    lastRow = .Cells(.Rows.Count, 1).End(xlUp).row
    
    ' Títulos y configuración
    titles = Array("BOX 18", "AUDIT", "COMMENTS")
    
    For i = 0 To UBound(titles)
        With .Cells(3, lastCol + 1 + i)
            .Value = titles(i)
            If InStr(1, titles(i), "AUDIT", vbTextCompare) > 0 Then
                .Interior.Color = RGB(146, 208, 80)
            Else
                .Interior.Color = RGB(9, 61, 147)
            End If
            .Font.Color = RGB(255, 255, 255)
            .HorizontalAlignment = xlCenter
            .Font.Bold = True
        End With
    Next i
    
    ' Fórmula
    .Range(.Cells(4, lastCol + 1), .Cells(lastRow, lastCol + 1)).FormulaR1C1 = "=VLOOKUP(RC[-2],'TD T4 ONLY QC'!C1:C9,9,0)"
    .Range(.Cells(4, lastCol + 2), .Cells(lastRow, lastCol + 2)).FormulaR1C1 = "=RC[-2]-RC[-1]"
    .Columns("A:M").AutoFit

End With

'Hoja 15 - BOX H = BOX 55
With Sheets(15)
    .Activate
    lastCol = .Cells(3, .Columns.Count).End(xlToLeft).Column
    lastRow = .Cells(.Rows.Count, 1).End(xlUp).row
    
    ' Títulos y configuración
    titles = Array("BOX 55", "AUDIT", "COMMENTS")
    
    For i = 0 To UBound(titles)
        With .Cells(3, lastCol + 1 + i)
            .Value = titles(i)
            If InStr(1, titles(i), "AUDIT", vbTextCompare) > 0 Then
                .Interior.Color = RGB(146, 208, 80)
            Else
                .Interior.Color = RGB(9, 61, 147)
            End If
            .Font.Color = RGB(255, 255, 255)
            .HorizontalAlignment = xlCenter
            .Font.Bold = True
        End With
    Next i
    
    ' Fórmula
    .Range(.Cells(4, lastCol + 1), .Cells(lastRow, lastCol + 1)).FormulaR1C1 = "=VLOOKUP(RC[-2],'TD T4 ONLY QC'!C1:C20,20,0)"
    .Range(.Cells(4, lastCol + 2), .Cells(lastRow, lastCol + 2)).FormulaR1C1 = "=RC[-2]-RC[-1]"
    .Columns("A:M").AutoFit

End With

'Hoja 16 - BOX I = BOX 56
With Sheets(16)
    .Activate
    lastCol = .Cells(3, .Columns.Count).End(xlToLeft).Column
    lastRow = .Cells(.Rows.Count, 1).End(xlUp).row
    
    ' Títulos y configuración
    titles = Array("BOX 56", "BOX 14", "BOX 24", "AUDIT 1", "AUDIT 2", "COMMENTS")
    
    For i = 0 To UBound(titles)
        With .Cells(3, lastCol + 1 + i)
            .Value = titles(i)
            If InStr(1, titles(i), "AUDIT", vbTextCompare) > 0 Then
                .Interior.Color = RGB(146, 208, 80)
            Else
                .Interior.Color = RGB(9, 61, 147)
            End If
            .Font.Color = RGB(255, 255, 255)
            .HorizontalAlignment = xlCenter
            .Font.Bold = True
        End With
    Next i
    
    ' Fórmula
    .Range(.Cells(4, lastCol + 1), .Cells(lastRow, lastCol + 1)).FormulaR1C1 = "=VLOOKUP(RC[-2],'TD T4 ONLY QC'!C1:C29,21,0)"
    .Range(.Cells(4, lastCol + 2), .Cells(lastRow, lastCol + 2)).FormulaR1C1 = "=VLOOKUP(RC[-3],'TD T4 ONLY QC'!C1:C29,4,0)"
    .Range(.Cells(4, lastCol + 3), .Cells(lastRow, lastCol + 3)).FormulaR1C1 = "=VLOOKUP(RC[-4],'TD T4 ONLY QC'!C1:C29,13,0)"
    .Range(.Cells(4, lastCol + 4), .Cells(lastRow, lastCol + 4)).FormulaR1C1 = "=RC[-4]-RC[-3]"
    .Range(.Cells(4, lastCol + 5), .Cells(lastRow, lastCol + 5)).FormulaR1C1 = "=IF(OR(RC[-2]=R1C11,RC[-2]=0,RC[-3]=RC[-2]),""OK"",""REVIEW"")"
    .Columns("A:M").AutoFit

End With

'Hoja 17 - BOX D =BOX 20
With Sheets(17)
    .Activate
    lastCol = .Cells(3, .Columns.Count).End(xlToLeft).Column
    lastRow = .Cells(.Rows.Count, 1).End(xlUp).row
    
    ' Títulos y configuración
    titles = Array("BOX 20", "AUDIT", "COMMENTS")
    
    For i = 0 To UBound(titles)
        With .Cells(3, lastCol + 1 + i)
            .Value = titles(i)
            If InStr(1, titles(i), "AUDIT", vbTextCompare) > 0 Then
                .Interior.Color = RGB(146, 208, 80)
            Else
                .Interior.Color = RGB(9, 61, 147)
            End If
            .Font.Color = RGB(255, 255, 255)
            .HorizontalAlignment = xlCenter
            .Font.Bold = True
        End With
    Next i
    
    ' Fórmula
    .Range(.Cells(4, lastCol + 1), .Cells(lastRow, lastCol + 1)).FormulaR1C1 = "=VLOOKUP(RC[-2],'TD T4 ONLY QC'!C1:C11,11,0)"
    .Range(.Cells(4, lastCol + 2), .Cells(lastRow, lastCol + 2)).FormulaR1C1 = "=RC[-2]-RC[-1]"
    .Columns("A:M").AutoFit

End With

'Hoja 18 - BOX F=BOX 44
With Sheets(18)
    .Activate
    lastCol = .Cells(3, .Columns.Count).End(xlToLeft).Column
    lastRow = .Cells(.Rows.Count, 1).End(xlUp).row
    
    ' Títulos y configuración
    titles = Array("BOX 44", "AUDIT", "COMMENTS")
    
    For i = 0 To UBound(titles)
        With .Cells(3, lastCol + 1 + i)
            .Value = titles(i)
            If InStr(1, titles(i), "AUDIT", vbTextCompare) > 0 Then
                .Interior.Color = RGB(146, 208, 80)
            Else
                .Interior.Color = RGB(9, 61, 147)
            End If
            .Font.Color = RGB(255, 255, 255)
            .HorizontalAlignment = xlCenter
            .Font.Bold = True
        End With
    Next i
    
    ' Fórmula
    .Range(.Cells(4, lastCol + 1), .Cells(lastRow, lastCol + 1)).FormulaR1C1 = "=VLOOKUP(RC[-2],'TD T4 ONLY QC'!C1:C29,17,0)"
    .Range(.Cells(4, lastCol + 2), .Cells(lastRow, lastCol + 2)).FormulaR1C1 = "=RC[-2]-RC[-1]"
    .Columns("A:M").AutoFit

End With

'Hoja 19 - BOX O -Code RJ = BOX 66+BOX 67
With Sheets(19)
    .Activate
    lastCol = .Cells(3, .Columns.Count).End(xlToLeft).Column
    lastRow = .Cells(.Rows.Count, 1).End(xlUp).row
    
    ' Títulos y configuración
    titles = Array("BOX 66", "BOX 67", "AUDIT", "COMMENTS")
    
    For i = 0 To UBound(titles)
        With .Cells(3, lastCol + 1 + i)
            .Value = titles(i)
            If InStr(1, titles(i), "AUDIT", vbTextCompare) > 0 Then
                .Interior.Color = RGB(146, 208, 80)
            Else
                .Interior.Color = RGB(9, 61, 147)
            End If
            .Font.Color = RGB(255, 255, 255)
            .HorizontalAlignment = xlCenter
            .Font.Bold = True
        End With
    Next i
    
    ' Fórmula
    .Range(.Cells(4, lastCol + 1), .Cells(lastRow, lastCol + 1)).FormulaR1C1 = "=VLOOKUP(RC[-2],'TD T4 WITHOUT BN'!C1:C29,25,0)"
    .Range(.Cells(4, lastCol + 2), .Cells(lastRow, lastCol + 2)).FormulaR1C1 = "=VLOOKUP(RC[-3],'TD T4 WITHOUT BN'!C1:C29,26,0)"
    .Range(.Cells(4, lastCol + 3), .Cells(lastRow, lastCol + 3)).FormulaR1C1 = "=(RC[-2]+RC[-1])-RC[-3]"
    .Columns("A:M").AutoFit

End With

'Hoja 20 - BOX L=BOX40
With Sheets(20)
    .Activate
    lastCol = .Cells(3, .Columns.Count).End(xlToLeft).Column
    lastRow = .Cells(.Rows.Count, 1).End(xlUp).row
    
    ' Títulos y configuración
    titles = Array("BOX 40", "AUDIT", "COMMENTS")
    
    For i = 0 To UBound(titles)
        With .Cells(3, lastCol + 1 + i)
            .Value = titles(i)
            If InStr(1, titles(i), "AUDIT", vbTextCompare) > 0 Then
                .Interior.Color = RGB(146, 208, 80)
            Else
                .Interior.Color = RGB(9, 61, 147)
            End If
            .Font.Color = RGB(255, 255, 255)
            .HorizontalAlignment = xlCenter
            .Font.Bold = True
        End With
    Next i
    
    ' Fórmula
    .Range(.Cells(4, lastCol + 1), .Cells(lastRow, lastCol + 1)).FormulaR1C1 = "=VLOOKUP(RC[-2],'TD T4 ONLY QC'!C1:C29,24,0)"
    .Range(.Cells(4, lastCol + 2), .Cells(lastRow, lastCol + 2)).FormulaR1C1 = "=RC[-2]-RC[-1]"
    .Columns("A:M").AutoFit

End With

'Hoja 21 - BOX W =BOX 34
With Sheets(21)
    .Activate
    lastCol = .Cells(3, .Columns.Count).End(xlToLeft).Column
    lastRow = .Cells(.Rows.Count, 1).End(xlUp).row
    
    ' Títulos y configuración
    titles = Array("BOX 34", "AUDIT", "COMMENTS")
    
    For i = 0 To UBound(titles)
        With .Cells(3, lastCol + 1 + i)
            .Value = titles(i)
            If InStr(1, titles(i), "AUDIT", vbTextCompare) > 0 Then
                .Interior.Color = RGB(146, 208, 80)
            Else
                .Interior.Color = RGB(9, 61, 147)
            End If
            .Font.Color = RGB(255, 255, 255)
            .HorizontalAlignment = xlCenter
            .Font.Bold = True
        End With
    Next i
    
    ' Fórmula
    .Range(.Cells(4, lastCol + 1), .Cells(lastRow, lastCol + 1)).FormulaR1C1 = "=VLOOKUP(RC[-2],'TD T4 ONLY QC'!C1:C29,23,0)"
    .Range(.Cells(4, lastCol + 2), .Cells(lastRow, lastCol + 2)).FormulaR1C1 = "=RC[-2]-RC[-1]"
    .Columns("A:M").AutoFit

End With

'Hoja 22 - BOX 30 = BOX V
With Sheets(22)
    .Activate
    lastCol = .Cells(3, .Columns.Count).End(xlToLeft).Column
    lastRow = .Cells(.Rows.Count, 1).End(xlUp).row
    
    ' Títulos y configuración
    titles = Array("BOX 30", "AUDIT", "COMMENTS")
    
    For i = 0 To UBound(titles)
        With .Cells(3, lastCol + 1 + i)
            .Value = titles(i)
            If InStr(1, titles(i), "AUDIT", vbTextCompare) > 0 Then
                .Interior.Color = RGB(146, 208, 80)
            Else
                .Interior.Color = RGB(9, 61, 147)
            End If
            .Font.Color = RGB(255, 255, 255)
            .HorizontalAlignment = xlCenter
            .Font.Bold = True
        End With
    Next i
    
    ' Fórmula
    .Range(.Cells(4, lastCol + 1), .Cells(lastRow, lastCol + 1)).FormulaR1C1 = "=VLOOKUP(RC[-2],'TD T4 WITHOUT BN'!C1:C29,22,0)"
    .Range(.Cells(4, lastCol + 2), .Cells(lastRow, lastCol + 2)).FormulaR1C1 = "=RC[-2]-RC[-1]"
    .Columns("A:M").AutoFit

End With

'Hoja 23 - BOX 46= BOX N
With Sheets(23)
    .Activate
    lastCol = .Cells(3, .Columns.Count).End(xlToLeft).Column
    lastRow = .Cells(.Rows.Count, 1).End(xlUp).row
    
    ' Títulos y configuración
    titles = Array("BOX 46", "AUDIT", "COMMENTS")
    
    For i = 0 To UBound(titles)
        With .Cells(3, lastCol + 1 + i)
            .Value = titles(i)
            If InStr(1, titles(i), "AUDIT", vbTextCompare) > 0 Then
                .Interior.Color = RGB(146, 208, 80)
            Else
                .Interior.Color = RGB(9, 61, 147)
            End If
            .Font.Color = RGB(255, 255, 255)
            .HorizontalAlignment = xlCenter
            .Font.Bold = True
        End With
    Next i
    
    ' Fórmula
    .Range(.Cells(4, lastCol + 1), .Cells(lastRow, lastCol + 1)).FormulaR1C1 = "=VLOOKUP(RC[-2],'TD T4 WITHOUT BN'!C1:C29,18,0)"
    .Range(.Cells(4, lastCol + 2), .Cells(lastRow, lastCol + 2)).FormulaR1C1 = "=RC[-1]-RC[-2]"
    .Columns("A:M").AutoFit

End With

'Hoja 24 - F235=BOX 85
With Sheets(24)
    .Activate
    lastCol = .Cells(3, .Columns.Count).End(xlToLeft).Column
    lastRow = .Cells(.Rows.Count, 1).End(xlUp).row
    
    ' Títulos y configuración
    titles = Array("BOX 85", "AUDIT", "COMMENTS")
    
    For i = 0 To UBound(titles)
        With .Cells(3, lastCol + 1 + i)
            .Value = titles(i)
            If InStr(1, titles(i), "AUDIT", vbTextCompare) > 0 Then
                .Interior.Color = RGB(146, 208, 80)
            Else
                .Interior.Color = RGB(9, 61, 147)
            End If
            .Font.Color = RGB(255, 255, 255)
            .HorizontalAlignment = xlCenter
            .Font.Bold = True
        End With
    Next i
    
    ' Fórmula
    .Range(.Cells(4, lastCol + 1), .Cells(lastRow, lastCol + 1)).FormulaR1C1 = "=VLOOKUP(RC[-2],'TD T4 ONLY QC'!C1:C29,28,0)"
    .Range(.Cells(4, lastCol + 2), .Cells(lastRow, lastCol + 2)).FormulaR1C1 = "=RC[-2]-RC[-1]"
    .Columns("A:M").AutoFit

End With

'Hoja 25 - NO SIN FIELD IN 000 AND THE SIN SHOULD HAVE 9 DIGITS - (DATA BASE IS CCYR + SQ01 IT002)
With Sheets(25)
    .Activate
    lastCol = .Cells(3, .Columns.Count).End(xlToLeft).Column
    lastCol = lastCol - 8
    lastRow = .Cells(.Rows.Count, 1).End(xlUp).row
    
    ' Títulos y configuración
    titles = Array("SIN", "AUDIT", "COMMENTS")
    
    For i = 0 To UBound(titles)
        With .Cells(3, lastCol + 1 + i)
            .Value = titles(i)
            If InStr(1, titles(i), "AUDIT", vbTextCompare) > 0 Then
                .Interior.Color = RGB(146, 208, 80)
            Else
                .Interior.Color = RGB(9, 61, 147)
            End If
            .Font.Color = RGB(255, 255, 255)
            .HorizontalAlignment = xlCenter
            .Font.Bold = True
        End With
    Next i
    
    ' Fórmula
    .Range(.Cells(4, lastCol + 1), .Cells(lastRow, lastCol + 1)).FormulaR1C1 = "=VLOOKUP(RC[-5],R3C11:R825C13,3,0)"
    .Range(.Cells(4, lastCol + 2), .Cells(lastRow, lastCol + 2)).FormulaR1C1 = "=AND(RC[-1]>0,LEN(RC[-1])=9)"
    .Columns("A:M").AutoFit

End With


'-------------------------- CAMBIO FINAL: pone la fecha en la que se ejecuto la macro --------------------------
wbPlantillaRL1vsT4.Sheets("BOX FB1= BOX 27").Visible = xlSheetHidden
wbPlantillaRL1vsT4.Sheets("Audits").Range("M1").Value = Date
Sheets("Audits").Activate
wbPlantillaRL1vsT4.Save
wbPlantillaRL1vsT4.Close

End Sub

