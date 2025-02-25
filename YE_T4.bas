Attribute VB_Name = "YE_T4"
Option Explicit
'variables para todo el proyecto

Dim BASE, wbPlantillaT4, wbBasesPlantilla, wbBases, NewWorkbook As Workbook
Dim ws, wsBuscar As Worksheet
Dim Fecha1, Fecha2, Fecha3, mes, Mes_Texto, BPA, año, dia, ruta, Ruta_pais, Ruta_Año, Ruta_Audi, ruta_Base, texto, ExisteArchivo, T4File, T4AFile As String
Dim hoja As Integer
Dim lastRow, lastCol, i, h, buscarCol, Clmn, LR As Long
Dim titles As Variant
Dim buscarRango, rng As Range
Dim Quest As VbMsgBoxResult
    
Sub Ejecutar_YE_T4()

' Verificar si hay datos en las celdas I8 y M8
If ThisWorkbook.Sheets("Home Page").Range("I8").Value = "" Or ThisWorkbook.Sheets("Home Page").Range("E18").Value = "" Then
    MsgBox "Incomplete data, please enter the data before executing.", vbExclamation
    Exit Sub
End If

' Llama a cada una de las funciones
InicializarVariables
CrearCarpetas
Desarrollo_T4

MsgBox "T4 audits completed. Please access the document to perform the relevant reviews.", vbInformation

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

Sub Desarrollo_T4()

'-------------------------- CREAR ARCHIVO DE LAS BASES --------------------------

'Solcita al usuario abrir el doc de la base del RL1
MsgBox "Please select the T4 database file downloaded from sap", vbInformation
ruta_Base = Application.GetOpenFilename("Archivos Excel (*.xls; *.xlsx), *.xls; *.xlsx")
Application.AskToUpdateLinks = False
    
' Abre el reporte seleccionado por el usuario
If ruta_Base <> "Falso" Then
    
    'Verifica si el archivo de las bases existe o no
    ExisteArchivo = Dir(Ruta_Año & "\" & año & mes & " Archivo Bases Boxes Validation" & ".xlsx")

    If ExisteArchivo = "" Then 'Si no existe el archvio de las bases en la carpeta
    MsgBox "First run the RL1  macro audit before running this one", vbExclamation
    Exit Sub

    Else 'Si el documento de las bases ya existe

        Set BASE = Workbooks.Open(Filename:=ruta_Base, UpdateLinks:=0)
        Set wbBases = Workbooks.Open(Ruta_Año & "\" & año & mes & " Archivo Bases Boxes Validation" & ".xlsx")
        BASE.Activate
        Worksheets(1).Activate
        lastRow = Cells(Rows.Count, "A").End(xlUp).row
        Range("A1:P" & lastRow).Select
        Selection.Copy
        
        'Pega datos en la plantilla
        wbBases.Activate
        Sheets("ORIGINAL T4").Activate
        Sheets("ORIGINAL T4").Range("A1").PasteSpecial Paste:=xlPasteAll
        Columns("A:P").AutoFit
        Application.CutCopyMode = False
        wbBases.Save
        wbBases.Close
        BASE.Close
  
    End If
End If

'Realiza cambios de formato para el archivo de las bases
Set wbBases = Workbooks.Open(Ruta_Año & "\" & año & mes & " Archivo Bases Boxes Validation" & ".xlsx")
wbBases.Activate

'Filtra para que solo quede TTA
Sheets("ORIGINAL T4").Activate
lastRow = Cells(Rows.Count, "A").End(xlUp).row
Rows("1:1").AutoFilter
Rows("1:1").AutoFilter Field:=14, Criteria1:="=#TTA"
Range("A1:P" & lastRow).SpecialCells(xlCellTypeVisible).Copy
Sheets("BASE T4").Activate
Range("A1").PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
Application.CutCopyMode = False

'Eliminala hoja original
Application.DisplayAlerts = False
On Error Resume Next
Sheets("ORIGINAL T4").Delete
On Error GoTo 0 '
Application.DisplayAlerts = True
wbBases.Save

'Poner todos los boxes para que el rango no cambie
lastRow = Sheets("BASE T4").Cells(Sheets("BASE T4").Rows.Count, 1).End(xlUp).row
ThisWorkbook.Sheets("Anexxes").Range("H2:H25").Copy
wbBases.Sheets("BASE T4").Cells(lastRow + 1, "L").PasteSpecial Paste:=xlPasteValues
Application.CutCopyMode = False
ThisWorkbook.Sheets("Anexxes").Range("I2:I25").Copy
wbBases.Sheets("BASE T4").Cells(lastRow + 1, "A").PasteSpecial Paste:=xlPasteValues
Application.CutCopyMode = False

'Crea tabla dinamica
Sheets("BASE T4").Activate

    Dim ult_Tabla As Long
    ult_Tabla = Cells(Rows.Count, "A").End(xlUp).row
    Dim rangoTabla1 As Range
    Set rangoTabla1 = Sheets("BASE T4").Range("A1:P" & ult_Tabla)
    ActiveSheet.ListObjects.Add(xlSrcRange, rangoTabla1, , xlYes).Name = "Tabla1"
    
    'Crear tabla dinamica
    Dim celdaTablaDinamica1 As Range
    Set celdaTablaDinamica1 = Sheets("TD T4").Range("A1")
    Dim tablaDinamica1 As PivotTable
    
    'Activa campos y le pone formato tabular
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        rangoTabla1, Version:=6).CreatePivotTable TableDestination:= _
        celdaTablaDinamica1, TableName:="tablaDinamica1", DefaultVersion:=6
        
    Sheets("TD T4").Select
    With ActiveSheet.PivotTables("tablaDinamica1").PivotFields("PERNR")
        .Orientation = xlRowField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("tablaDinamica1").PivotFields("BUSNM")
        .Orientation = xlRowField
        .Position = 2
    End With
    With ActiveSheet.PivotTables("tablaDinamica1").PivotFields("WRKAR")
        .Orientation = xlRowField
        .Position = 3
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
    ActiveSheet.PivotTables("tablaDinamica1").RowAxisLayout xlTabularRow
    ActiveSheet.PivotTables("tablaDinamica1").RepeatAllLabels xlRepeatLabels

wbBases.Save

'-------------------------- CREAR ARCHIVO DE LA PLANTILLA DEL T4 --------------------------

'Hace copia de la plantilla del T4
On Error Resume Next
Kill Ruta_Año & "\" & año & mes & " T4 Audits" & ".xlsx"
On Error GoTo 0

Set wbPlantillaT4 = Workbooks.Open(ruta & "\" & "T4 Audits.xlsx")
wbPlantillaT4.Activate
ActiveWorkbook.SaveCopyAs Filename:=Ruta_Año & "\" & año & mes & " T4 Audits" & ".xlsx"
wbPlantillaT4.Close False

'Abre los archivos correspondientes
Set wbPlantillaT4 = Workbooks.Open(Ruta_Año & "\" & año & mes & " T4 Audits" & ".xlsx")

'Pasa la tabla dinamica como valores al documento del T4
wbBases.Activate
wbBases.Sheets("TD T4").Activate
lastRow = Cells(Rows.Count, "A").End(xlUp).row
Range("A2:AE" & lastRow - 1).Select
Range("A2:AE" & lastRow - 1).Copy
wbPlantillaT4.Activate
Sheets("TD T4").Activate
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

'Formula del concatenar el KEY NUMBER
Range("A1").Value = "KEY NUMBER"
lastRow = Cells(Rows.Count, "B").End(xlUp).row
Sheets("TD T4").Range("A2:A" & lastRow) = "=+CONCATENATE(RC[1],RC[2],RC[3])"
Columns("A:AE").AutoFit

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


'Pasa la tabla dinamica como valores al documento del RL1
wbBases.Activate
wbBases.Sheets("TD RL1").Activate
lastRow = Cells(Rows.Count, "A").End(xlUp).row
Range("A2:AE" & lastRow - 1).Select
Range("A2:AE" & lastRow - 1).Copy
wbPlantillaT4.Activate
Sheets("TD RL1").Activate
Range("C1").Select
Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False
Application.CutCopyMode = False
With Range("A1:AB1")
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

lastRow = Cells(Rows.Count, "A").End(xlUp).row
For i = lastRow To 2 Step -1
        If Cells(i, "B").Value = "999999" Then
            Rows(i).Delete
        End If
Next i

'Formula del concatenar el KEY NUMBER
Range("A1").Value = "KEY NUMBER #2"
lastRow = Cells(Rows.Count, "C").End(xlUp).row
Sheets("TD RL1").Range("A2:A" & lastRow) = "=CONCATENATE(RC[2],RC[4])"
Columns("A:AD").AutoFit

'-------------------------- Ciclo para pegar las 4 columnas principales --------------------------
'NOTA:
    'Hoja 4 - BOX 14 >= BOX 26 UP TO CPP MAX
    'Hoja 5 - NATIVE EES
    'Hoja 6 - BOX 22 <= BOX 14 (EXCEPT FOR SEVERANCE PAYMENTS)
    'Hoja 7 - B14 >= 0 AND BOX 22 = BLANK (SHOULD BE BLANK ONLY FOR EXEMPT EMPLOYEES IN TD1 (FIT)), CHECK IF THEY HAVE BENEFITS ACTIVE
    'Hoja 8 - BOX 14 CAN NOT BE LESS THAN BOX 30 + BOX 34 + BOX 40
    'Hoja 9 - BOX 16A + BOX 16 = BOX 27
    'Hoja 10 - BOX 24 <= BOX 14
    'Hoja 11 - BOX 28
    'Hoja 12 - BOX 50 SHOULD HAVE 7 DIGITS (AFTER PA ENTRIES)
    'Hoja 13 - BOX 24 AND BOX 26 >= 0, SHOULD NOT BE BLANK
    'Hoja 14 - BOX 20 >0 SHOULD HAVE BOX 52
    'Hoja 15 - BOX 45  SHOULD NOT BE IN BLANK OR 0 FOR T4 SILP IN XML FILE
    'Hoja 16 - BOX 015 AND SHOULD NOT BE IN BLANK OR 0 FOR T4A SLIP IN XML FILE

For hoja = 4 To 17

    'Seleccina los datos que va a copiar
    wbPlantillaT4.Sheets("TD T4").Activate
    lastRow = Cells(Rows.Count, "B").End(xlUp).row
    Range("B1:D" & lastRow).Select
    Range("B1:D" & lastRow).Copy
    
    'Pega datos en cada hoja de los reportes
    Set ws = wbPlantillaT4.Sheets(hoja)
    ws.Activate
    Range("A3").Select
    Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    
    'Seleccina los datos que va a copiar
    wbPlantillaT4.Sheets("TD T4").Activate
    lastRow = Cells(Rows.Count, "B").End(xlUp).row
    Range("A1:A" & lastRow).Select
    Range("A1:A" & lastRow).Copy
    
    'Pega datos en cada hoja de los reportes
    Set ws = wbPlantillaT4.Sheets(hoja)
    ws.Activate
    Range("D3").Select
    Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    
    'Formato de color de los campos de la tabla
    With Range("A3:D3")
        .Font.Bold = True
        .Interior.Color = RGB(9, 61, 147)
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
    End With
    Columns("A:D").AutoFit
    
Next hoja

wbBases.Save
wbBases.Close
wbPlantillaT4.Save


'-------------------------- REPARTE LAS COLUMNAS DE LOS BOXES PARA LAS AUDIRORIAS --------------------------

Dim wsTDT4, wsDestino As Worksheet
Dim col As Range
Dim destCol As Long
Dim Posicion_hojas, criterios As Variant

'Setear la hoja de donde toma las columnas
Set wsTDT4 = wbPlantillaT4.Sheets("TD T4")

'Definir las hojas y los criterios, el orden de los criterios esta segun el orden de las hojas en el archivo
'Nota:
    'Hoja 4 - BOX 14 >= BOX 26 UP TO CPP MAX
    'Hoja 5 - NATIVE EES
    'Hoja 6 - BOX 22 <= BOX 14 (EXCEPT FOR SEVERANCE PAYMENTS)
    'Hoja 7 - B14 >= 0 AND BOX 22 = BLANK (SHOULD BE BLANK ONLY FOR EXEMPT EMPLOYEES IN TD1 (FIT)), CHECK IF THEY HAVE BENEFITS ACTIVE
    'Hoja 8 - BOX 14 CAN NOT BE LESS THAN BOX 30 + BOX 34 + BOX 40
    'Hoja 9 - BOX 16A + BOX 16 = BOX 27
    'Hoja 10 - BOX 24 <= BOX 14
    'Hoja 11 - BOX 28
    'Hoja 12 - BOX 50 SHOULD HAVE 7 DIGITS (AFTER PA ENTRIES)
    'Hoja 13 - BOX 24 AND BOX 26 >= 0, SHOULD NOT BE BLANK
    'Hoja 14 - BOX 20 >0 SHOULD HAVE BOX 52
    'Hoja 15 - BOX 45  SHOULD NOT BE IN BLANK OR 0 FOR T4 SILP IN XML FILE
    'Hoja 16 - BOX 015 AND SHOULD NOT BE IN BLANK OR 0 FOR T4A SLIP IN XML FILE
    'Hoja 17 - BOX 19 = BOX 18 APPLICABLE RATE
    
Posicion_hojas = Array(4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17) 'Las hojas en las que va a poner las columas

'Criterios para cada hoja. IMPORTANE: cada renglon es el crioterio de la hoja por posicion
criterios = Array( _
    Array("B14", "B26", "F71"), _
    Array("B14", "B22", "F71"), _
    Array("B14", "B22", "F66", "F67"), _
    Array("B14", "B22", "F40"), _
    Array("B14", "F30", "F34", "F40"), _
    Array("B16A", "B16", "B27", "B27A"), _
    Array("B14", "B24", "F71"), _
    Array("B28"), _
    Array("B50"), _
    Array("B24", "B26"), _
    Array("B20", "F52"), _
    Array("B45"), _
    Array("B15"), _
    Array("B18", "B19") _
)

'Iterar sobre las hojas
For i = LBound(Posicion_hojas) To UBound(Posicion_hojas)

    'Asignar la hoja de destino
    Set wsDestino = wbPlantillaT4.Sheets(Posicion_hojas(i))
    lastRow = wsDestino.Cells(wsDestino.Rows.Count, 1).End(xlUp).row
    
    'Indica que inicie a pegar los datos desde la columna E
    destCol = 5
    
    'Realiza la busqueda
    For Each col In wsTDT4.Rows(1).Cells
        
        If Not IsError(Application.Match(col.Value, criterios(i), 0)) Then
            wsTDT4.Range(col, wsTDT4.Cells(lastRow, col.Column)).Copy
            wsDestino.Cells(3, destCol).PasteSpecial Paste:=xlPasteValues
            destCol = destCol + 1
        End If
    Next col

    Application.CutCopyMode = False
    
Next i

'Formatos de las celdas de los titulos
For hoja = 4 To 17

    Set ws = wbPlantillaT4.Sheets(hoja)
    ws.Activate
    lastCol = Cells(3, Columns.Count).End(xlToLeft).Column
    With Range(Cells(3, 5), Cells(3, lastCol))
        .Interior.Color = RGB(243, 156, 18)
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
    End With
    Columns("A:I").AutoFit
    
Next hoja

wbPlantillaT4.Save

'-------------------------- CREA LAS FORMULAS DE LAS AUDITORIAS --------------------------


'Hoja 4 - BOX 14 >= BOX 26 UP TO CPP MAX

With Sheets(4)
    .Activate
    lastCol = .Cells(3, .Columns.Count).End(xlToLeft).Column
    lastRow = .Cells(.Rows.Count, 1).End(xlUp).row

    ' Títulos y configuración
    titles = Array("BOXJ", "AUDIT 1", "DIFF AUDIT 1", "DIFF AUDIT 2", "AUDIT 2", "COMMENTS")
    
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
    .Range(.Cells(4, lastCol + 2), .Cells(lastRow, lastCol + 2)).FormulaR1C1 = "=RC[-4]>=RC[-3]"
    .Range(.Cells(4, lastCol + 3), .Cells(lastRow, lastCol + 3)).FormulaR1C1 = "=IFERROR(IF(RC[-1]=FALSE,RC[-4]-RC[-5],"" ""),"" "")"
    .Range(.Cells(4, lastCol + 4), .Cells(lastRow, lastCol + 4)).FormulaR1C1 = "=IFERROR(ROUND(IF(RC[-2]=FALSE,RC[-3]-RC[-1],"" ""),10),"" "")"
    .Range(.Cells(4, lastCol + 5), .Cells(lastRow, lastCol + 5)).FormulaR1C1 = "=IF(AND(RC[-3]=FALSE, RC[-1]<>0),IF(AND(RC[-7]=0,RC[-5]>0),""OK-NATIVE EMPLOYEE"",""REVIEW""),"" "")"
    .Columns("A:M").AutoFit
End With

    'Hace el buscar del Box J
    wbPlantillaT4.Sheets(4).Activate
    lastRow = Cells(Rows.Count, "B").End(xlUp).row
    Set wsBuscar = wbPlantillaT4.Sheets("TD RL1")
    
    On Error Resume Next
    buscarCol = Application.WorksheetFunction.Match("BOX J", wsBuscar.Rows(1), 0)
    On Error GoTo 0
    
    'Hace el buscar
    Set buscarRango = wsBuscar.Range(wsBuscar.Cells(1, 1), wsBuscar.Cells(lastRow, buscarCol))
    h = 4
    Sheets(4).Range("H4:H" & lastRow) = "=IFERROR(VLOOKUP(CONCATENATE(A" & h & ",C" & h & "),'" & wsBuscar.Name & "'!" & buscarRango.Address & "," & buscarCol & ",0),"" "")"

'Hoja 5 - NATIVE EES

With Sheets(5)
    .Activate
    lastCol = .Cells(3, .Columns.Count).End(xlToLeft).Column
    lastRow = .Cells(.Rows.Count, 1).End(xlUp).row

    ' Títulos y configuración
    titles = Array("AUDIT", "COMMENTS")
    
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
    .Range(.Cells(4, lastCol + 1), .Cells(lastRow, lastCol + 1)).FormulaR1C1 = "=IF(AND(RC[-3]=0,RC[-2]="""",RC[-1]>0), ""OK NATIVE EE"", "" "")"
    .Columns("A:M").AutoFit
End With

'Hoja 6 - BOX 22 <= BOX 14 (EXCEPT FOR SEVERANCE PAYMENTS)

With Sheets(6)
    .Activate
    lastCol = .Cells(3, .Columns.Count).End(xlToLeft).Column
    lastRow = .Cells(.Rows.Count, 1).End(xlUp).row

    ' Títulos y configuración
    titles = Array("AUDIT 1", "AUDIT 2", "COMMENTS")
    
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
    .Range(.Cells(4, lastCol + 1), .Cells(lastRow, lastCol + 1)).FormulaR1C1 = "=RC[-3]<=RC[-4]"
    .Range(.Cells(4, lastCol + 2), .Cells(lastRow, lastCol + 2)).FormulaR1C1 = "=IF(RC[-1]=FALSE,IF((RC[-3]+RC[-2])>0,""Ok - Severance Payment"",""Review""),"" "")"
    .Columns("A:M").AutoFit
End With

'Hoja 7 - B14 >= 0 AND BOX 22 = BLANK (SHOULD BE BLANK ONLY FOR EXEMPT EMPLOYEES IN TD1 (FIT)), CHECK IF THEY HAVE BENEFITS ACTIVE

With Sheets(7)
    .Activate
    .Range("E2").Value = BPA
    .Range("D2").Value = "BPA"
    lastCol = .Cells(3, .Columns.Count).End(xlToLeft).Column
    lastRow = .Cells(.Rows.Count, 1).End(xlUp).row

    ' Títulos y configuración
    titles = Array("AUDIT 1", "AUDIT 2", "AUDIT 3", "COMMENTS")
    
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
    .Range(.Cells(4, lastCol + 1), .Cells(lastRow, lastCol + 1)).FormulaR1C1 = "=IF(AND(RC[-3]>0,RC[-2]=0),""VERDADERO -REVIE"",""FALSO-OK"")"
    .Range(.Cells(4, lastCol + 2), .Cells(lastRow, lastCol + 2)).FormulaR1C1 = "=IF(RC[-1]=""FALSO-OK"","" "",IF(RC[-4]<R2C5,""OK- Box14 lower than BPA"",""REVIEW""))"
    .Range(.Cells(4, lastCol + 3), .Cells(lastRow, lastCol + 3)).FormulaR1C1 = "=IF(RC[-2]=""FALSO-OK"","" "",IF(RC[-5]-RC[-3]=0,""Income related to TB only"",""Review""))"
    .Columns("A:M").AutoFit
    
End With

'Hoja 8 - BOX 14 CAN NOT BE LESS THAN BOX 30 + BOX 34 + BOX 40

With Sheets(8)
    .Activate
    lastCol = .Cells(3, .Columns.Count).End(xlToLeft).Column
    lastRow = .Cells(.Rows.Count, 1).End(xlUp).row

    ' Títulos y configuración
    titles = Array("AUDIT", "COMMENTS")
    
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
    .Range(.Cells(4, lastCol + 1), .Cells(lastRow, lastCol + 1)).FormulaR1C1 = "=RC[-4]>=(RC[-3]+RC[-2]+RC[-1])"
    .Columns("A:M").AutoFit
End With

'Hoja 9 - BOX 16A + BOX 16 = BOX 27 + BOX 27A

With Sheets(9)
    .Activate
    lastCol = .Cells(3, .Columns.Count).End(xlToLeft).Column
    lastRow = .Cells(.Rows.Count, 1).End(xlUp).row

    ' Títulos y configuración
    titles = Array("AUDIT", "COMMENTS")
    
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
    .Range(.Cells(4, lastCol + 1), .Cells(lastRow, lastCol + 1)).FormulaR1C1 = "=(RC[-4]+RC[-3])-RC[-2]-RC[-1]"
    .Columns("A:N").AutoFit
End With


'Hoja 10 - BOX 24 <= BOX 14

With Sheets(10)
    .Activate
    lastCol = .Cells(3, .Columns.Count).End(xlToLeft).Column
    lastRow = .Cells(.Rows.Count, 1).End(xlUp).row

    ' Títulos y configuración
    titles = Array("AUDIT 1", "AUDIT 2", "COMMENTS")
    
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
    .Range(.Cells(4, lastCol + 1), .Cells(lastRow, lastCol + 1)).FormulaR1C1 = "=RC[-2]<=RC[-3]"
    .Range(.Cells(4, lastCol + 2), .Cells(lastRow, lastCol + 2)).FormulaR1C1 = "=IF(RC[-1]=TRUE,"" "",IF(AND(RC[-4]=0,RC[-2]>0),""Ok- Native employee"",""Review""))"
    .Columns("A:M").AutoFit
End With

''Hoja 11 - BOX 28 --- AUDITORIA PENDIENTE
'
'With Sheets(11)
'    .Activate
'    lastCol = .Cells(3, .Columns.Count).End(xlToLeft).Column
'    lastRow = .Cells(.Rows.Count, 1).End(xlUp).row
'
'    ' Títulos y configuración
'    titles = Array("AUDIT 1", "AUDIT 2"",COMMENTS")
'
'    For i = 0 To UBound(titles)
'        With .Cells(3, lastCol + 1 + i)
'            .Value = titles(i)
'            If InStr(1, titles(i), "AUDIT", vbTextCompare) > 0 Then
'                .Interior.Color = RGB(146, 208, 80)
'            Else
'                .Interior.Color = RGB(9, 61, 147)
'            End If
'            .Font.Color = RGB(255, 255, 255)
'            .HorizontalAlignment = xlCenter
'            .Font.Bold = True
'        End With
'    Next i
'
'    ' Fórmula
'    .Range(.Cells(4, lastCol + 1), .Cells(lastRow, lastCol + 1)).FormulaR1C1 = "=IF(COUNTIF(R4C1:R" & lastRow & "C1,RC[-4])>=2,""REVIEW"",""OK"")"
'    .Range(.Cells(4, lastCol + 2), .Cells(lastRow, lastCol + 2)).FormulaR1C1 = "=IF(COUNTIF(R4C1:R" & lastRow & "C1,RC[-4])>=2,""REVIEW"",""OK"")"
'    .Columns("A:M").AutoFit
'End With
'
'
''Hoja 12 - BOX 50 SHOULD HAVE 7 DIGITS (AFTER PA ENTRIES) --- AUDITORIA PENDIENTE
'
'With Sheets(12)
'    .Activate
'    lastCol = .Cells(3, .Columns.Count).End(xlToLeft).Column
'    lastRow = .Cells(.Rows.Count, 1).End(xlUp).row
'
'    ' Títulos y configuración
'    titles = Array("AUDIT 1", "AUDIT 2"",COMMENTS")
'
'    For i = 0 To UBound(titles)
'        With .Cells(3, lastCol + 1 + i)
'            .Value = titles(i)
'            If InStr(1, titles(i), "AUDIT", vbTextCompare) > 0 Then
'                .Interior.Color = RGB(146, 208, 80)
'            Else
'                .Interior.Color = RGB(9, 61, 147)
'            End If
'            .Font.Color = RGB(255, 255, 255)
'            .HorizontalAlignment = xlCenter
'            .Font.Bold = True
'        End With
'    Next i
'
'    ' Fórmula
'    .Range(.Cells(4, lastCol + 1), .Cells(lastRow, lastCol + 1)).FormulaR1C1 = "=IF(COUNTIF(R4C1:R" & lastRow & "C1,RC[-4])>=2,""REVIEW"",""OK"")"
'    .Range(.Cells(4, lastCol + 2), .Cells(lastRow, lastCol + 2)).FormulaR1C1 = "=IF(COUNTIF(R4C1:R" & lastRow & "C1,RC[-4])>=2,""REVIEW"",""OK"")"
'    .Columns("A:M").AutoFit
'End With

'Hoja 13 - BOX 24 AND BOX 26 >= 0, SHOULD NOT BE BLANK

With Sheets(13)
    .Activate
    lastCol = .Cells(3, .Columns.Count).End(xlToLeft).Column
    lastRow = .Cells(.Rows.Count, 1).End(xlUp).row

    ' Títulos y configuración
    titles = Array("AUDIT", "COMMENTS")
    
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
    .Range(.Cells(4, lastCol + 1), .Cells(lastRow, lastCol + 1)).FormulaR1C1 = "=AND(RC[-2],RC[-1])>=0"
    .Columns("A:M").AutoFit
End With

'Hoja 14 - BOX 20 >0 SHOULD HAVE BOX 52

With Sheets(14)
    .Activate
    lastCol = .Cells(3, .Columns.Count).End(xlToLeft).Column
    lastRow = .Cells(.Rows.Count, 1).End(xlUp).row

    ' Títulos y configuración
    titles = Array("AUDIT", "COMMENTS")
    
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
    .Range(.Cells(4, lastCol + 1), .Cells(lastRow, lastCol + 1)).FormulaR1C1 = "=IF(AND(RC[-2]>0,RC[-1]>0),""OK"",""REVIEW"")"
    .Columns("A:M").AutoFit
End With

'Hoja 15 - BOX 45  SHOULD NOT BE IN BLANK OR 0 FOR T4 SILP IN XML FILE --- AUDITORIA PENDIENTE

With Sheets(15)
    .Activate
    .Range("E4").Value = "audit in separate file Box 45_Box 015 Audit"
    .Range("E4").HorizontalAlignment = xlCenter
    .Range("E4").Font.Bold = True
    .Columns("A:M").AutoFit
End With

'Hoja 16 - BOX 015 AND SHOULD NOT BE IN BLANK OR 0 FOR T4A SLIP IN XML FILE --- AUDITORIA PENDIENTE

With Sheets(16)
    .Activate
    .Range("E4").Value = "audit in separate file Box 45_Box 015 Audit"
    .Range("E4").HorizontalAlignment = xlCenter
    .Range("E4").Font.Bold = True
    .Columns("A:M").AutoFit
End With

'Hoja 17 - BOX 19 = BOX 18 APPLICABLE RATE

With Sheets(17)
    .Activate
    lastCol = .Cells(3, .Columns.Count).End(xlToLeft).Column
    lastCol = lastCol - 13
    lastRow = .Cells(.Rows.Count, 1).End(xlUp).row

    ' Títulos y configuración
    titles = Array("RATE", "BOX18*RATE", "AUDIT 1", "AUDIT 2", "COMMENTS")
    
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
    .Range(.Cells(4, lastCol + 1), .Cells(lastRow, lastCol + 1)).FormulaR1C1 = "=VLOOKUP(RC[-5],R3C16:R15C18,3,0)"
    .Range(.Cells(4, lastCol + 2), .Cells(lastRow, lastCol + 2)).FormulaR1C1 = "=RC[-3]*RC[-1]"
    .Range(.Cells(4, lastCol + 3), .Cells(lastRow, lastCol + 3)).FormulaR1C1 = "=RC[-3]-RC[-1]"
    .Range(.Cells(4, lastCol + 4), .Cells(lastRow, lastCol + 4)).FormulaR1C1 = "=RC[-4]/RC[-5]"
    
    .Columns("A:M").AutoFit
End With

'-------------------------- CAMBIO FINAL: pone la fecha en la que se ejecuto la macro --------------------------
wbPlantillaT4.Sheets("AUDIT LIST").Range("M1").Value = Date
Sheets("AUDIT LIST").Activate
wbPlantillaT4.Save
wbPlantillaT4.Close

End Sub


Sub Ejecutar_YE_Y4_45_50()

'Revisa las variables y crea las carpetas si no estan creadas
InicializarVariables
CrearCarpetas

On Error Resume Next
Kill Ruta_Año & "\" & año & mes & " Box 45_Box 015 Audit.xlsx"
On Error GoTo 0

'T4 Box 45 Audit
Quest = MsgBox("Would you like to run the Audit for 'BOX 45 SHOULD NOT BE IN BLANK OR 0 FOR T4 SLIP IN XML FILE'", vbQuestion + vbYesNo + vbDefaultButton2, "T4 BOX 45 AUDIT")
If (Quest = vbYes) Then 'If the user selects Yes, this will do the whole audit

    ' Crea un nuevo libro de trabajo
    Set NewWorkbook = Workbooks.Add
    NewWorkbook.Sheets(1).Name = "T4 Box 45 Audit" ' Cambia el nombre de la hoja a T4
    T4File = Application.GetOpenFilename ' Pregunta al usuario por el archivo necesario
    
    ' Importa el archivo en el nuevo libro de trabajo
    NewWorkbook.XmlImport URL:=T4File, ImportMap:=Nothing, Overwrite:=True, Destination:=NewWorkbook.Sheets(1).Range("$A$1") ' Imports file in Excel
    
    NewWorkbook.Sheets(1).Cells.Find(What:="empr_dntl_ben_rpt_cd", After:=ActiveCell, LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False, SearchFormat:=False).Select ' Busca la columna que se necesita
    Clmn = Selection.Cells(1, 1).Column ' Guarda en una variable la columna necesaria para filtrar por 0
    NewWorkbook.Sheets(1).UsedRange.AutoFilter Field:=Clmn, Criteria1:="0" ' Filtra el rango
    Set rng = NewWorkbook.Sheets(1).AutoFilter.Range.Offset(1).SpecialCells(xlCellTypeVisible) ' Establece el rango para colorear
    rng.Interior.Color = RGB(255, 204, 204) ' Establece el nuevo color para la fila "0"
    NewWorkbook.Sheets(1).ShowAllData ' Quita el filtro
    LR = NewWorkbook.Sheets(1).UsedRange.Rows.Count ' Cuenta las filas
    NewWorkbook.Sheets(1).Rows(LR).EntireRow.Delete ' Elimina la última fila
    
    ' Guarda el nuevo libro en la ruta especificada
    NewWorkbook.SaveAs Filename:=Ruta_Año & "\" & año & mes & " Box 45_Box 015 Audit.xlsx", FileFormat:=xlOpenXMLWorkbook ' Guardar como archivo Excel
End If

    
'T4A Box 015 Audit
Quest = MsgBox("Would you like to run the Audit for 'BOX 015 SHOULD NOT BE IN BLANK OR 0 FOR T4A SLIP IN XML FILE'", vbQuestion + vbYesNo + vbDefaultButton2, "T4A BOX 45 AUDIT")
If (Quest = vbYes) Then 'If the user selects Yes, this will do the whole audit

    Sheets.Add(After:=Sheets("T4 Box 45 Audit")).Name = "T4A Box 015 Audit" 'Changes the sheet name
    T4AFile = Application.GetOpenFilename 'Asks user to select file needed
    NewWorkbook.XmlImport URL:=T4AFile, ImportMap:=Nothing, Overwrite:=True, Destination:=NewWorkbook.Sheets(2).Range("$A$1") ' Imports file in Excel
    NewWorkbook.Sheets(2).Cells.Find(What:="payr_dntl_ben_rpt_cd", After:=ActiveCell, LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False, SearchFormat:=False).Select 'Finds column to look for the zero, "payr_dntl_ben_rpt_cd" for T4A
    Clmn = Selection.Cells(1, 1).Column 'Saves in a variable the column field needed to filter by 0
    NewWorkbook.Sheets(2).UsedRange.AutoFilter Field:=Clmn, Criteria1:="0"  'Filters range
    Set rng = ActiveSheet.AutoFilter.Range.Offset(1).SpecialCells(xlCellTypeVisible) 'Sets the range to color
    rng.Interior.Color = RGB(255, 204, 204) 'Sets new color to "0" row
    ActiveSheet.ShowAllData 'Unfilters
    LR = ActiveSheet.UsedRange.Rows.Count 'Counts rows
    Rows(LR).EntireRow.Delete 'Deletes row
    
End If

NewWorkbook.Sheets(1).Activate
NewWorkbook.Save
NewWorkbook.Close
MsgBox "T4 Box 45 & T4A Box 015 audits completed. Please access the document to perform the relevant reviews.", vbInformation

End Sub
