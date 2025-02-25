Attribute VB_Name = "YE_RL1"
Option Explicit
'variables para todo el proyecto

Dim BASE, wbPlantillaRL1, wbBasesPlantilla, wbBases As Workbook
Dim ws As Worksheet
Dim Fecha1, Fecha2, Fecha3, mes, Mes_Texto, BPA, año, dia, ruta, Ruta_pais, Ruta_Año, Ruta_Audi, ruta_Base, texto, ExisteArchivo As String
Dim hoja As Integer
Dim lastRow, lastCol, i As Long
Dim titles As Variant
    
Sub Ejecutar_YE_RL1()

' Verificar si hay datos en las celdas I8 y M8
If ThisWorkbook.Sheets("Home Page").Range("I8").Value = "" Or ThisWorkbook.Sheets("Home Page").Range("E18").Value = "" Then
    MsgBox "Incomplete data, please enter the data before executing.", vbExclamation
    Exit Sub
End If

' Llama a cada una de las funciones
InicializarVariables
CrearCarpetas
Desarrollo_RL1

MsgBox "RL1 audits completed. Please access the document to perform the relevant reviews.", vbInformation

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

Sub Desarrollo_RL1()

'-------------------------- CREAR ARCHIVO DE LAS BASES --------------------------

'Solcita al usuario abrir el doc de la base del RL1
MsgBox "Please select the RL1 database file downloaded from sap", vbInformation
ruta_Base = Application.GetOpenFilename("Archivos Excel (*.xls; *.xlsx), *.xls; *.xlsx")
Application.AskToUpdateLinks = False
    
' Abre el reporte seleccionado por el usuario
If ruta_Base <> "Falso" Then
    
    'Verifica si el archivo de las bases existe o no
    ExisteArchivo = Dir(Ruta_Año & "\" & año & mes & " Archivo Bases Boxes Validation" & ".xlsx")

    If ExisteArchivo = "" Then 'Si no existe el archvio de las bases en la carpeta
   
        Set BASE = Workbooks.Open(Filename:=ruta_Base, UpdateLinks:=0)
        BASE.Activate
        Worksheets(1).Activate
        lastRow = Cells(Rows.Count, "A").End(xlUp).row
        Range("A1:P" & lastRow).Select
        Selection.Copy
        
        'Crea el archivo nuevo
        Workbooks.Add
        ActiveSheet.Range("A1").PasteSpecial Paste:=xlPasteAll
        Columns("A:P").AutoFit
        Application.CutCopyMode = False
        
        'Guarda el archivo con el nombre nuevo
        ActiveWorkbook.SaveAs Ruta_Año & "\" & año & mes & " Archivo Bases Boxes Validation" & ".xlsx"
        Workbooks(año & mes & " Archivo Bases Boxes Validation" & ".xlsx").Save
        Sheets(1).Activate
        ActiveSheet.Name = "ORIGINAL RL1"
        Sheets.Add(After:=Sheets(Sheets.Count)).Name = "BASE RL1"
        Sheets.Add(After:=Sheets(Sheets.Count)).Name = "TD RL1"
        Sheets.Add(After:=Sheets(Sheets.Count)).Name = "ORIGINAL T4"
        Sheets.Add(After:=Sheets(Sheets.Count)).Name = "BASE T4"
        Sheets.Add(After:=Sheets(Sheets.Count)).Name = "TD T4"
        Sheets.Add(After:=Sheets(Sheets.Count)).Name = "TD T4 WITHOUT BN"
        Sheets.Add(After:=Sheets(Sheets.Count)).Name = "TD T4 ONLY QC"
        Workbooks(año & mes & " Archivo Bases Boxes Validation" & ".xlsx").Save
        Workbooks(año & mes & " Archivo Bases Boxes Validation" & ".xlsx").Close
        BASE.Close
        
    
    Else 'Si el documento de las bases ya existe
        On Error Resume Next
        Kill Ruta_Año & "\" & año & mes & " Archivo Bases Boxes Validation" & ".xlsx"
        On Error GoTo 0
        
        Set BASE = Workbooks.Open(Filename:=ruta_Base, UpdateLinks:=0)
        BASE.Activate
        Worksheets(1).Activate
        lastRow = Cells(Rows.Count, "A").End(xlUp).row
        Range("A1:P" & lastRow).Select
        Selection.Copy
        
        'Crea el archivo nuevo
        Workbooks.Add
        ActiveSheet.Range("A1").PasteSpecial Paste:=xlPasteAll
        Columns("A:P").AutoFit
        Application.CutCopyMode = False
        
        'Guarda el archivo con el nombre nuevo
        ActiveWorkbook.SaveAs Ruta_Año & "\" & año & mes & " Archivo Bases Boxes Validation" & ".xlsx"
        Workbooks(año & mes & " Archivo Bases Boxes Validation" & ".xlsx").Save
        Sheets(1).Activate
        ActiveSheet.Name = "ORIGINAL RL1"
        Sheets.Add(After:=Sheets(Sheets.Count)).Name = "BASE RL1"
        Sheets.Add(After:=Sheets(Sheets.Count)).Name = "TD RL1"
        Sheets.Add(After:=Sheets(Sheets.Count)).Name = "ORIGINAL T4"
        Sheets.Add(After:=Sheets(Sheets.Count)).Name = "BASE T4"
        Sheets.Add(After:=Sheets(Sheets.Count)).Name = "TD T4"
        Workbooks(año & mes & " Archivo Bases Boxes Validation" & ".xlsx").Save
        Workbooks(año & mes & " Archivo Bases Boxes Validation" & ".xlsx").Close
        BASE.Close
  
    End If
End If

'Realiza cambios de formato para el archivo de las bases
Set wbBases = Workbooks.Open(Ruta_Año & "\" & año & mes & " Archivo Bases Boxes Validation" & ".xlsx")
wbBases.Activate

'Filtra para que solo quede TTA
Sheets("ORIGINAL RL1").Activate
lastRow = Cells(Rows.Count, "A").End(xlUp).row
Rows("1:1").AutoFilter
Rows("1:1").AutoFilter Field:=14, Criteria1:="=#TTA"
Range("A1:P" & lastRow).SpecialCells(xlCellTypeVisible).Copy
Sheets("BASE RL1").Activate
Range("A1").PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
Application.CutCopyMode = False

'Eliminala hoja original
Application.DisplayAlerts = False
On Error Resume Next
Sheets("ORIGINAL RL1").Delete
On Error GoTo 0 '
Application.DisplayAlerts = True
wbBases.Save

'Codigo para que cambie lo de las B
Sheets("BASE RL1").Activate
Columns("M:M").Insert Shift:=xlToRight
Range("M1").Value = "BOX NAME"

' Recorrer cada fila desde la 2 hasta la última
lastRow = Cells(Rows.Count, "A").End(xlUp).row
For i = 2 To lastRow
    texto = Sheets("BASE RL1").Cells(i, "L").Value
    Select Case texto
        Case "B01": Sheets("BASE RL1").Cells(i, "M").Value = "Box A"
        Case "B02": Sheets("BASE RL1").Cells(i, "M").Value = "Box B"
        Case "B02S": Sheets("BASE RL1").Cells(i, "M").Value = "Box Bs"
        Case "B02B": Sheets("BASE RL1").Cells(i, "M").Value = "Box BB"
        Case "B03": Sheets("BASE RL1").Cells(i, "M").Value = "Box C"
        Case "B04": Sheets("BASE RL1").Cells(i, "M").Value = "Box D"
        Case "B05": Sheets("BASE RL1").Cells(i, "M").Value = "Box E"
        Case "B06": Sheets("BASE RL1").Cells(i, "M").Value = "Box F"
        Case "B07": Sheets("BASE RL1").Cells(i, "M").Value = "Box G"
        Case "B22": Sheets("BASE RL1").Cells(i, "M").Value = "Box H"
        Case "B23": Sheets("BASE RL1").Cells(i, "M").Value = "Box I"
        Case "B10": Sheets("BASE RL1").Cells(i, "M").Value = "Box J"
        Case "B11": Sheets("BASE RL1").Cells(i, "M").Value = "Box K"
        Case "B12": Sheets("BASE RL1").Cells(i, "M").Value = "Box L"
        Case "B13": Sheets("BASE RL1").Cells(i, "M").Value = "Box M"
        Case "B14": Sheets("BASE RL1").Cells(i, "M").Value = "Box N"
        Case "B15": Sheets("BASE RL1").Cells(i, "M").Value = "Box O"
        Case "B16": Sheets("BASE RL1").Cells(i, "M").Value = "Box P"
        Case "B17": Sheets("BASE RL1").Cells(i, "M").Value = "Box Q"
        Case "B18": Sheets("BASE RL1").Cells(i, "M").Value = "Box R"
        Case "B19": Sheets("BASE RL1").Cells(i, "M").Value = "Box S"
        Case "B20": Sheets("BASE RL1").Cells(i, "M").Value = "Box T"
        Case "B21": Sheets("BASE RL1").Cells(i, "M").Value = "Box U"
        Case "B08": Sheets("BASE RL1").Cells(i, "M").Value = "Box V"
        Case "B09": Sheets("BASE RL1").Cells(i, "M").Value = "Box W"
        Case Else: Sheets("BASE RL1").Cells(i, "M").Value = texto ' Mantener el mismo valor
    End Select
Next i


'Poner todos los boxes para que el rango no cambie
lastRow = Sheets("BASE RL1").Cells(Sheets("BASE RL1").Rows.Count, 1).End(xlUp).row
ThisWorkbook.Sheets("Anexxes").Range("G2:G23").Copy
wbBases.Sheets("BASE RL1").Cells(lastRow + 1, "M").PasteSpecial Paste:=xlPasteValues
Application.CutCopyMode = False
ThisWorkbook.Sheets("Anexxes").Range("I2:I23").Copy
wbBases.Sheets("BASE RL1").Cells(lastRow + 1, "A").PasteSpecial Paste:=xlPasteValues
Application.CutCopyMode = False

'Crea tabla dinamica
Sheets("BASE RL1").Activate

    Dim ult_Tabla As Long
    ult_Tabla = Cells(Rows.Count, "A").End(xlUp).row
    Dim rangoTabla1 As Range
    Set rangoTabla1 = Sheets("BASE RL1").Range("A1:Q" & ult_Tabla)
    ActiveSheet.ListObjects.Add(xlSrcRange, rangoTabla1, , xlYes).Name = "Tabla1"
    
    'Crear tabla dinamica
    Dim celdaTablaDinamica1 As Range
    Set celdaTablaDinamica1 = Sheets("TD RL1").Range("A1")
    Dim tablaDinamica1 As PivotTable
    
    'Activa campos y le pone formato tabular
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        rangoTabla1, Version:=6).CreatePivotTable TableDestination:= _
        celdaTablaDinamica1, TableName:="tablaDinamica1", DefaultVersion:=6
        
    Sheets("TD RL1").Select
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
    With ActiveSheet.PivotTables("tablaDinamica1").PivotFields("BOX NAME")
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

wbBases.Save

'-------------------------- CREAR ARCHIVO DE LA PLANTILLA DEL RL1 --------------------------

'Hace copia de la plantilla del RL1
On Error Resume Next
Kill Ruta_Año & "\" & año & mes & " RL1 Audits" & ".xlsx"
On Error GoTo 0

Set wbPlantillaRL1 = Workbooks.Open(ruta & "\" & "RL1 Audits.xlsx")
wbPlantillaRL1.Activate
ActiveWorkbook.SaveCopyAs Filename:=Ruta_Año & "\" & año & mes & " RL1 Audits" & ".xlsx"
wbPlantillaRL1.Close False

'Abre los archivos correspondientes
Set wbPlantillaRL1 = Workbooks.Open(Ruta_Año & "\" & año & mes & " RL1 Audits" & ".xlsx")

'Pasa la tabla dinamica como valores al documento del RL1
wbBases.Activate
wbBases.Sheets("TD RL1").Activate
lastRow = Cells(Rows.Count, "A").End(xlUp).row
Range("A2:AB" & lastRow - 1).Select
Range("A2:AB" & lastRow - 1).Copy
wbPlantillaRL1.Activate
Sheets("TD RL1").Activate
Range("B1").Select
Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False
Application.CutCopyMode = False
With Range("A1:AA1")
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
Sheets("TD RL1").Range("A2:A" & lastRow) = "=+CONCATENATE(RC[1],RC[2],RC[3])"
Columns("A:Y").AutoFit

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

'-------------------------- Ciclo para pegar las 4 columnas principales --------------------------
'NOTA:
    'Hoja 3 - BOX A >= BOX L+ Box J
    'Hoja 4 - BoxA >= Box G up to QPP max
    'Hoja 5 - BOX E < = BOX  A (Except for severance payments)
    'Hoja 6 - BoxA>=0 and Box E=Blank (should be blank only for exempt employees in TD1 (FIT))
    'Hoja 7 - Box A can not be less than Box J + Box L + Box V + Box W
    'Hoja 8 - if BOX O >0 then Code cannot be blank
    'Hoja 9 - Only one RL1 per employee and all QC employees have at least one T4 slip
    



For hoja = 3 To 9

    'Seleccina los datos que va a copiar
    wbPlantillaRL1.Sheets("TD RL1").Activate
    lastRow = Cells(Rows.Count, "B").End(xlUp).row
    Range("B1:D" & lastRow).Select
    Range("B1:D" & lastRow).Copy
    
    'Pega datos en cada hoja de los reportes
    Set ws = wbPlantillaRL1.Sheets(hoja)
    ws.Activate
    Range("A3").Select
    Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    
    'Seleccina los datos que va a copiar
    wbPlantillaRL1.Sheets("TD RL1").Activate
    lastRow = Cells(Rows.Count, "B").End(xlUp).row
    Range("A1:A" & lastRow).Select
    Range("A1:A" & lastRow).Copy
    
    'Pega datos en cada hoja de los reportes
    Set ws = wbPlantillaRL1.Sheets(hoja)
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

'-------------------------- REPARTE LAS COLUMNAS DE LOS BOXES PARA LAS AUDIRORIAS --------------------------

Dim wsTDRL1, wsDestino As Worksheet
Dim col As Range
Dim destCol As Long
Dim Posicion_hojas, criterios As Variant

'Setear la hoja de donde toma las columnas
Set wsTDRL1 = wbPlantillaRL1.Sheets("TD RL1")

'Definir las hojas y los criterios, el orden de los criterios esta segun el orden de las hojas en el archivo
'NOTA:
    'Hoja 3 - BOX A >= BOX L+ Box J
    'Hoja 4 - BoxA >= Box G up to QPP max
    'Hoja 5 - BOX E < = BOX  A (Except for severance payments)
    'Hoja 6 - BoxA>=0 and Box E=Blank (should be blank only for exempt employees in TD1 (FIT))
    'Hoja 7 - Box A can not be less than Box J + Box L + Box V + Box W
    'Hoja 8 - if BOX O >0 then Code cannot be blank
    'Hoja 9 - Only one RL1 per employee and all QC employees have at least one T4 slip
    
Posicion_hojas = Array(3, 4, 5, 6, 7, 8) 'Las hojas en las que va a poner las columas

'Criterios para cada hoja. IMPORTANE: cada renglon es el crioterio de la hoja por posicion
criterios = Array( _
    Array("Box A", "Box L", "Box J"), _
    Array("Box A", "Box G"), _
    Array("Box E", "Box A", "Box O"), _
    Array("Box A", "Box E", "Box J", "Box L"), _
    Array("Box A", "Box J", "Box L", "Box V", "Box W"), _
    Array("Box O") _
)

'Iterar sobre las hojas
For i = LBound(Posicion_hojas) To UBound(Posicion_hojas)

    'Asignar la hoja de destino
    Set wsDestino = wbPlantillaRL1.Sheets(Posicion_hojas(i))
    lastRow = wsDestino.Cells(wsDestino.Rows.Count, 1).End(xlUp).row
    
    'Indica que inicie a pegar los datos desde la columna E
    destCol = 5
    
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
For hoja = 3 To 9

    Set ws = wbPlantillaRL1.Sheets(hoja)
    ws.Activate
    lastCol = Cells(3, Columns.Count).End(xlToLeft).Column
    With Range(Cells(3, 5), Cells(3, lastCol))
        .Interior.Color = RGB(243, 156, 18)
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
    End With
    Columns("A:D").AutoFit
    
Next hoja

wbPlantillaRL1.Save

'-------------------------- CREA LAS FORMULAS DE LAS AUDITORIAS --------------------------


'Hoja 3 - BOX A >= BOX L+ Box J

With Sheets(3)
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
    .Range(.Cells(4, lastCol + 1), .Cells(lastRow, lastCol + 1)).FormulaR1C1 = "=RC[-3]>=(RC[-1]+RC[-2])"
    .Columns("A:M").AutoFit
End With

'Hoja 4 - BoxA >= Box G up to QPP max

With Sheets(4)
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
    .Range(.Cells(4, lastCol + 1), .Cells(lastRow, lastCol + 1)).FormulaR1C1 = "=RC[-2]>=RC[-1]"
    .Columns("A:M").AutoFit
End With

'Hoja 5 - BOX E < = BOX  A (Except for severance payments)

With Sheets(5)
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
    .Range(.Cells(4, lastCol + 2), .Cells(lastRow, lastCol + 2)).FormulaR1C1 = "=IFERROR(IF(RC[-1]=TRUE,"" "",IF(AND(RC[-1]=FALSE,RC[-2]>0),""OK- Severance Payment"",""Review"")),"" "")"
    .Columns("A:M").AutoFit
End With


'Hoja 6 - BoxA>=0 and Box E=Blank (should be blank only for exempt employees in TD1 (FIT))

With Sheets(6)
    .Activate
    .Range("F2").Value = BPA
    .Range("E2").Value = "BPA"
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
    .Range(.Cells(4, lastCol + 1), .Cells(lastRow, lastCol + 1)).FormulaR1C1 = "=IF(AND(RC[-4]>0,RC[-3]=0),""VERDADERO-REVIEW"",""FALSO-OK"")"
    .Range(.Cells(4, lastCol + 2), .Cells(lastRow, lastCol + 2)).FormulaR1C1 = "=IF(RC[-1]=""FALSO-OK"", "" "",IF(RC[-5]<R2C6,""OK- BoxA lower than BPA"",""REVIEW""))"
    .Range(.Cells(4, lastCol + 3), .Cells(lastRow, lastCol + 3)).FormulaR1C1 = "=IF(RC[-2]=""FALSO-OK"", "" "",IF(RC[-6]-(RC[-4]+RC[-3])=0,""OK- Income related to TB only"",""REVIEW""))"
    .Columns("A:M").AutoFit
    
End With

'Hoja 7 - Box A can not be less than Box J + Box L + Box V + Box W

With Sheets(7)
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
    .Range(.Cells(4, lastCol + 1), .Cells(lastRow, lastCol + 1)).FormulaR1C1 = "=RC[-5]>=(RC[-4]+RC[-3]+RC[-2]+RC[-1])"
    .Columns("A:M").AutoFit
End With

'Hoja 8 - if BOX O >0 then Code cannot be blank

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
    .Range(.Cells(4, lastCol + 1), .Cells(lastRow, lastCol + 1)).FormulaR1C1 = "=IF(RC[-1]>0,""REVIEW CODE RJ"",""OK"")"
    .Columns("A:M").AutoFit
End With


'Hoja 9 - Only one RL1 per employee and all QC employees have at least one T4 slip

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
    .Range(.Cells(4, lastCol + 1), .Cells(lastRow, lastCol + 1)).FormulaR1C1 = "=IF(COUNTIF(R4C1:R" & lastRow & "C1,RC[-4])>=2,""REVIEW"",""OK"")"
    .Columns("A:M").AutoFit
End With

'-------------------------- CAMBIO FINAL: pone la fecha en la que se ejecuto la macro --------------------------
wbPlantillaRL1.Sheets("AUDIT LIST").Range("M1").Value = Date
Sheets("AUDIT LIST").Activate
wbPlantillaRL1.Save
wbPlantillaRL1.Close

End Sub
