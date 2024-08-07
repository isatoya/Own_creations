Attribute VB_Name = "Organizar_doc"
Option Explicit
'variables para todo el proyecto

    Dim Fecha1, Fecha2, mes, Mes_Texto, año, catorcena As String
    Dim ruta, Ruta_Año, Ruta_Catorcena, Ruta_organizado As String
    Dim wbConsolidado, wbCreditosInicial As Workbook
    Dim CelsFecha As Range
    Dim lastRow, i As Long
    Dim CelsAG As Object
    
Sub Organizar_docs_local()

    ' Verificar si hay datos en las celdas I8 y M8
    If ThisWorkbook.Sheets("Principal").Range("L7").Value = "" Or ThisWorkbook.Sheets("Principal").Range("H7").Value = "" Or ThisWorkbook.Sheets("Principal").Range("H13").Value = "" Then
        MsgBox "Datos incompletos, por favor ingrese los datos antes de ejecutar.", vbExclamation
        Exit Sub
    End If

    ' Llama a cada una de las funciones
    CrearCarpetas
    crea_reportes
    
    MsgBox "Los exceles estan abiertos, por favor ingrese a realiza las conexiónes por medio de formulas entre ellos.", vbInformation

End Sub

Sub InicializarVariables()
'Definicion de las variables
    
    'Fechas
    mes = ThisWorkbook.Sheets("Principal").Range("M7").Text
    Mes_Texto = ThisWorkbook.Sheets("Principal").Range("H11").Value
    año = ThisWorkbook.Sheets("Principal").Range("M13").Value
    catorcena = ThisWorkbook.Sheets("Principal").Range("H13").Text
    Fecha1 = ThisWorkbook.Sheets("Principal").Range("H7").Value
    Fecha2 = ThisWorkbook.Sheets("Principal").Range("I7").Value
    
    'Rutas ------'ORIGINAL: G:\H2R\Mexico\PAYROLL\Novedades\CATORCENAS 2024\CATORCENA 08-2024\PAGOS A TERCEROS\SINDICATOS
    ruta = "G:\H2R\Mexico\PAYROLL\Novedades\"
    Ruta_Año = ruta & "CATORCENAS " & año
    Ruta_Catorcena = Ruta_Año & "\" & "CATORCENA " & catorcena & "-" & año & "\PAGOS A TERCEROS\SINDICATOS"
    Ruta_organizado = Ruta_Catorcena & "\" & "DOCS ORGANIZADOS"
    
    'ruta = "G:\H2R\Mexico\PAYROLL\Novedades\"
    'Ruta_Año = ruta & "CATORCENAS " & año
    'Ruta_Catorcena = "C:\Users\imontoy\Desktop\CATORCENA 08 - 2024"
    'Ruta_organizado = "C:\Users\imontoy\Desktop\CATORCENA 08 - 2024" & "\" & "DOCS ORGANIZADOS"
    
    

End Sub

Sub CrearCarpetas()
    
    'Creación y validacion de las carpetas
    InicializarVariables
    
    '''''''''''''
    ''CATORCENA''
    '''''''''''''
    Ruta_Catorcena = Ruta_Año & "\" & "CATORCENA " & catorcena & "-" & año & "\PAGOS A TERCEROS\SINDICATOS"
    If Dir(Ruta_Catorcena, vbDirectory + vbHidden) = "" Then
        'Comprueba que la carpeta no exista para crear el directorio.
        If Dir(Ruta_Catorcena & vbDirectory + vbHidden) = "" Then MkDir Ruta_Catorcena
    End If
    
    '''''''''''''''''''''''
    ''Documneto orgnizado''
    '''''''''''''''''''''''
    Ruta_organizado = Ruta_Catorcena & "\" & "DOCS ORGANIZADOS"
    If Dir(Ruta_organizado, vbDirectory + vbHidden) = "" Then
        'Comprueba que la carpeta no exista para crear el directorio.
        If Dir(Ruta_organizado & vbDirectory + vbHidden) = "" Then MkDir Ruta_organizado
    End If
        
End Sub


Sub crea_reportes()

    InicializarVariables
    Dim wbPlantilla As Workbook
    Dim wbZPYMX034 As Workbook
    Dim wbZPYMX025_completo As Workbook
    Dim wbZPYMX025_MX02 As Workbook
    
    
'------------------------------ Crea copia de los archivos originales ------------------------------

'Hace copia del archivo de relacion sociedad
    On Error Resume Next
    Kill Ruta_organizado & "\" & "ZPYMX025 CAT " & catorcena & " - " & año & "_Relacion Sociedad COMPLETO" & ".xlsx"
    On Error GoTo 0
    
    Set wbPlantilla = Workbooks.Open(Ruta_Catorcena & "\" & "Relacion sociedad.xlsx")
    wbPlantilla.Activate
    ActiveWorkbook.SaveCopyAs Filename:=Ruta_organizado & "\" & "ZPYMX025 CAT " & catorcena & " - " & año & "_Relacion Sociedad COMPLETO" & ".xlsx"
    wbPlantilla.Close False

'Hace copia del archivo de MX02_Relacion sociedad -
    On Error Resume Next
    Kill Ruta_organizado & "\" & "ZPYMX025 CAT " & catorcena & " - " & año & " MX02_Relacion Sociedad" & ".xlsx"
    On Error GoTo 0
    
    Set wbPlantilla = Workbooks.Open(Ruta_Catorcena & "\" & "MX02_Relacion Sociedad.xlsx")
    wbPlantilla.Activate
    ActiveWorkbook.SaveCopyAs Filename:=Ruta_organizado & "\" & "ZPYMX025 CAT " & catorcena & " - " & año & " MX02_Relacion Sociedad" & ".xlsx"
    wbPlantilla.Close False
    
'Hace copia del archivo de ZPYMXO34 CAT
    On Error Resume Next
    Kill Ruta_organizado & "\" & "ZPYMX034 CAT " & catorcena & ".xlsx"
    On Error GoTo 0
    
    Set wbPlantilla = Workbooks.Open(Ruta_Catorcena & "\" & "ZPYMX034 CAT " & catorcena & ".xlsx")
    wbPlantilla.Activate
    ActiveWorkbook.SaveCopyAs Filename:=Ruta_organizado & "\" & "ZPYMX034 CAT " & catorcena & ".xlsx"
    wbPlantilla.Close False



'------------------------------ Crea copia de los archivos originales ------------------------------


'----------- TRABAJA EN EL ZPYMX034 -----------

'Hace cambios de formato en el arhchivo ZPYMX034
    Set wbZPYMX034 = Workbooks.Open(Ruta_organizado & "\" & "ZPYMX034 CAT " & catorcena & ".xlsx")
    wbZPYMX034.Activate
    Sheets("Sheet1").Name = "ORIGINAL"
    Sheets("ORIGINAL").Copy After:=Sheets(Sheets.Count)
    ActiveSheet.Name = "MODIFICACIONES"
    Sheets("MODIFICACIONES").Tab.Color = RGB(255, 192, 0)
    Sheets("ORIGINAL").Activate
    Range("A2:J500").HorizontalAlignment = xlLeft
    Columns("A:J").AutoFit
    
'Cambia texto a numero en el archivo ZPYMX034
    Sheets("ORIGINAL").Activate
    lastRow = Cells(Rows.Count, "A").End(xlUp).Row
    For i = 2 To lastRow
        If IsNumeric(Range("B" & i).Value) Then
            Range("B" & i).Value = Val(Range("B" & i).Value)
        End If
    Next i
    
    For i = 2 To lastRow
        If IsNumeric(Range("D" & i).Value) Then
            Range("D" & i).Value = Val(Range("D" & i).Value)
        End If
    Next i
    
    For i = 2 To lastRow
        If IsNumeric(Range("J" & i).Value) Then
            Range("J" & i).Value = Val(Range("J" & i).Value)
        End If
    Next i
       
'Hace sumatoria de importe en la hoja original
    lastRow = Cells(Rows.Count, "F").End(xlUp).Row
    Range("F" & lastRow + 1).Formula = "=SUM(F2:F" & lastRow & ")"
    With Range("F" & lastRow + 1)
        .Interior.Color = RGB(255, 192, 0)
        .Font.Bold = True
    End With
    
    Range("F" & lastRow + 3).Value = "Formula del archivo de la ZPYMX025 RELACION COMPLETA"
    With Range("F" & lastRow + 3)
        .Interior.Color = RGB(255, 192, 0)
        .Font.Bold = True
    End With

'Resta entre los datos en el arhchivo ZPYMX034
    Sheets("ORIGINAL").Activate
    lastRow = Cells(Rows.Count, "E").End(xlUp).Row
    Range("F" & lastRow + 5).Formula = "=F" & lastRow + 1 & "-F" & lastRow + 3
    With Range("F" & lastRow + 5)
        .Font.Color = RGB(255, 0, 0)
        .Font.Bold = True
    End With
    Range("G" & lastRow + 5).Value = "Cuota Nacional"
    wbZPYMX034.Save
    
'Hace cambios de formato en la hoja de "MODIFICACIONES"
    Sheets("MODIFICACIONES").Activate
    With Range("A1:J1")
        .Font.Color = RGB(0, 0, 0)
        .Interior.Color = RGB(0, 176, 240)
        .RowHeight = 15
    End With
    Range("A2:J5000").HorizontalAlignment = xlLeft
    Columns("A:Z").AutoFit
    
    lastRow = Cells(Rows.Count, "A").End(xlUp).Row
    For i = 2 To lastRow
        If IsNumeric(Range("B" & i).Value) Then
            Range("B" & i).Value = Val(Range("B" & i).Value)
        End If
    Next i
    
    For i = 2 To lastRow
        If IsNumeric(Range("D" & i).Value) Then
            Range("D" & i).Value = Val(Range("D" & i).Value)
        End If
    Next i
    
    For i = 2 To lastRow
        If IsNumeric(Range("J" & i).Value) Then
            Range("J" & i).Value = Val(Range("J" & i).Value)
        End If
    Next i

'En la hoja de Modificaciones cambia el numero de la cuenta contable
    lastRow = Cells(Rows.Count, "J").End(xlUp).Row
    ActiveSheet.Range("J2:J" & lastRow).Value = "20206055"

'Cambia los numeros de negativos a positivos
    Columns("F:F").Select
    Selection.Replace What:="-", Replacement:="+", LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False
      
'Suma del importe
    lastRow = Cells(Rows.Count, "F").End(xlUp).Row
    Range("F" & lastRow + 1).Formula = "=SUM(F2:F" & lastRow & ")"
    With Range("F" & lastRow + 1)
        .Interior.Color = RGB(255, 192, 0)
        .Font.Bold = True
    End With
    ActiveWindow.Zoom = 80
    Columns("A:J").AutoFit
   
'Crea plantilla de titutlos
    lastRow = Cells(Rows.Count, "E").End(xlUp).Row
    
    With Range("E" & lastRow + 4)
        .Value = "TOTAL CHONA"
        .Font.Bold = True
        .HorizontalAlignment = xlRight
    End With
    
    With Range("E" & lastRow + 5)
        .Value = "TOTAL SECCION 40"
        .Font.Bold = True
        .HorizontalAlignment = xlRight
    End With
    
    With Range("E" & lastRow + 6)
        .Value = "TOTAL SECCION 51"
        .Font.Bold = True
        .HorizontalAlignment = xlRight
    End With
    
    With Range("E" & lastRow + 7)
        .Value = "TOTAL PAGOS"
        .Font.Bold = True
        .Interior.Color = RGB(255, 192, 0)
        .HorizontalAlignment = xlRight
    End With
    
    With Range("F" & lastRow + 3)
        .Value = "ZPYMX034"
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
    End With
    
    With Range("H" & lastRow + 3)
        .Value = "ZPYMX025"
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
    End With
    
    With Range("H" & lastRow + 7)
        .Font.Bold = True
        .Interior.Color = RGB(255, 192, 0)
    End With
    Range("H" & lastRow + 4 & ":H" & lastRow + 7).NumberFormat = "$#,##0.00"
    'Suma total
    Range("H" & lastRow + 7).Formula = "=SUM(H" & lastRow + 4 & ":H" & lastRow + 6 & ")"
    
'Sumas con condicionales
    Dim rangoE As Range
    Dim rangoF As Range
    Set rangoE = Range("E2:E" & lastRow)
    Set rangoF = Range("F2:F" & lastRow)
    
    'Contar si
    Range("F" & lastRow + 4).FormulaR1C1 = "=SUMIF(" & rangoE.Address(ReferenceStyle:=xlR1C1) & ",""CARLOS OROPEZA CHONA""," & rangoF.Address(ReferenceStyle:=xlR1C1) & ")"
    Range("F" & lastRow + 5).FormulaR1C1 = "=SUMIF(" & rangoE.Address(ReferenceStyle:=xlR1C1) & ",""SINDICATO DE TRABAJADORES DE CEMENTO SECCION 40""," & rangoF.Address(ReferenceStyle:=xlR1C1) & ")"
    Range("F" & lastRow + 6).FormulaR1C1 = "=SUMIF(" & rangoE.Address(ReferenceStyle:=xlR1C1) & ",""SINDICATO NACIONAL DE LA INDUSTRIA DEL CEMENTO SECCION 51""," & rangoF.Address(ReferenceStyle:=xlR1C1) & ")"
    
    'Suma total
    Range("F" & lastRow + 7).Formula = "=SUM(F" & lastRow + 4 & ":F" & lastRow + 6 & ")"
    Range("F" & lastRow + 4 & ":F" & lastRow + 7).NumberFormat = "$#,##0.00"
    
    With Range("F" & lastRow + 7)
        .Font.Bold = True
        .Interior.Color = RGB(255, 192, 0)
    End With
    
'Formulas de diferencias, resta F - H con respuesta en la I
    Range("I" & lastRow + 4 & ":I" & lastRow + 7).NumberFormat = "$#,##0.00"
    Range("I" & lastRow + 4).Formula = "=F" & lastRow + 4 & "+H" & lastRow + 4
    Range("I" & lastRow + 5).Formula = "=F" & lastRow + 5 & "+H" & lastRow + 5
    Range("I" & lastRow + 6).Formula = "=F" & lastRow + 6 & "+H" & lastRow + 6
    Range("I" & lastRow + 7).Formula = "=F" & lastRow + 7 & "+H" & lastRow + 7

'Crea plantilla de pagos
    Sheets.Add(After:=Sheets(Sheets.Count)).Name = "PAGOS"
    Sheets("PAGOS").Tab.Color = RGB(255, 0, 0)
    Sheets("PAGOS").Activate
    ActiveWindow.Zoom = 71
    Range("A1").Value = "Proveedor (31)"
    Range("B1").Value = "Numero SAP"
    Range("C1").Value = "Div."
    Range("D1").Value = "Nombre Sindicato"
    Range("E1").Value = "Importe"
    Range("F1").Value = "Destino"
    Range("G1").Value = "Cuenta con (40)"
    Range("H1").Value = "Solicitud"
    Rows("1").RowHeight = 29.25
    Rows("2:7").RowHeight = 25
    Columns("A").ColumnWidth = 15.43
    Columns("B").ColumnWidth = 14.43
    Columns("C").ColumnWidth = 8
    Columns("D").ColumnWidth = 67.29
    Columns("E").ColumnWidth = 19.57
    Columns("F").ColumnWidth = 20.86
    Columns("G").ColumnWidth = 12.43
    Columns("H").ColumnWidth = 10.14
    Range("A2:A3").Merge
    Range("A4:A5").Merge
    Range("A6:A7").Merge
    Range("B2:B3").Merge
    Range("B4:B5").Merge
    Range("B6:B7").Merge
    Range("C2:C3").Merge
    Range("C4:C5").Merge
    Range("C6:C7").Merge
    Range("D2:D3").Merge
    Range("D4:D5").Merge
    Range("D6:D7").Merge
    Range("E2:E3").Merge
    Range("E4:E5").Merge
    Range("E6:E7").Merge
    Range("F2:F3").Merge
    Range("F4:F5").Merge
    Range("F6:F7").Merge
    Range("G2:G3").Merge
    Range("G4:G5").Merge
    Range("G6:G7").Merge
    Range("H2:H3").Merge
    Range("H4:H5").Merge
    Range("H6:H7").Merge
    
    With Range("A1:H7").Borders
        .LineStyle = xlContinuous
        .Color = RGB(0, 0, 0)
        .Weight = xlThin
    End With
    
    Range("A10:H10").Merge
    Range("A10").Value = "PAGOS ADICIONALES AYUDA DE DEFUNCIÓN"
    Range("A11").Value = "Proveedor (31)"
    Range("B11").Value = "Numero SAP"
    Range("C11").Value = "Div."
    Range("D11").Value = "Nombre Sindicato"
    Range("E11").Value = "Importe"
    Range("F11").Value = "Destino"
    Range("G11").Value = "Cuenta con (40)"
    Range("H11").Value = "Solicitud"
    Rows("10").RowHeight = 22.5
    Rows("11").RowHeight = 29.25
    Rows("12:15").RowHeight = 21.75
    Range("A12:A13").Merge
    Range("A14:A15").Merge
    Range("B12:B13").Merge
    Range("B14:B15").Merge
    Range("C12:C13").Merge
    Range("C14:C15").Merge
    Range("D12:D13").Merge
    Range("D14:D15").Merge
    Range("E12:E13").Merge
    Range("E14:E15").Merge
    Range("F12:F13").Merge
    Range("F14:F15").Merge
    Range("G12:G13").Merge
    Range("G14:G15").Merge
    Range("H12:H13").Merge
    Range("H14:H15").Merge
    
    With Range("A10:H15").Borders
        .LineStyle = xlContinuous
        .Color = RGB(0, 0, 0)
        .Weight = xlThin
    End With
    
    'Color de fondo
    Range("A1:H1").Interior.Color = RGB(214, 220, 228)
    Range("A2:A6").Interior.Color = RGB(214, 220, 228)
    Range("G2:G6").Interior.Color = RGB(214, 220, 228)
    Range("A10").Interior.Color = RGB(214, 220, 228)
    Range("A11:H11").Interior.Color = RGB(214, 220, 228)
    Range("A12:A14").Interior.Color = RGB(214, 220, 228)
    Range("G12:G14").Interior.Color = RGB(214, 220, 228)
    Range("E18").Interior.Color = RGB(226, 239, 218)
    
    'Formato de texto
    Rows("1:1").Font.Bold = True
    Rows("10:11").Font.Bold = True
    Columns("A:H").Font.Name = "Tahoma"
    Columns("A:H").Font.Size = 11
    Range("A2:A6").Font.Color = RGB(255, 0, 0)
    Range("C2:D6").Font.Color = RGB(255, 0, 0)
    Range("F2:G6").Font.Color = RGB(255, 0, 0)
    Range("A12:A14").Font.Color = RGB(255, 0, 0)
    Range("C12:D14").Font.Color = RGB(255, 0, 0)
    Range("F12:G14").Font.Color = RGB(255, 0, 0)
    Range("B2:B6").Font.Color = RGB(0, 128, 128)
    Range("E2:E6").Font.Color = RGB(0, 128, 128)
    Range("H2:H6").Font.Color = RGB(0, 128, 128)
    Range("B12:B14").Font.Color = RGB(0, 128, 128)
    Range("E12:E14").Font.Color = RGB(0, 128, 128)
    Range("H12:H14").Font.Color = RGB(0, 128, 128)
    Range("A:H").HorizontalAlignment = xlCenter
    Range("A:H").VerticalAlignment = xlCenter
    Range("D2:D6").HorizontalAlignment = xlLeft
    Range("D2:D6").VerticalAlignment = xlCenter
    Range("D12:D15").Select
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
    End With
    Columns("E").NumberFormat = "$#,##0.00"
    Range("E8").Font.Color = RGB(0, 128, 128)
    Range("E8").Font.Bold = True
    Range("E16").Font.Color = RGB(0, 128, 128)
    Range("E16").Font.Bold = True
    Range("E18").Font.Bold = True
    
    'Ajustar texto
    Columns("A:H").Select
    With Selection
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
    End With
    
    'Copiar texto
    Range("A2").Value = "30445"
    Range("A4").Value = "4900000"
    Range("A6").Value = "4900010"
    Range("C2").Value = "1412"
    Range("C4").Value = "1421"
    Range("C6").Value = "1411"
    Range("D2").Value = "CARLOS OROPEZA CHONA"
    Range("D4").Value = "SINDICATO DE TRABAJADORES DE CEMENTO SECCION 40"
    Range("D6").Value = "SINDICATO NACIONAL DE LA INDUSTRIA DEL CEMENTO SECCION 51"
    Range("F2").Value = "PLANTA CEMENTOS ACAPULCO"
    Range("F4").Value = "PLANTA CEMENTOS ORIZABA"
    Range("F6").Value = "PLANTA CEMENTOS APAXCO"
    Range("G2").Value = "20206055"
    Range("G4").Value = "20206055"
    Range("G6").Value = "20206055"
    Range("G12").Value = "20206055"
    Range("G14").Value = "20206055"
    
    'Suma
    Range("E8").Formula = "=SUM(E2:E6)"
    Range("E16").Formula = "=SUM(E12:E15)"
    Range("E18").Formula = "=E8+E16"
    
    'Formulas en importe
    Range("E2").Formula = "='MODIFICACIONES'!F" & lastRow + 4
    Range("E4").Formula = "='MODIFICACIONES'!F" & lastRow + 5
    Range("E6").Formula = "='MODIFICACIONES'!F" & lastRow + 6

    Sheets("MODIFICACIONES").Activate
    Rows("1:1").Select
    With ActiveWindow
        .SplitColumn = 0
        .SplitRow = 1
    End With
    ActiveWindow.FreezePanes = True
    wbZPYMX034.Save
    
'----------- Trabaja en el wbZPYMX025_Relacion sociedad completo -----------

'Trabaja en wbZPYMX025_completo
    Set wbZPYMX025_completo = Workbooks.Open(Ruta_organizado & "\" & "ZPYMX025 CAT " & catorcena & " - " & año & "_Relacion Sociedad COMPLETO" & ".xlsx")
    wbZPYMX025_completo.Activate
    Sheets("Sheet1").Name = "ZPYMX025 CAT " & catorcena & " - " & año
    Columns("O:P").Select
    Selection.Delete
    
' Aplicar el formato a la fila 1
    With Range("A1:U1")
        .Interior.Color = RGB(0, 176, 240)
        .RowHeight = 15
    End With
    Range("A2:U5000").HorizontalAlignment = xlLeft
    Columns("A:Z").AutoFit

    
'Pone autifiltros y fija la primera fila
    Sheets(1).Rows(1).AutoFilter
    Sheets(1).Rows(1).Select
    Rows("1:1").Select
    With ActiveWindow
        .SplitColumn = 0
        .SplitRow = 1
    End With
    ActiveWindow.FreezePanes = True
    
'Hace cambios de texto a nuemero
    lastRow = Cells(Rows.Count, "F").End(xlUp).Row
    For i = 2 To lastRow
        If IsNumeric(Range("A" & i).Value) Then
            Range("A" & i).Value = Val(Range("A" & i).Value)
        End If
    Next i
    
    For i = 2 To lastRow
        If IsNumeric(Range("G" & i).Value) Then
            Range("G" & i).Value = Val(Range("G" & i).Value)
        End If
    Next i

    For i = 2 To lastRow
        If IsNumeric(Range("K" & i).Value) Then
            Range("K" & i).Value = Val(Range("K" & i).Value)
        End If
    Next i
    
    For i = 2 To lastRow
        If IsNumeric(Range("M" & i).Value) Then
            Range("M" & i).Value = Val(Range("M" & i).Value)
        End If
    Next i

    For i = 2 To lastRow
        If IsNumeric(Range("O" & i).Value) Then
            Range("O" & i).Value = Val(Range("O" & i).Value)
        End If
    Next i
    
    For i = 2 To lastRow
        If IsNumeric(Range("P" & i).Value) Then
            Range("P" & i).Value = Val(Range("P" & i).Value)
        End If
    Next i
    
    For i = 2 To lastRow
        If IsNumeric(Range("U" & i).Value) Then
            Range("U" & i).Value = Val(Range("U" & i).Value)
        End If
    Next i
    
   
'Elimina las filas que tengan importe 0
    lastRow = Cells(Rows.Count, "S").End(xlUp).Row
    For i = lastRow To 1 Step -1
        If Cells(i, "S").Value = 0 Then
            Rows(i).Delete
        End If
    Next i
    
'Hace sumatoria del importe
    lastRow = Cells(Rows.Count, "S").End(xlUp).Row
    Range("S" & lastRow + 1).Formula = "=SUM(S2:S" & lastRow & ")"
    With Range("S" & lastRow + 1)
        .Interior.Color = RGB(255, 192, 0)
        .Font.Bold = True
    End With
    ActiveWindow.Zoom = 80
    Columns("A:Z").AutoFit
   
'Organiza de menor a mayor
    ActiveSheet.AutoFilter.Sort.SortFields. _
        Add Key:=Range("G1"), SortOn:=xlSortOnValues, Order:=xlAscending, _
        DataOption:=xlSortNormal
    With ActiveSheet.AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    wbZPYMX025_completo.Save
    


'----------- Trabaja en el wbZPYMX025_ MX02 relacion sociedad -----------

    Set wbZPYMX025_MX02 = Workbooks.Open(Ruta_organizado & "\" & "ZPYMX025 CAT " & catorcena & " - " & año & " MX02_Relacion Sociedad" & ".xlsx")
    wbZPYMX025_MX02.Activate
    Sheets("Hoja1").Name = "ZPYMX025 CAT " & catorcena & " - " & año
    wbZPYMX025_MX02.Sheets("ACREEDOR 30445 LOCAL").Move Before:=wbZPYMX025_MX02.Sheets(2)
    wbZPYMX025_MX02.Sheets("ACREEDOR 4900000 LOCAL ").Move Before:=wbZPYMX025_MX02.Sheets(3)

'Hoja inicial "ZPYMX025 CAT " & catorcena & " - " & año ---------------------

    Sheets(1).Activate
    lastRow = Cells(Rows.Count, "A").End(xlUp).Row
    Columns("O:P").Select
    Selection.Delete
    With Range("A1:U1")
        .Interior.Color = RGB(0, 176, 240)
        .RowHeight = 15
    End With
    Range("A2:U5000").HorizontalAlignment = xlLeft
    Columns("A:Z").AutoFit
    Sheets(1).Rows(1).AutoFilter
    Sheets(1).Rows(1).Select
    Rows("1:1").Select
    Rows("1:1").Select
    With ActiveWindow
        .SplitColumn = 0
        .SplitRow = 1
    End With
    ActiveWindow.FreezePanes = True
    
    'Convertir en numero
    For i = 2 To lastRow
        If IsNumeric(Range("A" & i).Value) Then
            Range("A" & i).Value = Val(Range("A" & i).Value)
        End If
    Next i
    
    For i = 2 To lastRow
        If IsNumeric(Range("G" & i).Value) Then
            Range("G" & i).Value = Val(Range("G" & i).Value)
        End If
    Next i
    
    For i = 2 To lastRow
        If IsNumeric(Range("K" & i).Value) Then
            Range("K" & i).Value = Val(Range("K" & i).Value)
        End If
    Next i
    
    For i = 2 To lastRow
        If IsNumeric(Range("M" & i).Value) Then
            Range("M" & i).Value = Val(Range("M" & i).Value)
        End If
    Next i
    
    For i = 2 To lastRow
        If IsNumeric(Range("O" & i).Value) Then
            Range("O" & i).Value = Val(Range("O" & i).Value)
        End If
    Next i
    
    For i = 2 To lastRow
        If IsNumeric(Range("P" & i).Value) Then
            Range("P" & i).Value = Val(Range("P" & i).Value)
        End If
    Next i
    
    For i = 2 To lastRow
        If IsNumeric(Range("U" & i).Value) Then
            Range("U" & i).Value = Val(Range("U" & i).Value)
        End If
    Next i
    
    'Elimina las filas que tengan importe 0
    lastRow = Cells(Rows.Count, "A").End(xlUp).Row
    
    For i = lastRow To 1 Step -1
        If Cells(i, "S").Value = 0 Then
            Rows(i).Delete
        End If
    Next i
    
    'Organiza de menor a mayor
    ActiveSheet.AutoFilter.Sort.SortFields. _
        Add Key:=Range("G1"), SortOn:=xlSortOnValues, Order:=xlAscending, _
        DataOption:=xlSortNormal
    With ActiveSheet.AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    'Hace sumatoria del importe
    lastRow = Cells(Rows.Count, "S").End(xlUp).Row
    Range("S" & lastRow + 1).Formula = "=SUM(S2:S" & lastRow & ")"
    With Range("S" & lastRow + 1)
        .Interior.Color = RGB(255, 192, 0)
        .Font.Bold = True
    End With
    ActiveWindow.Zoom = 80
    Columns("A:U").AutoFit
    wbZPYMX025_MX02.Save
 
'Hoja ACREEDOR 30445 LOCAL ---------------------

    Sheets(2).Activate
    lastRow = Cells(Rows.Count, "A").End(xlUp).Row
    ActiveSheet.Tab.Color = RGB(255, 192, 0)
    Columns("O:P").Select
    Selection.Delete
    With Range("A1:U1")
        .Interior.Color = RGB(0, 176, 240)
        .RowHeight = 15
    End With
    Range("A2:U5000").HorizontalAlignment = xlLeft
    Columns("A:Z").AutoFit
    Sheets(2).Rows(1).AutoFilter
    Sheets(2).Rows(1).Select
    Rows("1:1").Select
    Rows("1:1").Select
    With ActiveWindow
        .SplitColumn = 0
        .SplitRow = 1
    End With
    ActiveWindow.FreezePanes = True
    
    'Convertir en numero
    For i = 2 To lastRow
        If IsNumeric(Range("A" & i).Value) Then
            Range("A" & i).Value = Val(Range("A" & i).Value)
        End If
    Next i
    
    For i = 2 To lastRow
        If IsNumeric(Range("G" & i).Value) Then
            Range("G" & i).Value = Val(Range("G" & i).Value)
        End If
    Next i
    
    For i = 2 To lastRow
        If IsNumeric(Range("K" & i).Value) Then
            Range("K" & i).Value = Val(Range("K" & i).Value)
        End If
    Next i
    
    For i = 2 To lastRow
        If IsNumeric(Range("M" & i).Value) Then
            Range("M" & i).Value = Val(Range("M" & i).Value)
        End If
    Next i
    
    For i = 2 To lastRow
        If IsNumeric(Range("O" & i).Value) Then
            Range("O" & i).Value = Val(Range("O" & i).Value)
        End If
    Next i
    
    For i = 2 To lastRow
        If IsNumeric(Range("P" & i).Value) Then
            Range("P" & i).Value = Val(Range("P" & i).Value)
        End If
    Next i
    
    For i = 2 To lastRow
        If IsNumeric(Range("U" & i).Value) Then
            Range("U" & i).Value = Val(Range("U" & i).Value)
        End If
    Next i
    
    'Elimina las filas que tengan importe 0
    lastRow = Cells(Rows.Count, "A").End(xlUp).Row
    
    For i = lastRow To 1 Step -1
        If Cells(i, "S").Value = 0 Then
            Rows(i).Delete
        End If
    Next i
    
    'Organiza de menor a mayor
    ActiveSheet.AutoFilter.Sort.SortFields. _
        Add Key:=Range("G1"), SortOn:=xlSortOnValues, Order:=xlAscending, _
        DataOption:=xlSortNormal
    With ActiveSheet.AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

    'Hace sumatoria del importe
    lastRow = Cells(Rows.Count, "A").End(xlUp).Row
    Range("S" & lastRow + 1).Formula = "=SUM(S2:S" & lastRow & ")"
    With Range("S" & lastRow + 1)
        .Interior.Color = RGB(255, 192, 0)
        .Font.Bold = True
    End With
    ActiveWindow.Zoom = 80
    Columns("A:U").AutoFit
    wbZPYMX025_MX02.Save

'Hoja ACREEDOR  4900000 LOCAL ---------------------
    Sheets(3).Activate
    lastRow = Cells(Rows.Count, "A").End(xlUp).Row
    ActiveSheet.Tab.Color = RGB(255, 192, 0)
    Columns("O:P").Select
    Selection.Delete
    With Range("A1:U1")
        .Interior.Color = RGB(0, 176, 240)
        .RowHeight = 15
    End With
    Range("A2:U5000").HorizontalAlignment = xlLeft
    Columns("A:Z").AutoFit
    Sheets(3).Rows(1).AutoFilter
    Sheets(3).Rows(1).Select
    Rows("1:1").Select
    Rows("1:1").Select
    With ActiveWindow
        .SplitColumn = 0
        .SplitRow = 1
    End With
    ActiveWindow.FreezePanes = True
    
    'Convertir en numero
    For i = 2 To lastRow
        If IsNumeric(Range("A" & i).Value) Then
            Range("A" & i).Value = Val(Range("A" & i).Value)
        End If
    Next i
    
    For i = 2 To lastRow
        If IsNumeric(Range("G" & i).Value) Then
            Range("G" & i).Value = Val(Range("G" & i).Value)
        End If
    Next i
    
    For i = 2 To lastRow
        If IsNumeric(Range("K" & i).Value) Then
            Range("K" & i).Value = Val(Range("K" & i).Value)
        End If
    Next i
    
    For i = 2 To lastRow
        If IsNumeric(Range("M" & i).Value) Then
            Range("M" & i).Value = Val(Range("M" & i).Value)
        End If
    Next i
    
    For i = 2 To lastRow
        If IsNumeric(Range("O" & i).Value) Then
            Range("O" & i).Value = Val(Range("O" & i).Value)
        End If
    Next i
    
    For i = 2 To lastRow
        If IsNumeric(Range("P" & i).Value) Then
            Range("P" & i).Value = Val(Range("P" & i).Value)
        End If
    Next i
    
    For i = 2 To lastRow
        If IsNumeric(Range("U" & i).Value) Then
            Range("U" & i).Value = Val(Range("U" & i).Value)
        End If
    Next i
    
    'Elimina las filas que tengan importe 0
    lastRow = Cells(Rows.Count, "A").End(xlUp).Row
    
    For i = lastRow To 1 Step -1
        If Cells(i, "S").Value = 0 Then
            Rows(i).Delete
        End If
    Next i
    
    'Organiza de menor a mayor
    ActiveSheet.AutoFilter.Sort.SortFields. _
        Add Key:=Range("G1"), SortOn:=xlSortOnValues, Order:=xlAscending, _
        DataOption:=xlSortNormal
    With ActiveSheet.AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

    'Hace sumatoria del importe
    lastRow = Cells(Rows.Count, "A").End(xlUp).Row
    Range("S" & lastRow + 1).Formula = "=SUM(S2:S" & lastRow & ")"
    With Range("S" & lastRow + 1)
        .Interior.Color = RGB(255, 192, 0)
        .Font.Bold = True
    End With
    ActiveWindow.Zoom = 80
    Columns("A:U").AutoFit
    wbZPYMX025_MX02.Save
    
'Hoja ACREEDOR 4900010 LOCAL ---------------------
    Sheets(4).Activate
    lastRow = Cells(Rows.Count, "A").End(xlUp).Row
    ActiveSheet.Tab.Color = RGB(255, 192, 0)
    Columns("O:P").Select
    Selection.Delete
    With Range("A1:U1")
        .Interior.Color = RGB(0, 176, 240)
        .RowHeight = 15
    End With
    Range("A2:U5000").HorizontalAlignment = xlLeft
    Columns("A:Z").AutoFit
    Sheets(4).Rows(1).AutoFilter
    Sheets(4).Rows(1).Select
    Rows("1:1").Select
    With ActiveWindow
        .SplitColumn = 0
        .SplitRow = 1
    End With
    ActiveWindow.FreezePanes = True
    
    'Convertir en numero
    For i = 2 To lastRow
        If IsNumeric(Range("A" & i).Value) Then
            Range("A" & i).Value = Val(Range("A" & i).Value)
        End If
    Next i
    
    For i = 2 To lastRow
        If IsNumeric(Range("G" & i).Value) Then
            Range("G" & i).Value = Val(Range("G" & i).Value)
        End If
    Next i
    
    For i = 2 To lastRow
        If IsNumeric(Range("K" & i).Value) Then
            Range("K" & i).Value = Val(Range("K" & i).Value)
        End If
    Next i
    
    For i = 2 To lastRow
        If IsNumeric(Range("M" & i).Value) Then
            Range("M" & i).Value = Val(Range("M" & i).Value)
        End If
    Next i
    
    For i = 2 To lastRow
        If IsNumeric(Range("O" & i).Value) Then
            Range("O" & i).Value = Val(Range("O" & i).Value)
        End If
    Next i
    
    For i = 2 To lastRow
        If IsNumeric(Range("P" & i).Value) Then
            Range("P" & i).Value = Val(Range("P" & i).Value)
        End If
    Next i
    
    For i = 2 To lastRow
        If IsNumeric(Range("U" & i).Value) Then
            Range("U" & i).Value = Val(Range("U" & i).Value)
        End If
    Next i
    
    'Elimina las filas que tengan importe 0
    lastRow = Cells(Rows.Count, "A").End(xlUp).Row
    
    For i = lastRow To 1 Step -1
        If Cells(i, "S").Value = 0 Then
            Rows(i).Delete
        End If
    Next i
    
    'Organiza de menor a mayor
    ActiveSheet.AutoFilter.Sort.SortFields. _
        Add Key:=Range("G1"), SortOn:=xlSortOnValues, Order:=xlAscending, _
        DataOption:=xlSortNormal
    With ActiveSheet.AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

    'Hace sumatoria del importe
    lastRow = Cells(Rows.Count, "A").End(xlUp).Row
    Range("S" & lastRow + 1).Formula = "=SUM(S2:S" & lastRow & ")"
    With Range("S" & lastRow + 1)
        .Interior.Color = RGB(255, 192, 0)
        .Font.Bold = True
    End With
    ActiveWindow.Zoom = 80
    Columns("A:U").AutoFit
    wbZPYMX025_MX02.Save
    
End Sub



