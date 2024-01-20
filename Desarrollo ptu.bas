Attribute VB_Name = "M�dulo1"
Option Explicit

    Dim A�o As Integer
    Dim ruta As String
    Dim ruta2 As String
    Dim ruta3 As String
    Dim ruta4 As String

Sub InicializarVariables()

    ruta = ThisWorkbook.Path
    A�o = UserForm1.ComboBox1.Value
    ruta2 = ruta & "\" & CStr(A�o) & "\1. PW PTU"
    ruta3 = ruta & "\" & CStr(A�o) & "\3. SIND MAL ALTO - DIAS PESOS"
    ruta4 = ruta & "\" & CStr(A�o) & "\4. FACTOR DIAS PESOS"
    
End Sub

Sub PW_PTU()

    'Codigo para crear un nuevo archivo en blanco para crear el informe
    InicializarVariables

    ' Crear un nuevo libro de Excel
    Dim NuevoLibro As Workbook
    Set NuevoLibro = Workbooks.Add

    ' Renombrar las hojas
    Dim i As Integer
    For i = 1 To 5
        NuevoLibro.Sheets.Add(After:=NuevoLibro.Sheets(NuevoLibro.Sheets.Count)).Name = "Sheet" & i
    Next i
    Application.DisplayAlerts = False
    NuevoLibro.Sheets("Hoja1").Delete
    Application.DisplayAlerts = True
    
    NuevoLibro.SaveAs ruta2 & "\PW PTU " & A�o & ".xlsx" 'Guardar libro

End Sub


Sub Organizar_paginas()
'Pasa cada hoja de los informes descargados al archivo nuevo, con el objetivo de unificarlos
    
    InicializarVariables

    'Abre el archivo ZHR929
    Workbooks.Open ruta2 & "\ZHR929" & ".xlsx"
    ActiveSheet.Cells.Select
    Selection.Copy
    Workbooks("PW PTU " & A�o & ".xlsx").Activate
    Sheets("Sheet1").Activate
    Range("A1").PasteSpecial xlPasteAll
    ActiveSheet.Name = "ZHR929 FLEX Y CCTO"
    Application.CutCopyMode = False
    Workbooks("ZHR929.xlsx").Close SaveChanges:=False
    
    
    'Abre el  archivo ZHRMX27
    Workbooks.Open ruta2 & "\ZHRMX27" & ".xlsx"
    ActiveSheet.Cells.Select
    Selection.Copy
    Workbooks("PW PTU " & A�o & ".xlsx").Activate
    Sheets("Sheet2").Activate
    Range("A1").PasteSpecial xlPasteAll
    ActiveSheet.Name = "ZHRMX27 FECHA DEL REING"
    Application.CutCopyMode = False
    Workbooks("ZHRMX27.xlsx").Close SaveChanges:=False
    
    
    'Abre el archivo ZPYMX025 (PTU DIAS DETALLADO)
    Workbooks.Open ruta2 & "\ZPYMX025" & ".xlsx"
    ActiveSheet.Cells.Select
    Selection.Copy
    Workbooks("PW PTU " & A�o & ".xlsx").Activate
    Sheets("Sheet3").Activate
    Range("A1").PasteSpecial xlPasteAll
    ActiveSheet.Name = "PTU DIAS DETALLADO"
    Application.CutCopyMode = False
    Workbooks("ZPYMX025.xlsx").Close SaveChanges:=False
    
    
    'Abre el archivo ZPYMX025_V2 (PTU PESOS DETALLADO)
    Workbooks.Open ruta2 & "\ZPYMX025_V2" & ".xlsx"
    ActiveSheet.Cells.Select
    Selection.Copy
    Workbooks("PW PTU " & A�o & ".xlsx").Activate
    Sheets("Sheet4").Activate
    Range("A1").PasteSpecial xlPasteAll
    ActiveSheet.Name = "PTU PESOS DETALLADO"
    Application.CutCopyMode = False
    Workbooks("ZPYMX025_V2.xlsx").Close SaveChanges:=False
    
    
    'Abre el ZPYMX025_AUSENTISMOS (AUSENTISMOS SIND)
    Workbooks.Open ruta2 & "\ZPYMX025_AUSENTISMOS" & ".xlsx"
    ActiveSheet.Cells.Select
    Selection.Copy
    Workbooks("PW PTU " & A�o & ".xlsx").Activate
    Sheets("Sheet5").Activate
    Range("A1").PasteSpecial xlPasteAll
    ActiveSheet.Name = "AUSENTISMOS SIND"
    Application.CutCopyMode = False
    Workbooks("ZPYMX025_AUSENTISMOS.xlsx").Close SaveChanges:=False
    
    
    Workbooks("PW PTU " & A�o & ".xlsx").Activate 'Guardar
    ActiveWorkbook.Save
    
    
End Sub

Sub Cambios_formato()
'Hacer cambios en todo el informe nuevo y hace el buscarv en la ZHR929
    
    InicializarVariables
    Dim lastRow As Long
    Dim lastrow_put_dias As Long
    Dim lastrow_27 As Long
    Dim lastrow_put_pesos As Long
    Dim lastrow_put_ausen As Long
    
    'En la ZHR929 FLEX Y CCTO: Pone titulos en las columnas y hace cambios de formato
        Workbooks("PW PTU " & A�o & ".xlsx").Activate
        Sheets("ZHR929 FLEX Y CCTO").Activate
        Range("C1").Value = "Fecha de ingreso"
        Range("D1").Value = "Fecha de baja"
        Columns("E:E").Insert Shift:=xlToLeft, CopyOrigin:=xlFormatFromLeftOrAbove
        Range("E1").Value = "Fecha de reingreso"
        Range("S1").Value = "PTU DIAS DEF"
        Range("T1").Value = "AJ DIAS MEDICOS"
        Range("U1").Value = "TOTAL DIAS"
        Range("V1").Value = "PTU PESOS DEF"
        Range("W1").Value = "OBSERVACION"
        
        With Range("A1:W1")
            .Interior.Color = RGB(192, 192, 192)
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
        End With
        Columns("A:W").AutoFit
    
        'Realiza el buscarV para la fecha de ingreso
        Sheets("ZHR929 FLEX Y CCTO").Select
        lastRow = ActiveSheet.Cells(Rows.Count, 1).End(xlUp).Row
        With Worksheets("ZHR929 FLEX Y CCTO").Range("E2:E" & lastRow)
        .Formula = "=IF(ISERROR(VLOOKUP(RC[-4], 'ZHRMX27 FECHA DEL REING'!C[-4]:C[4], 9, 0)),"""",VLOOKUP(RC[-4], 'ZHRMX27 FECHA DEL REING'!C[-4]:C[4], 9, 0))"
        End With
    
        Columns("E").Copy 'Copia y pega como valores para quitar las formulas
        Columns("E").PasteSpecial Paste:=xlPasteValues
        Application.CutCopyMode = False


    'En la PTU DIAS DETALLADO: Cambia a formato numero y organiza titulos
        Sheets("PTU DIAS DETALLADO").Activate
        Columns("Z:Z").Insert Shift:=xlToRight
        lastrow_put_dias = ActiveSheet.Cells(Rows.Count, 1).End(xlUp).Row
        With Range("Z2:Z" & lastrow_put_dias)
            .Formula = "=VALUE(SUBSTITUTE(RC[-1],"" "", """"))"
        End With
        
        Range("Z1").Value = "Ctd."
        Columns("Z").Copy
        Columns("Z").PasteSpecial Paste:=xlPasteValues
        Application.CutCopyMode = False
        Columns("Z:Z").NumberFormat = "0.00"
        Columns("Y").Delete

            With Range("A1:AB1")
                .Interior.Color = RGB(192, 192, 192)
                .Font.Bold = True
                .HorizontalAlignment = xlCenter
            End With
            
        Columns("A:AB").AutoFit


    'En la PTU PESOS DETALLADO: Cambia a formato numero y organiza titulos
        Sheets("PTU PESOS DETALLADO").Activate
        Columns("AC:AC").Insert Shift:=xlToRight
        lastrow_put_pesos = ActiveSheet.Cells(Rows.Count, 1).End(xlUp).Row
        With Range("AC2:AC" & lastrow_put_pesos)
            .Formula = "=VALUE(SUBSTITUTE(RC[-1],"" "", """"))"
        End With
        
        Range("AC1").Value = "Importe"
        Columns("AC").Copy
        Columns("AC").PasteSpecial Paste:=xlPasteValues
        Application.CutCopyMode = False
        Columns("AC:AC").NumberFormat = "0.00"
        Columns("AB").Delete

            With Range("A1:AD1")
                .Interior.Color = RGB(192, 192, 192)
                .Font.Bold = True
                .HorizontalAlignment = xlCenter
            End With
            
        Columns("A:AD").AutoFit
                
    'En la AUSENTISMOS SIND: Cambia a formato numero y organiza titulos
        Sheets("AUSENTISMOS SIND").Activate
        Columns("AC:AC").Insert Shift:=xlToRight
        lastrow_put_ausen = ActiveSheet.Cells(Rows.Count, 1).End(xlUp).Row
        With Range("AC2:AC" & lastrow_put_ausen)
            .Formula = "=VALUE(SUBSTITUTE(RC[-1],"" "", """"))"
        End With
        
        Range("AC1").Value = "Ctd."
        Columns("AC").Copy
        Columns("AC").PasteSpecial Paste:=xlPasteValues
        Application.CutCopyMode = False
        Columns("AC:AC").NumberFormat = "0.00"
        Columns("AB").Delete

            With Range("A1:AE1")
                .Interior.Color = RGB(192, 192, 192)
                .Font.Bold = True
                .HorizontalAlignment = xlCenter
            End With
            
        Columns("A:AE").AutoFit
    
    'Cambiar formato de la columna salario a numero
    Sheets("ZHRMX27 FECHA DEL REING").Activate
    lastrow_27 = ActiveSheet.Cells(Rows.Count, 1).End(xlUp).Row
    With Range("L2:L" & lastrow_27)
        .Formula = "=VALUE(SUBSTITUTE(RC[-1],"" "", """"))"
    End With
    Columns("L").Copy
    Columns("L").PasteSpecial Paste:=xlPasteValues
    
    'Buscarv del salario entre la hoja 929 y la 27
    Sheets("ZHR929 FLEX Y CCTO").Activate
    With Worksheets("ZHR929 FLEX Y CCTO").Range("I2:I" & lastRow)
        .Formula = "=VLOOKUP(RC[-8],'ZHRMX27 FECHA DEL REING'!C[-8]:C[3],12,0)"
    End With
    Columns("I").Copy
    Columns("I").PasteSpecial Paste:=xlPasteValues
    
    
    ActiveWorkbook.Save 'Guardar
    
End Sub

Sub Filtrar()
'Elimina los datos del a�o anterior al que se esta realizando el informe en las hojas PTU DIAS y PTU PESOS
    
    InicializarVariables
    

    'ELIMINA DATOS DEL A�O ANTERIOR EN LA PTU DIAS DETALLADO

        Workbooks("PW PTU " & A�o & ".xlsx").Activate
        Sheets("PTU DIAS DETALLADO").Activate
        Columns("U:U").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    
        'Llena la nueva columna con los digitos del a�o segun la columna Per.para
        Range("U2:U" & Cells(Rows.Count, "V").End(xlUp).Row).Formula = "=LEFT(V2,4)"
        Columns("U:U").Copy
        Columns("U:U").PasteSpecial Paste:=xlPasteValues
        Application.CutCopyMode = False
    
        'Evalua
        Dim celda As Range
        For Each celda In Range("U2:U" & Cells(Rows.Count, "U").End(xlUp).Row)
            If Left(celda.Value, 4) = CStr(A�o - 1) Then
                celda.Interior.Color = RGB(173, 216, 230)
            End If
        Next celda
    
        ' Eliminar las filas que est�n resaltadas
        On Error Resume Next
        Application.ScreenUpdating = False
        
        Dim y As Long
        Dim lastRow_y As Long
        
        lastRow_y = Cells(Rows.Count, "U").End(xlUp).Row
        
        For y = lastRow_y To 2 Step -1
            If Cells(y, "U").Interior.Color = RGB(173, 216, 230) Then
                Cells(y, "U").EntireRow.Delete
            End If
        Next y
        
        Application.ScreenUpdating = True
        
        Columns("U:U").Delete 'Elimina la columna que creamos
    
    
    'ELIMINA DATOS DEL A�O ANTERIOR EN LA PTU PESOS DETALLADO

        Workbooks("PW PTU " & A�o & ".xlsx").Activate
        Sheets("PTU PESOS DETALLADO").Activate
        Columns("W:W").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    
        ' Aplicar f�rmula para extraer los primeros 4 d�gitos de la columna V y pegar como valores
        Range("W2:W" & Cells(Rows.Count, "X").End(xlUp).Row).Formula = "=LEFT(X2,4)"
        Columns("W:W").Copy
        Columns("W:W").PasteSpecial Paste:=xlPasteValues
        Application.CutCopyMode = False ' Limpiar el portapapeles
    
        'Evalua
        Dim celda1 As Range
        For Each celda1 In Range("W2:W" & Cells(Rows.Count, "W").End(xlUp).Row)
            If Left(celda1.Value, 4) = CStr(A�o - 1) Then
                celda1.Interior.Color = RGB(173, 216, 230) ' Color azul claro
            End If
        Next celda1
    
        ' Eliminar las filas que est�n resaltadas
        
        On Error Resume Next
        Application.ScreenUpdating = False
        
        Dim z As Long
        Dim lastRow_z As Long
        
        lastRow_z = Cells(Rows.Count, "W").End(xlUp).Row
        
        For z = lastRow_z To 2 Step -1
            If Cells(z, "W").Interior.Color = RGB(173, 216, 230) Then
                Cells(z, "W").EntireRow.Delete
            End If
        Next z
        
        Application.ScreenUpdating = True
        
        Columns("W:W").Delete 'Elimina la columna que creamos
        
        ActiveWorkbook.Save 'Guarda
    

End Sub

Sub TD_Dias()
'Crea la tabla dinamica de la hoja PTU DIAS DETALLADO entre los siguientes campos: N personal, Sociedad e Cantidad

    InicializarVariables
    Dim ultimaFila As Long
    ultimaFila = Cells(Rows.Count, "A").End(xlUp).Row
    
    Workbooks("PW PTU " & A�o & ".xlsx").Activate 'Activa la hoja
    Sheets("PTU DIAS DETALLADO").Activate
    
    'Rango
    Dim rangoDatos As Range
    Set rangoDatos = Range("A1:AB" & ultimaFila)
    
    'Ubicacion de la tabla dinamica
    Dim celdaTabla As Range
    Set celdaTabla = Range("AE1")
    
    'Creacion de la tabla dinamica
    Dim tablaDinamica As PivotTable
    Set tablaDinamica = ActiveSheet.PivotTableWizard(SourceType:=xlDatabase, SourceData:=rangoDatos, TableDestination:=celdaTabla)
    
    With tablaDinamica
        .PivotFields("N� pers.").Orientation = xlRowField '
        .PivotFields("Soc.").Orientation = xlColumnField
        .AddDataField .PivotFields("Ctd."), "Suma de Ctd.", xlSum
        '.PivotFields("Ctd.").Orientation = xlDataField
    End With
    
    With tablaDinamica.PivotFields("Soc.")
    On Error Resume Next
    .PivotItems("(blank)").Visible = False
    On Error GoTo 0
    End With
    
    tablaDinamica.TableStyle2 = ""
    ActiveWorkbook.Save 'Guardar
    
End Sub

Sub TD_Pesos()
'Crea la tabla dinamica de la hoja PTU PESOS DETALLADO entre los siguientes campos: N personal, Sociedad e Importe
    
    InicializarVariables
    Dim ultimaFila2 As Long
    ultimaFila2 = Cells(Rows.Count, "A").End(xlUp).Row
    
    Workbooks("PW PTU " & A�o & ".xlsx").Activate 'Activar la hoja
    Sheets("PTU PESOS DETALLADO").Activate
    
    'Rango
    Dim rangoDatos As Range
    Set rangoDatos = Range("A1:AD" & ultimaFila2)
    
    'Ubicacion de la tabla dinamica
    Dim celdaTabla As Range
    Set celdaTabla = Range("AF1")
    
    'Creacion de la tabla dinamica
    Dim tablaDinamica As PivotTable
    Set tablaDinamica = ActiveSheet.PivotTableWizard(SourceType:=xlDatabase, SourceData:=rangoDatos, TableDestination:=celdaTabla)
    
    With tablaDinamica
        .PivotFields("N� pers.").Orientation = xlRowField '
        .PivotFields("Soc.").Orientation = xlColumnField
        .AddDataField .PivotFields("Importe"), "Suma de Ctd.", xlSum
        '.PivotFields("Importe").Orientation = xlDataField
    End With
    
    tablaDinamica.TableStyle2 = ""
    ActiveWorkbook.Save 'Guardar
    
End Sub

Sub TD_ausentismos()
'Crea la tabla dinamica de la hoja AUSENTISMOS entre los siguientes campos: N personal, Sociedad e Cantidad

    InicializarVariables
    Dim ultimaFila3 As Long
    ultimaFila3 = Cells(Rows.Count, "A").End(xlUp).Row
    
    Workbooks("PW PTU " & A�o & ".xlsx").Activate
    Sheets("AUSENTISMOS SIND").Activate
    
    'Rango
    Dim rangoDatos As Range
    Set rangoDatos = Range("A1:AE" & ultimaFila3)
    
    'Ubicacion de la tabla
    Dim celdaTabla As Range
    Set celdaTabla = Range("AG1")
    
    'Creacion de tabla dinamica
    Dim tablaDinamica As PivotTable
    Set tablaDinamica = ActiveSheet.PivotTableWizard(SourceType:=xlDatabase, SourceData:=rangoDatos, TableDestination:=celdaTabla)
    
    With tablaDinamica
        .PivotFields("N� pers.").Orientation = xlRowField
        .PivotFields("CC-n.").Orientation = xlColumnField
        .PivotFields("Texto expl.CC-n�mina").Orientation = xlColumnField
        .AddDataField .PivotFields("Ctd."), "Suma de Ctd.", xlSum
        '.PivotFields("Ctd.").Orientation = xlDataField
    End With
    
    tablaDinamica.TableStyle2 = ""
    ActiveWorkbook.Save 'Guardar
    
End Sub

Sub Buscarv()

    InicializarVariables
    Dim lastRow As Long


    'Realiza el buscarV para PTU DIAS
    Workbooks("PW PTU " & A�o & ".xlsx").Activate
    Sheets("ZHR929 FLEX Y CCTO").Select
    lastRow = ActiveSheet.Cells(Rows.Count, 1).End(xlUp).Row
    With Worksheets("ZHR929 FLEX Y CCTO").Range("S2:S" & lastRow)
    .Formula = "=VLOOKUP(RC[-18], 'PTU DIAS DETALLADO'!C[12]:C[18], 6, 0)"
    End With
    
        'Copia y pega como valores para quitar las formulas
        Columns("S").Copy
        Columns("S").PasteSpecial Paste:=xlPasteValues
        Application.CutCopyMode = False
        
        
    'Realizar buscarv Para PTU PESOS
    With Worksheets("ZHR929 FLEX Y CCTO").Range("V2:V" & lastRow)
    .Formula = "=IF(ISERROR(VLOOKUP(RC[-21], 'PTU PESOS DETALLADO'!C[10]:C[18], 9, 0)),"""",VLOOKUP(RC[-21], 'PTU PESOS DETALLADO'!C[10]:C[18], 9, 0))"
    End With
    
        'Copia y pega como valores para quitar las formulas
        Columns("V").Copy
        Columns("V").PasteSpecial Paste:=xlPasteValues
        Application.CutCopyMode = False

    
End Sub

Sub Filtro_ZHR929()
'Filtra columnas H y S para poner observaciones segun el caso
'Crea el condicional "Si" o "No" segun la observacion en la columna W
'Realiza un filtro final para extraer los empleados Sindicalizados y con PTU

    InicializarVariables
    Dim UltFila As Long
    UltFila = Cells(Rows.Count, "A").End(xlUp).Row
    Dim celda As Range
    Dim i As Integer

    
    'Filro

    Workbooks("PW PTU " & A�o & ".xlsx").Activate
    Sheets("ZHR929 FLEX Y CCTO").Activate
    Rows("1:1").AutoFilter
    ActiveSheet.Range("H:H").AutoFilter Field:=8, Criteria1:="<>TIEMPO INDETERMINADO"
    ActiveSheet.Range("S:S").AutoFilter Field:=19, Criteria1:="<60"

    'Coloca la observacion
    
    For Each celda In Range("W2:W" & UltFila).SpecialCells(xlCellTypeVisible)
        celda.Value = "NO TIENE DERECHO A PTU <60 CON CTTO DETERMINADO"
    Next celda
    ActiveSheet.AutoFilterMode = False 'Quitar filtros
    
    'Crea una columna nueva
    Columns("W:W").Insert Shift:=xlToLeft, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("W1").Value = "TIENE O NO DERECHO A PTU"
    
    
    'Escribe el condicional "Si" o "No"
        For i = 2 To UltFila
            ' Verifica si la celda en la columna X contiene la frase especificada
            If InStr(1, Cells(i, "X").Value, "NO TIENE DERECHO A PTU <60 CON CTTO DETERMINADO") > 0 Then
                Cells(i, "W").Value = "NO"
            Else
                Cells(i, "W").Value = " "
            End If
        Next i
    
    ActiveWorkbook.Save 'Guarda
    

End Sub

Sub Unificado_y_planta()
'Crea el nuevo informe "Informe Dias Pesos A�o"



    InicializarVariables
    Dim UltFila As Long
    Workbooks("PW PTU " & A�o & ".xlsx").Activate
    UltFila = Cells(Rows.Count, "A").End(xlUp).Row
    
    'Filtra los empleados sindicalizados
    Sheets("ZHR929 FLEX Y CCTO").Activate
    On Error Resume Next
    ActiveSheet.AutoFilterMode = False ' Elimina cualquier filtro existente
    Rows("1:1").AutoFilter
    ActiveSheet.Range("G:G").AutoFilter Field:=7, Criteria1:="SINDICALIZADO CAT"
    ActiveSheet.Range("W:W").AutoFilter Field:=23, Criteria1:="SI"
    
    On Error GoTo 0 ' Restablece el manejo de errores a su estado normal

    'Crea un libro con las hojas correspondientes
    Workbooks.Add.SaveAs ruta3 & "\" & "1. Informe Dias Pesos A�o " & CStr(A�o) & " V1.xlsx"
    Sheets("Hoja1").Activate
    ActiveSheet.Name = "UNIFICADO"
    Range("A2:E2").Merge
    Range("A2").Value = "LISTADO DE PERSONAL SINDICALIZADO PTU A�O "
    Range("A2").Font.Bold = True
    ActiveWorkbook.Sheets.Add(After:=Sheets(Sheets.Count)).Name = "MX02"
    ActiveWorkbook.Sheets.Add(After:=Sheets(Sheets.Count)).Name = "MX08"
  
    
    'Crea tabla de la hoja Unificado
    Workbooks("PW PTU " & A�o & ".xlsx").Activate
    Sheets("ZHR929 FLEX Y CCTO").Activate
    Range("A:B").Resize(UltFila, 2).Copy
    
    Workbooks("1. Informe Dias Pesos A�o " & CStr(A�o) & " V1.xlsx").Activate
    Sheets("UNIFICADO").Activate
    Range("A5").PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    
    Workbooks("PW PTU " & A�o & ".xlsx").Activate
    Sheets("ZHR929 FLEX Y CCTO").Activate
    Range("G1:G" & UltFila).Copy
    
    Workbooks("1. Informe Dias Pesos A�o " & CStr(A�o) & " V1.xlsx").Activate
    Sheets("UNIFICADO").Activate
    Range("C5").PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    
    Workbooks("PW PTU " & A�o & ".xlsx").Activate
    Sheets("ZHR929 FLEX Y CCTO").Activate
    Range("M1:M" & UltFila).Copy
    
    Workbooks("1. Informe Dias Pesos A�o " & CStr(A�o) & " V1.xlsx").Activate
    Sheets("UNIFICADO").Activate
    Range("D5").PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    
    Workbooks("PW PTU " & A�o & ".xlsx").Activate
    Sheets("ZHR929 FLEX Y CCTO").Activate
    Range("U1:U" & UltFila).Copy
    
    Workbooks("1. Informe Dias Pesos A�o " & CStr(A�o) & " V1.xlsx").Activate
    Sheets("UNIFICADO").Activate
    Range("E5").PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    
    Workbooks("PW PTU " & A�o & ".xlsx").Activate
    Sheets("ZHR929 FLEX Y CCTO").Activate
    Range("V1:V" & UltFila).Copy
    
    Workbooks("1. Informe Dias Pesos A�o " & CStr(A�o) & " V1.xlsx").Activate
    Sheets("UNIFICADO").Activate
    Range("F5").PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    
    Workbooks("PW PTU " & A�o & ".xlsx").Activate
    Sheets("ZHR929 FLEX Y CCTO").Activate
    Range("J1:J" & UltFila).Copy
    
    Workbooks("1. Informe Dias Pesos A�o " & CStr(A�o) & " V1.xlsx").Activate
    Sheets("UNIFICADO").Activate
    Range("G5").PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    
    'Cambia titulos de la tabla
    Range("A5").Value = "N de personal"
    Range("B5").Value = "Nombre editado del empleado o canditato"
    Range("C5").Value = "Area de personal"
    Range("D5").Value = "Centro de trabajo"
    Range("E5").Value = "Total dias"
    Range("F5").Value = "Total pesos"
    Range("G5").Value = "CIA"
    
        With Columns("A:G")
            .AutoFit
            .HorizontalAlignment = xlCenter
        End With
        
        With Range("A5:G5")
            .Interior.Color = RGB(192, 192, 192)
            .Font.Bold = True
        End With
    
    ActiveWorkbook.Save 'Guarda
    
    'Filtrar por planta MX02
    Rows("5:5").AutoFilter
    ActiveSheet.Range("G:G").AutoFilter Field:=7, Criteria1:="MX02"
    
        'Copia y pega los datos de la primera planta en la hoja MX02
        ActiveSheet.Range("A5:G" & Cells(Rows.Count, "A").End(xlUp).Row).SpecialCells(xlCellTypeVisible).Copy
        Sheets("MX02").Activate
        Range("A5").PasteSpecial Paste:=xlPasteValues
        Application.CutCopyMode = False
        
        Range("A2:E2").Merge
        Range("A2").Value = "LISTADO DE PERSONAL SINDICALIZADO PTU A�O "
        Range("A2").Font.Bold = True
        
        With Range("A5:G5")
            .Interior.Color = RGB(192, 192, 192)
            .Font.Bold = True
        End With
        
        With Columns("A:G")
            .AutoFit
            .HorizontalAlignment = xlCenter
        End With
        
    
    'Filtra por la planta MX08
    Sheets("UNIFICADO").Activate
    ActiveSheet.AutoFilterMode = False
    Rows("5:5").AutoFilter
    ActiveSheet.Range("G:G").AutoFilter Field:=7, Criteria1:="MX08"
    
        'Copia y pega los datos de la primera planta en la hoja MX08
        ActiveSheet.Range("A5:G" & Cells(Rows.Count, "A").End(xlUp).Row).SpecialCells(xlCellTypeVisible).Copy
        Sheets("MX08").Activate
        Range("A5").PasteSpecial Paste:=xlPasteValues
        Application.CutCopyMode = False
        
        Range("A2:E2").Merge
        Range("A2").Value = "LISTADO DE PERSONAL SINDICALIZADO PTU A�O "
        Range("A2").Font.Bold = True
        
        With Range("A5:G5")
            .Interior.Color = RGB(192, 192, 192)
            .Font.Bold = True
        End With
        
        With Columns("A:G")
            .AutoFit
            .HorizontalAlignment = xlCenter
        End With
        
    ActiveWorkbook.Save 'Guarda


End Sub

Sub Sind_Alto()
'Crea el nuevo informe "1. Mayores salarios SIND"

    InicializarVariables

    'Crea el nuevo libro en la carpeta 3.
    Workbooks.Add.SaveAs ruta3 & "\" & "1. Mayores salarios SIND " & CStr(A�o) & " V1.xlsx"
    Sheets("Hoja1").Activate
    ActiveSheet.Name = "Mayores salarios " & CStr(A�o)
    Range("A2:E2").Merge
    Range("A2").Value = "LISTADO DE PERSONAL SINDICALIZADO PTU " & CStr(A�o)
    Range("A2").Font.Bold = True

    'Va a hacer el filtro para sacar los salarios mayores de MX02
    Workbooks("1. Informe Dias Pesos A�o " & CStr(A�o) & " V1.xlsx").Activate
    Sheets("MX02").Activate

    ' Encontrar la �ltima fila en la columna F
    Dim ultimaFila As Long
    ultimaFila = Cells(Rows.Count, "F").End(xlUp).Row
    Range("F5:F" & ultimaFila).Sort Key1:=Range("F5"), Order1:=xlDescending, Header:=xlYes
    
    'Copiar y pegar
    Range("A5:G8").Copy
    Workbooks("1. Mayores salarios SIND " & CStr(A�o) & " V1.xlsx").Activate
    Sheets("Mayores salarios " & CStr(A�o)).Activate
    Range("A5").PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    
       
    Workbooks("1. Informe Dias Pesos A�o " & CStr(A�o) & " V1.xlsx").Activate
    Sheets("MX08").Activate
    Dim ultimaFila2 As Long
    ultimaFila2 = Cells(Rows.Count, "F").End(xlUp).Row
    Range("F5:F" & ultimaFila2).Sort Key1:=Range("F5"), Order1:=xlDescending, Header:=xlYes
    
    Range("A6:G6").Copy
    Workbooks("1. Mayores salarios SIND " & CStr(A�o) & " V1.xlsx").Activate
    Sheets("Mayores salarios " & CStr(A�o)).Activate
    Range("A9").PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    
    With Range("A5:G5")
            .Interior.Color = RGB(192, 192, 192)
            .Font.Bold = True
    End With
        
    With Columns("A:G")
            .AutoFit
            .HorizontalAlignment = xlCenter
    End With
    
    ActiveWorkbook.Save
    ActiveWorkbook.Close
    
    'Que me cierre el 1. Informe Dias Pesos A�o
    Workbooks("1. Informe Dias Pesos A�o " & CStr(A�o) & " V1.xlsx").Activate
    ActiveWorkbook.Save
    ActiveWorkbook.Close
    
    
End Sub

Sub Plantilla_MX08()

    InicializarVariables
        
    'Verificar la carpeta�plantillas
    If Dir(ruta & "\PLANTILLAS\", vbDirectory) = "" Then
        MsgBox "La carpeta PLANTILLAS no existe en la ruta proporcionada.", vbExclamation, "Error"
        Exit Sub
    End If
    
    'Ruta de la plantilla
    Dim rutaPlantilla As String
    rutaPlantilla = ruta & "\PLANTILLAS\" & "MX0X PTU AAAA.xlsx"
    
    'Verificar que si existe el archivo de la plantilla
    If Dir(rutaPlantilla) = "" Then
        MsgBox "El archivo MX0X PTU AAAA.xlsx no existe en la carpeta PLANTILLAS.", vbExclamation, "Error"
        Exit Sub
    End If
    
    'Copiar archivo y abrirlo
    FileCopy rutaPlantilla, ruta4 & "\" & "MX08 PTU " & CStr(A�o) & ".xlsx"
    Workbooks.Open ruta4 & "\" & "MX08 PTU " & CStr(A�o) & ".xlsx", UpdateLinks:=False
    Workbooks("MX08 PTU " & CStr(A�o) & ".xlsx").Activate
    Sheets("CIA").Activate
    ActiveSheet.Name = "MX08"
    
    'Filtra los empleados
    Workbooks("PW PTU " & A�o & ".xlsx").Activate
    Sheets("ZHR929 FLEX Y CCTO").Activate
    ActiveSheet.AutoFilterMode = False
    Rows("1:1").AutoFilter
    ActiveSheet.Range("J:J").AutoFilter Field:=10, Criteria1:="MX08"
    ActiveSheet.Range("W:W").AutoFilter Field:=23, Criteria1:="SI"
    
    On Error Resume Next 'Sociedad
    Columns("J:J").Resize(Rows.Count - 1, 1).Offset(1, 0).SpecialCells(xlCellTypeVisible).Copy
    Workbooks("MX08 PTU " & CStr(A�o) & ".xlsx").Activate
    Sheets("MX08").Activate
    Range("A10").PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    On Error GoTo 0
    
    Workbooks("PW PTU " & A�o & ".xlsx").Activate 'Numero de personal
    Sheets("ZHR929 FLEX Y CCTO").Activate
    Columns("A:A").Resize(Rows.Count - 1, 1).Offset(1, 0).SpecialCells(xlCellTypeVisible).Copy
    Workbooks("MX08 PTU " & CStr(A�o) & ".xlsx").Activate
    Sheets("MX08").Activate
    Range("B10").PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False

    Workbooks("PW PTU " & A�o & ".xlsx").Activate 'Nombre del empleado
    Sheets("ZHR929 FLEX Y CCTO").Activate
    Columns("B:B").Resize(Rows.Count - 1, 1).Offset(1, 0).SpecialCells(xlCellTypeVisible).Copy
    Workbooks("MX08 PTU " & CStr(A�o) & ".xlsx").Activate
    Sheets("MX08").Activate
    Range("C10").PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    
    Workbooks("PW PTU " & A�o & ".xlsx").Activate 'Apres
    Sheets("ZHR929 FLEX Y CCTO").Activate
    Columns("F:F").Resize(Rows.Count - 1, 1).Offset(1, 0).SpecialCells(xlCellTypeVisible).Copy
    Workbooks("MX08 PTU " & CStr(A�o) & ".xlsx").Activate
    Sheets("MX08").Activate
    Range("D10").PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    
    Workbooks("PW PTU " & A�o & ".xlsx").Activate 'DivPres
    Sheets("ZHR929 FLEX Y CCTO").Activate
    Columns("L:L").Resize(Rows.Count - 1, 1).Offset(1, 0).SpecialCells(xlCellTypeVisible).Copy
    Workbooks("MX08 PTU " & CStr(A�o) & ".xlsx").Activate
    Sheets("MX08").Activate
    Range("E10").PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    
    Workbooks("PW PTU " & A�o & ".xlsx").Activate 'TextoDivPers
    Sheets("ZHR929 FLEX Y CCTO").Activate
    Columns("M:M").Resize(Rows.Count - 1, 1).Offset(1, 0).SpecialCells(xlCellTypeVisible).Copy
    Workbooks("MX08 PTU " & CStr(A�o) & ".xlsx").Activate
    Sheets("MX08").Activate
    Range("F10").PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    
    Workbooks("PW PTU " & A�o & ".xlsx").Activate 'Status de ocupaci�n
    Sheets("ZHR929 FLEX Y CCTO").Activate
    Columns("P:P").Resize(Rows.Count - 1, 1).Offset(1, 0).SpecialCells(xlCellTypeVisible).Copy
    Workbooks("MX08 PTU " & CStr(A�o) & ".xlsx").Activate
    Sheets("MX08").Activate
    Range("G10").PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    
    Workbooks("PW PTU " & A�o & ".xlsx").Activate 'Sueldo Base Mensual
    Sheets("ZHR929 FLEX Y CCTO").Activate
    Columns("I:I").Resize(Rows.Count - 1, 1).Offset(1, 0).SpecialCells(xlCellTypeVisible).Copy
    Workbooks("MX08 PTU " & CStr(A�o) & ".xlsx").Activate
    Sheets("MX08").Activate
    Range("H10").PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    
    Workbooks("PW PTU " & A�o & ".xlsx").Activate 'Ptu Dias
    Sheets("ZHR929 FLEX Y CCTO").Activate
    Columns("U:U").Resize(Rows.Count - 1, 1).Offset(1, 0).SpecialCells(xlCellTypeVisible).Copy
    Workbooks("MX08 PTU " & CStr(A�o) & ".xlsx").Activate
    Sheets("MX08").Activate
    Range("L10").PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    
    Workbooks("PW PTU " & A�o & ".xlsx").Activate 'Ptu pesos
    Sheets("ZHR929 FLEX Y CCTO").Activate
    Columns("V:V").Resize(Rows.Count - 1, 1).Offset(1, 0).SpecialCells(xlCellTypeVisible).Copy
    Workbooks("MX08 PTU " & CStr(A�o) & ".xlsx").Activate
    Sheets("MX08").Activate
    Range("N10").PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    
    ActiveWorkbook.Save
    ActiveWorkbook.Close

End Sub

Sub Plantilla_MX02()

    InicializarVariables
        
    'Verificar la carpeta�plantillas
    If Dir(ruta & "\PLANTILLAS\", vbDirectory) = "" Then
        MsgBox "La carpeta PLANTILLAS no existe en la ruta proporcionada.", vbExclamation, "Error"
        Exit Sub
    End If
    
    'Ruta de la plantilla
    Dim rutaPlantilla As String
    rutaPlantilla = ruta & "\PLANTILLAS\" & "MX0X PTU AAAA.xlsx"
    
    'Verificar que si existe el archivo de la plantilla
    If Dir(rutaPlantilla) = "" Then
        MsgBox "El archivo MX0X PTU AAAA.xlsx no existe en la carpeta PLANTILLAS.", vbExclamation, "Error"
        Exit Sub
    End If
    
    'Copiar archivo y abrirlo
    FileCopy rutaPlantilla, ruta4 & "\" & "MX02 PTU " & CStr(A�o) & ".xlsx"
    Workbooks.Open ruta4 & "\" & "MX02 PTU " & CStr(A�o) & ".xlsx", UpdateLinks:=False
    Workbooks("MX02 PTU " & CStr(A�o) & ".xlsx").Activate
    Sheets("CIA").Activate
    ActiveSheet.Name = "MX02"
    
    'Filtra los empleados
    Workbooks("PW PTU " & A�o & ".xlsx").Activate
    Sheets("ZHR929 FLEX Y CCTO").Activate
    ActiveSheet.AutoFilterMode = False
    Rows("1:1").AutoFilter
    ActiveSheet.Range("J:J").AutoFilter Field:=10, Criteria1:="MX02"
    ActiveSheet.Range("W:W").AutoFilter Field:=23, Criteria1:="SI"
    
    On Error Resume Next 'Sociedad
    Columns("J:J").Resize(Rows.Count - 1, 1).Offset(1, 0).SpecialCells(xlCellTypeVisible).Copy
    Workbooks("MX02 PTU " & CStr(A�o) & ".xlsx").Activate
    Sheets("MX02").Activate
    Range("A10").PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    On Error GoTo 0
    
    Workbooks("PW PTU " & A�o & ".xlsx").Activate 'Numero de personal
    Sheets("ZHR929 FLEX Y CCTO").Activate
    Columns("A:A").Resize(Rows.Count - 1, 1).Offset(1, 0).SpecialCells(xlCellTypeVisible).Copy
    Workbooks("MX02 PTU " & CStr(A�o) & ".xlsx").Activate
    Sheets("MX02").Activate
    Range("B10").PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False

    Workbooks("PW PTU " & A�o & ".xlsx").Activate 'Nombre del empleado
    Sheets("ZHR929 FLEX Y CCTO").Activate
    Columns("B:B").Resize(Rows.Count - 1, 1).Offset(1, 0).SpecialCells(xlCellTypeVisible).Copy
    Workbooks("MX02 PTU " & CStr(A�o) & ".xlsx").Activate
    Sheets("MX02").Activate
    Range("C10").PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    
    Workbooks("PW PTU " & A�o & ".xlsx").Activate 'Apres
    Sheets("ZHR929 FLEX Y CCTO").Activate
    Columns("F:F").Resize(Rows.Count - 1, 1).Offset(1, 0).SpecialCells(xlCellTypeVisible).Copy
    Workbooks("MX02 PTU " & CStr(A�o) & ".xlsx").Activate
    Sheets("MX02").Activate
    Range("D10").PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    
    Workbooks("PW PTU " & A�o & ".xlsx").Activate 'DivPres
    Sheets("ZHR929 FLEX Y CCTO").Activate
    Columns("L:L").Resize(Rows.Count - 1, 1).Offset(1, 0).SpecialCells(xlCellTypeVisible).Copy
    Workbooks("MX02 PTU " & CStr(A�o) & ".xlsx").Activate
    Sheets("MX02").Activate
    Range("E10").PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    
    Workbooks("PW PTU " & A�o & ".xlsx").Activate 'TextoDivPers
    Sheets("ZHR929 FLEX Y CCTO").Activate
    Columns("M:M").Resize(Rows.Count - 1, 1).Offset(1, 0).SpecialCells(xlCellTypeVisible).Copy
    Workbooks("MX02 PTU " & CStr(A�o) & ".xlsx").Activate
    Sheets("MX02").Activate
    Range("F10").PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    
    Workbooks("PW PTU " & A�o & ".xlsx").Activate 'Status de ocupaci�n
    Sheets("ZHR929 FLEX Y CCTO").Activate
    Columns("P:P").Resize(Rows.Count - 1, 1).Offset(1, 0).SpecialCells(xlCellTypeVisible).Copy
    Workbooks("MX02 PTU " & CStr(A�o) & ".xlsx").Activate
    Sheets("MX02").Activate
    Range("G10").PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    
    Workbooks("PW PTU " & A�o & ".xlsx").Activate 'Sueldo Base Mensual
    Sheets("ZHR929 FLEX Y CCTO").Activate
    Columns("I:I").Resize(Rows.Count - 1, 1).Offset(1, 0).SpecialCells(xlCellTypeVisible).Copy
    Workbooks("MX02 PTU " & CStr(A�o) & ".xlsx").Activate
    Sheets("MX02").Activate
    Range("H10").PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    
    Workbooks("PW PTU " & A�o & ".xlsx").Activate 'Ptu Dias
    Sheets("ZHR929 FLEX Y CCTO").Activate
    Columns("U:U").Resize(Rows.Count - 1, 1).Offset(1, 0).SpecialCells(xlCellTypeVisible).Copy
    Workbooks("MX02 PTU " & CStr(A�o) & ".xlsx").Activate
    Sheets("MX02").Activate
    Range("L10").PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    
    Workbooks("PW PTU " & A�o & ".xlsx").Activate 'Ptu pesos
    Sheets("ZHR929 FLEX Y CCTO").Activate
    Columns("V:V").Resize(Rows.Count - 1, 1).Offset(1, 0).SpecialCells(xlCellTypeVisible).Copy
    Workbooks("MX02 PTU " & CStr(A�o) & ".xlsx").Activate
    Sheets("MX02").Activate
    Range("N10").PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    
    ActiveWorkbook.Save
    ActiveWorkbook.Close
    
    'Me guarde el archivo PW PTU
    Workbooks("PW PTU " & A�o & ".xlsx").Activate 'Ptu Dias
    ActiveSheet.AutoFilterMode = False
    ActiveWorkbook.Save
    ActiveWorkbook.Close
    
End Sub

