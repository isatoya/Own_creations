Attribute VB_Name = "BSWIFT"
Option Explicit
Dim wb, wbBase, wbPlantilla_BSWIFT, wbFiltrado, wbSTD As Workbook
Dim ws As Worksheet
Dim Fecha1, Fecha2, mes, Mes_Texto, año, dia, fecha_hoy, Ruta, Ruta_pais, Ruta_Año, Ruta_Audi, Ruta_Base, texto, ExisteArchivo As String
Dim hoja As Integer
Dim lastRow, lastRowA, lastRowH, lastCol, lastRowSAP, i As Long
Dim titles As Variant
Dim rng, cell As Range
    
Sub Ejecutar_BSWIFT()

' Verificar si hay datos en las celdas I8 y M8
'If ThisWorkbook.Sheets("Home Page").Range("I8").Value = "" Or ThisWorkbook.Sheets("Home Page").Range("M8").Value = "" Then
'    MsgBox "Incomplete data, please enter the data before executing.", vbExclamation
'    Exit Sub
'End If

' Llama a cada una de las funciones
InicializarVariables
CrearCarpetas
Desarrollo_BSWIFT

MsgBox "Complete audit. Please go to the path and verify the documents.", vbInformation

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
fecha_hoy = Format(Date, "mm.dd.yyyy")

'Rutas
Ruta = ThisWorkbook.Path & "\"
Ruta_pais = Ruta & "Benefits audits"
Ruta_Audi = Ruta_pais & "\" & "Bswift VS SAP"
'Ruta_Año = Ruta_Audi & "\" & año

End Sub
Sub CrearCarpetas()

'Carpeta pais
Ruta_pais = Ruta & "Benefits audits"
If Dir(Ruta_pais, vbDirectory + vbHidden) = "" Then
    If Dir(Ruta_pais & vbDirectory + vbHidden) = "" Then MkDir Ruta_pais
End If

'Carpeta auditorias
Ruta_Audi = Ruta_pais & "\" & "Bswift VS SAP"
If Dir(Ruta_Audi, vbDirectory + vbHidden) = "" Then
    If Dir(Ruta_Audi & vbDirectory + vbHidden) = "" Then MkDir Ruta_Audi
End If

'Carpeta año
'Ruta_Año = Ruta_Audi & "\" & año
'If Dir(Ruta_Año, vbDirectory + vbHidden) = "" Then
'    If Dir(Ruta_Audi & vbDirectory + vbHidden) = "" Then MkDir Ruta_Año
'End If
     
End Sub

Sub Desarrollo_BSWIFT()

'****************** 1) CREA COPIA DEL A PLANTILLA ******************

'Hace copia de la plantilla
On Error Resume Next
Kill Ruta_Audi & "\" & fecha_hoy & " Bswift vs SAP Audit" & ".xlsx"
On Error GoTo 0

Set wbPlantilla_BSWIFT = Workbooks.Open(Ruta & "\" & "Template Bswift vs SAP Audit.xlsx")
wbPlantilla_BSWIFT.Activate
ActiveWorkbook.SaveCopyAs Filename:=Ruta_Audi & "\" & fecha_hoy & " Bswift vs SAP Audit" & ".xlsx"
wbPlantilla_BSWIFT.Close False

'Crear un nuevo libro que es donde va a poner todo lo que separa
On Error Resume Next
Kill Ruta_Audi & "\" & fecha_hoy & " Filtered Bswift" & ".xlsx"
On Error GoTo 0
Set wbFiltrado = Workbooks.Add
wbFiltrado.SaveAs Ruta_Audi & "\" & fecha_hoy & " Filtered Bswift" & ".xlsx"
wbFiltrado.Sheets(1).Name = "Terminated"
Sheets.Add(After:=Sheets(Sheets.Count)).Name = "EE missing set up"
Sheets.Add(After:=Sheets(Sheets.Count)).Name = "EE discrepancies"
Sheets.Add(After:=Sheets(Sheets.Count)).Name = "ER missing set up"
Sheets.Add(After:=Sheets(Sheets.Count)).Name = "ER discrepancies"
Sheets.Add(After:=Sheets(Sheets.Count)).Name = "Tabaco"
Sheets.Add(After:=Sheets(Sheets.Count)).Name = "STD"
wbFiltrado.Save
wbFiltrado.Close

'Crear un nuevo libro de STD
On Error Resume Next
Kill Ruta_Audi & "\" & fecha_hoy & " STD" & ".xlsx"
On Error GoTo 0
Set wbSTD = Workbooks.Add
wbSTD.SaveAs Ruta_Audi & "\" & fecha_hoy & " STD" & ".xlsx"
wbSTD.Sheets(1).Name = "Original"
Sheets.Add(After:=Sheets(Sheets.Count)).Name = "Tabla dinamica"
Sheets.Add(After:=Sheets(Sheets.Count)).Name = "SAP-Bswift"
wbSTD.Save
wbSTD.Close

'Abre el documento de la plantilla
Set wbPlantilla_BSWIFT = Workbooks.Open(Ruta_Audi & "\" & fecha_hoy & " Bswift vs SAP Audit" & ".xlsx")

'****************** 2) PIDE AL USUARIO SELECCIONAR EL ACHIVO DESCARGADO DE BSWIFT ******************

'Solcita al usuario abrir el doc previamente descargado de BSWIFT
MsgBox "Please select the BSWIFT database file downloaded", vbInformation
Ruta_Base = Application.GetOpenFilename("Archivos Excel (*.xls; *.xlsx), *.xls; *.xlsx")
Application.AskToUpdateLinks = False
    
' Abre el reporte seleccionado por el usuario. Luego, copia y pega la info a la plantilla
If Ruta_Base <> "Falso" Then
    
        Set wbBase = Workbooks.Open(Filename:=Ruta_Base, UpdateLinks:=0)
        wbBase.Activate
        Worksheets(1).Activate
        lastRow = Cells(Rows.Count, "A").End(xlUp).row
        Range("A2:J" & lastRow).Select
        Selection.Copy
        wbPlantilla_BSWIFT.Activate
        Sheets("Bswift-SAP").Activate
        Sheets("Bswift-SAP").Range("A2").PasteSpecial Paste:=xlPasteAll
        Columns("A:J").AutoFit
        Application.CutCopyMode = False
        wbBase.Save
        wbBase.Close
           
    Else
       
End If


'****************** 3) PIDE AL USUARIO SELECCIONAR EL ACHIVO DESCARGADO DE BSWIFT DE HBE******************

'Solcita al usuario abrir el doc previamente descargado de BSWIFT de HBE
MsgBox "Please select the BSWIFT (HBE) database file downloaded", vbInformation
Ruta_Base = Application.GetOpenFilename("Archivos Excel (*.xls; *.xlsx; *.xlsb), *.xls; *.xlsx; *.xlsb")
Application.AskToUpdateLinks = False

' Abre el reporte seleccionado por el usuario. Luego, copia y pega la info a la plantilla
If Ruta_Base <> "Falso" Then
    
        Set wbBase = Workbooks.Open(Filename:=Ruta_Base, UpdateLinks:=0)
        wbBase.Activate
        Worksheets(1).Activate
        
        'Filtra para que solo quede /BT1 y 2995
        If ActiveSheet.AutoFilterMode Then
            ActiveSheet.AutoFilterMode = False
        End If
        lastRow = Cells(Rows.Count, "A").End(xlUp).row
        Rows("1:1").AutoFilter
        Rows("1:1").AutoFilter Field:=10, Criteria1:="=/BT1", Operator:=xlOr, Criteria2:="2995"
        
        'Empieza a pasar datos a la plantilla
        Range("E2:E" & lastRow).SpecialCells(xlCellTypeVisible).Copy  'Pasa sap ID
        wbPlantilla_BSWIFT.Sheets("Bswift-SAP").Activate
        lastRow = Cells(Rows.Count, "A").End(xlUp).row
        Cells(lastRow + 1, "A").PasteSpecial Paste:=xlPasteAll
        Application.CutCopyMode = False
        
        wbBase.Activate
        Worksheets(1).Activate
        lastRow = Cells(Rows.Count, "A").End(xlUp).row
        Range("G2:G" & lastRow).SpecialCells(xlCellTypeVisible).Copy 'Pasa first name
        wbPlantilla_BSWIFT.Sheets("Bswift-SAP").Activate
        lastRow = Cells(Rows.Count, "C").End(xlUp).row
        Cells(lastRow + 1, "C").PasteSpecial Paste:=xlPasteAll
        Application.CutCopyMode = False
        
        wbBase.Activate
        Worksheets(1).Activate
        lastRow = Cells(Rows.Count, "H").End(xlUp).row
        Range("H2:H" & lastRow).SpecialCells(xlCellTypeVisible).Copy 'Pasa last name
        wbPlantilla_BSWIFT.Sheets("Bswift-SAP").Activate
        lastRow = Cells(Rows.Count, "B").End(xlUp).row
        Cells(lastRow + 1, "B").PasteSpecial Paste:=xlPasteAll
        Application.CutCopyMode = False
        
        wbBase.Activate
        Worksheets(1).Activate
        lastRow = Cells(Rows.Count, "K").End(xlUp).row
        Range("K2:K" & lastRow).SpecialCells(xlCellTypeVisible).Copy 'Benefit Plan Type
        wbPlantilla_BSWIFT.Sheets("Bswift-SAP").Activate
        lastRow = Cells(Rows.Count, "D").End(xlUp).row
        Cells(lastRow + 1, "D").PasteSpecial Paste:=xlPasteAll
        Application.CutCopyMode = False
        
        wbBase.Activate
        Worksheets(1).Activate
        lastRow = Cells(Rows.Count, "N").End(xlUp).row
        Range("N2:N" & lastRow).SpecialCells(xlCellTypeVisible).Copy 'EE cost per PP 1
        wbPlantilla_BSWIFT.Sheets("Bswift-SAP").Activate
        lastRow = Cells(Rows.Count, "F").End(xlUp).row
        Cells(lastRow + 1, "F").PasteSpecial Paste:=xlPasteAll
        Application.CutCopyMode = False
        
        wbBase.Activate
        Worksheets(1).Activate
        lastRow = Cells(Rows.Count, "I").End(xlUp).row
        Range("I2:I" & lastRow).SpecialCells(xlCellTypeVisible).Copy 'Pay Frequency (No Codes)
        wbPlantilla_BSWIFT.Sheets("Bswift-SAP").Activate
        lastRow = Cells(Rows.Count, "I").End(xlUp).row
        Cells(lastRow + 1, "I").PasteSpecial Paste:=xlPasteAll
        Application.CutCopyMode = False
        
        wbBase.Activate
        Worksheets(1).Activate
        lastRow = Cells(Rows.Count, "M").End(xlUp).row
        Range("M2:M" & lastRow).SpecialCells(xlCellTypeVisible).Copy 'Effective date
        wbPlantilla_BSWIFT.Sheets("Bswift-SAP").Activate
        lastRow = Cells(Rows.Count, "E").End(xlUp).row
        Cells(lastRow + 1, "E").PasteSpecial Paste:=xlPasteAll
        Application.CutCopyMode = False
        
        'En la columna H se pone 0 a todo
        wbPlantilla_BSWIFT.Sheets("Bswift-SAP").Activate
        lastRowA = Cells(Rows.Count, "A").End(xlUp).row
        lastRowH = Cells(Rows.Count, "J").End(xlUp).row
        Range("H" & lastRowH + 1 & ":H" & lastRowA).Value = 0
        
        wbPlantilla_BSWIFT.Save
        
        
        'Pasa la otra parte del filtro: Filtra para que solo quede 5123, 5515 , 5516
        wbBase.Activate
        Worksheets(1).Activate
        If ActiveSheet.AutoFilterMode Then
            ActiveSheet.AutoFilterMode = False
        End If
        lastRow = Cells(Rows.Count, "A").End(xlUp).row
        Rows("1:1").AutoFilter
        Rows("1:1").AutoFilter Field:=10, Criteria1:=Array("5123", "5515", "5516"), Operator:=xlFilterValues
        
        'Empieza a pasar datos a la plantilla
        Range("E2:E" & lastRow).SpecialCells(xlCellTypeVisible).Copy  'Pasa sap ID
        wbPlantilla_BSWIFT.Sheets("Bswift-SAP").Activate
        lastRow = Cells(Rows.Count, "A").End(xlUp).row
        Cells(lastRow + 1, "A").PasteSpecial Paste:=xlPasteAll
        Application.CutCopyMode = False
        
        wbBase.Activate
        Worksheets(1).Activate
        lastRow = Cells(Rows.Count, "A").End(xlUp).row
        Range("G2:G" & lastRow).SpecialCells(xlCellTypeVisible).Copy 'Pasa first name
        wbPlantilla_BSWIFT.Sheets("Bswift-SAP").Activate
        lastRow = Cells(Rows.Count, "C").End(xlUp).row
        Cells(lastRow + 1, "C").PasteSpecial Paste:=xlPasteAll
        Application.CutCopyMode = False
        
        wbBase.Activate
        Worksheets(1).Activate
        lastRow = Cells(Rows.Count, "H").End(xlUp).row
        Range("H2:H" & lastRow).SpecialCells(xlCellTypeVisible).Copy 'Pasa last name
        wbPlantilla_BSWIFT.Sheets("Bswift-SAP").Activate
        lastRow = Cells(Rows.Count, "B").End(xlUp).row
        Cells(lastRow + 1, "B").PasteSpecial Paste:=xlPasteAll
        Application.CutCopyMode = False
        
        wbBase.Activate
        Worksheets(1).Activate
        lastRow = Cells(Rows.Count, "K").End(xlUp).row
        Range("K2:K" & lastRow).SpecialCells(xlCellTypeVisible).Copy 'Benefit Plan Type
        wbPlantilla_BSWIFT.Sheets("Bswift-SAP").Activate
        lastRow = Cells(Rows.Count, "D").End(xlUp).row
        Cells(lastRow + 1, "D").PasteSpecial Paste:=xlPasteAll
        Application.CutCopyMode = False
        
        wbBase.Activate
        Worksheets(1).Activate
        lastRow = Cells(Rows.Count, "N").End(xlUp).row
        Range("N2:N" & lastRow).SpecialCells(xlCellTypeVisible).Copy 'EE cost per PP 1
        wbPlantilla_BSWIFT.Sheets("Bswift-SAP").Activate
        lastRow = Cells(Rows.Count, "H").End(xlUp).row
        Cells(lastRow + 1, "H").PasteSpecial Paste:=xlPasteAll
        Application.CutCopyMode = False
        
        wbBase.Activate
        Worksheets(1).Activate
        lastRow = Cells(Rows.Count, "I").End(xlUp).row
        Range("I2:I" & lastRow).SpecialCells(xlCellTypeVisible).Copy 'Pay Frequency (No Codes)
        wbPlantilla_BSWIFT.Sheets("Bswift-SAP").Activate
        lastRow = Cells(Rows.Count, "I").End(xlUp).row
        Cells(lastRow + 1, "I").PasteSpecial Paste:=xlPasteAll
        Application.CutCopyMode = False
        
        wbBase.Activate
        Worksheets(1).Activate
        lastRow = Cells(Rows.Count, "M").End(xlUp).row
        Range("M2:M" & lastRow).SpecialCells(xlCellTypeVisible).Copy 'Effective date
        wbPlantilla_BSWIFT.Sheets("Bswift-SAP").Activate
        lastRow = Cells(Rows.Count, "E").End(xlUp).row
        Cells(lastRow + 1, "E").PasteSpecial Paste:=xlPasteAll
        Application.CutCopyMode = False
        
        'En la columna H se pone 0 a todo
        wbPlantilla_BSWIFT.Sheets("Bswift-SAP").Activate
        lastRowA = Cells(Rows.Count, "A").End(xlUp).row
        lastRowH = Cells(Rows.Count, "F").End(xlUp).row
        Range("F" & lastRowH + 1 & ":F" & lastRowA).Value = 0
        
        wbBase.Close SaveChanges:=False
        wbPlantilla_BSWIFT.Save
          
    Else
       
End If


'Redondea columnas F
wbPlantilla_BSWIFT.Sheets("Bswift-SAP").Activate
Columns("G:G").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
lastRow = Cells(Rows.Count, "A").End(xlUp).row
Range("G1").Value = "EE cost per PP 1"
Sheets("Bswift-SAP").Range("G2:G" & lastRow) = "=ROUND(RC[-1],2)"
Range("G:G").Copy
Range("G:G").PasteSpecial Paste:=xlPasteValues
Application.CutCopyMode = False
Columns("F:F").Delete

'Redondea columnas H
wbPlantilla_BSWIFT.Sheets("Bswift-SAP").Activate
Columns("I:I").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
lastRow = Cells(Rows.Count, "A").End(xlUp).row
Range("I1").Value = "ER Cost Per PP"
Sheets("Bswift-SAP").Range("I2:I" & lastRow) = "=ROUND(RC[-1],2)"
Range("I:I").Copy
Range("I:I").PasteSpecial Paste:=xlPasteValues
Application.CutCopyMode = False
Columns("H:H").Delete

'Formato de las columnas
Range("F:F").NumberFormat = "0.00"
Range("H:H").NumberFormat = "0.00"
Range("J:J").NumberFormat = "mm/dd/yyyy"
Cells.Borders.LineStyle = xlNone
Range("A:D").HorizontalAlignment = xlLeft
Range("E:J").HorizontalAlignment = xlCenter


'****************** 4) DESCARGA DOCUMENTOS DE SAP (PRIMER REPORTE) ******************

'Descarga reportes de sap
Dim SapGuiAuto As Object
Dim App As Object
Dim Connection As Object
Dim session As Object
    
Application.DisplayAlerts = False
Set SapGuiAuto = GetObject("SAPGUI")
Set App = SapGuiAuto.GetScriptingEngine
Set Connection = App.Children(0)
Set session = Connection.Children(0)

'Vuelve a la pantalla inicial
session.findById("wnd[0]/tbar[0]/okcd").Text = "/n"
session.findById("wnd[0]").sendVKey 0

'Entramos a la transaccion
session.findById("wnd[0]/tbar[0]/okcd").Text = "/nSQ01"
session.findById("wnd[0]").sendVKey 0

'Confirmamos el Environment
session.findById("wnd[0]/mbar/menu[5]/menu[0]").Select
session.findById("wnd[1]/usr/radRAD1").Select
session.findById("wnd[1]/tbar[0]/btn[2]").press

'Selecciona el grupo de usarios
session.findById("wnd[0]/tbar[1]/btn[19]").press
session.findById("wnd[1]/usr/cntlGRID1/shellcont/shell").selectColumn "DBGBNUM"
session.findById("wnd[1]/tbar[0]/btn[29]").press
session.findById("wnd[2]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN001-LOW").Text = "HR_ALL_SITE"
session.findById("wnd[2]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/cntlGRID1/shellcont/shell").selectedRows = "0"
session.findById("wnd[1]/tbar[0]/btn[0]").press

'Entramos al qry
session.findById("wnd[0]/usr/ctxtRS38R-QNUM").Text = "HR_SAI_0057"
session.findById("wnd[0]/tbar[1]/btn[8]").press

'Selecciona la variante HUS AUDIT
session.findById("wnd[0]/tbar[1]/btn[17]").press
session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").currentCellRow = 1
session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").selectedRows = "1"
session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").doubleClickCurrentCell

'Pega numeros de los empleados
session.findById("wnd[0]/usr/btn%_PNPPERNR_%_APP_%-VALU_PUSH").press
session.findById("wnd[1]/tbar[0]/btn[16]").press
wbPlantilla_BSWIFT.Activate
Worksheets("Bswift-SAP").Activate
lastRowSAP = ActiveSheet.Cells(Rows.Count, 1).End(xlUp).row
ActiveSheet.Range("A2:A" & lastRowSAP).Copy
session.findById("wnd[1]/tbar[0]/btn[24]").press
session.findById("wnd[1]/tbar[0]/btn[8]").press

'Ejecuta
session.findById("wnd[0]/tbar[1]/btn[8]").press

'Exporta
session.findById("wnd[0]/mbar/menu[0]/menu[3]/menu[2]").Select
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").Select
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").SetFocus
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/ctxtDY_PATH").Text = Ruta_Audi
session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = fecha_hoy & " IT0057.XLS"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 10
session.findById("wnd[1]/tbar[0]/btn[11]").press

'Organizar documento
Application.CutCopyMode = False
Workbooks.Open Ruta_Audi & "\" & fecha_hoy & " IT0057.XLS"
ActiveSheet.Cells.Select
Selection.Copy
Workbooks.Add
ActiveSheet.Paste
ActiveWorkbook.SaveAs Ruta_Audi & "\" & fecha_hoy & " IT0057.XLSX"
ActiveWorkbook.Close SaveChanges:=True
Workbooks(fecha_hoy & " IT0057.XLS").Close
Kill Ruta_Audi & "\" & fecha_hoy & " IT0057.XLS"

'Elimina filas
Set wb = Workbooks.Open(Ruta_Audi & "\" & fecha_hoy & " IT0057.XLSX")
wb.Activate
Rows("1:4").Delete
Rows("2").Delete
Columns("A").Delete

'Organiza de nuevo a mas antiguo y borra duplicados
Range("A1").Select
lastRow = Cells(Rows.Count, "A").End(xlUp).row
ActiveSheet.Range("A2:L" & lastRow).RemoveDuplicates Columns:=Array(1, 5), _
    Header:=xlYes
Range("A1:L1").Select
Selection.AutoFilter
ActiveWorkbook.Worksheets(1).AutoFilter.Sort.SortFields.Clear
ActiveWorkbook.Worksheets(1).AutoFilter.Sort.SortFields.Add Key:=Range( _
    "I1"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:= _
    xlSortNormal
With ActiveWorkbook.Worksheets(1).AutoFilter.Sort
    .Header = xlYes
    .MatchCase = False
    .Orientation = xlTopToBottom
    .SortMethod = xlPinYin
    .Apply
End With

'Copia y pega infomracion a la hoja de plantilla
lastRow = Cells(Rows.Count, "A").End(xlUp).row
Range("A2:L" & lastRow).Select
Selection.Copy
wbPlantilla_BSWIFT.Activate
Sheets("SAP-Bswift").Activate
Sheets("SAP-Bswift").Range("B2").PasteSpecial Paste:=xlPasteAll
Columns("A:M").AutoFit
wb.Save
wb.Close

'Crea la formula de la columna A
wbPlantilla_BSWIFT.Activate
Sheets("SAP-Bswift").Activate
lastRow = Cells(Rows.Count, "B").End(xlUp).row
Sheets("SAP-Bswift").Range("A2:A" & lastRow) = "=+RC[1]&RC[5]"
Range("A:A").Copy
Range("A:A").PasteSpecial Paste:=xlPasteValues
Application.CutCopyMode = False


'****************** 5) DESCARGA DOCUMENTOS DE SAP (SEGUNDO REPORTE) ******************
    
Application.DisplayAlerts = False
Set SapGuiAuto = GetObject("SAPGUI")
Set App = SapGuiAuto.GetScriptingEngine
Set Connection = App.Children(0)
Set session = Connection.Children(0)

'Vuelve a la pantalla inicial
session.findById("wnd[0]/tbar[0]/okcd").Text = "/n"
session.findById("wnd[0]").sendVKey 0

'Entramos a la transaccion
session.findById("wnd[0]/tbar[0]/okcd").Text = "/nSQ01"
session.findById("wnd[0]").sendVKey 0

'Confirmamos el Environment
session.findById("wnd[0]/mbar/menu[5]/menu[0]").Select
session.findById("wnd[1]/usr/radRAD1").Select
session.findById("wnd[1]/tbar[0]/btn[2]").press

'Selecciona el grupo de usarios
session.findById("wnd[0]/tbar[1]/btn[19]").press
session.findById("wnd[1]/usr/cntlGRID1/shellcont/shell").selectColumn "DBGBNUM"
session.findById("wnd[1]/tbar[0]/btn[29]").press
session.findById("wnd[2]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN001-LOW").Text = "HR_ALL_SITE"
session.findById("wnd[2]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/cntlGRID1/shellcont/shell").selectedRows = "0"
session.findById("wnd[1]/tbar[0]/btn[0]").press

'Entramos al qry
session.findById("wnd[0]/usr/ctxtRS38R-QNUM").Text = "HR_US_IT0000"
session.findById("wnd[0]/tbar[1]/btn[8]").press

'Selecciona la variante HUS AUDIT
session.findById("wnd[0]/tbar[1]/btn[17]").press
session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").currentCellRow = 4
session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").selectedRows = "4"
session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").doubleClickCurrentCell

'Pega numeros de los empleados
session.findById("wnd[0]/usr/btn%_PNPPERNR_%_APP_%-VALU_PUSH").press
session.findById("wnd[1]/tbar[0]/btn[16]").press
wbPlantilla_BSWIFT.Activate
Worksheets("Bswift-SAP").Activate
lastRowSAP = ActiveSheet.Cells(Rows.Count, 1).End(xlUp).row
ActiveSheet.Range("A2:A" & lastRowSAP).Copy
session.findById("wnd[1]/tbar[0]/btn[24]").press
session.findById("wnd[1]/tbar[0]/btn[8]").press

'Ejecuta
session.findById("wnd[0]/tbar[1]/btn[8]").press

'Exporta
session.findById("wnd[0]/mbar/menu[0]/menu[3]/menu[2]").Select
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").Select
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").SetFocus
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/ctxtDY_PATH").Text = Ruta_Audi
session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = fecha_hoy & " STATUS.XLS"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 10
session.findById("wnd[1]/tbar[0]/btn[11]").press

'Organizar documento
Application.CutCopyMode = False
Workbooks.Open Ruta_Audi & "\" & fecha_hoy & " STATUS.XLS"
ActiveSheet.Cells.Select
Selection.Copy
Workbooks.Add
ActiveSheet.Paste
ActiveWorkbook.SaveAs Ruta_Audi & "\" & fecha_hoy & " STATUS.XLSX"
ActiveWorkbook.Close SaveChanges:=True
Workbooks(fecha_hoy & " STATUS.XLS").Close
Kill Ruta_Audi & "\" & fecha_hoy & " STATUS.XLS"

'Elimina filas
Set wb = Workbooks.Open(Ruta_Audi & "\" & fecha_hoy & " STATUS.XLSX")
wb.Activate
Rows("1:4").Delete
Rows("2").Delete
Columns("A").Delete

'Copia y pega infomracion a la hoja de plantilla
lastRow = Cells(Rows.Count, "A").End(xlUp).row
Range("A2:W" & lastRow).Select
Selection.Copy
wbPlantilla_BSWIFT.Activate
Sheets("Status").Activate
Sheets("Status").Range("A2").PasteSpecial Paste:=xlPasteAll
Columns("A:W").AutoFit
wb.Save
wb.Close
wbPlantilla_BSWIFT.Save

'****************** 6) CREA LAS FORMULAS DE LA PARTE QUE ESTA EN AMARILLO ******************

'Cambia de texto a numeros la columna A
wbPlantilla_BSWIFT.Activate
Sheets("Bswift-SAP").Activate
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False

Set rng = Range("A2:A" & Cells(Rows.Count, 1).End(xlUp).row)
For Each cell In rng
    If IsNumeric(cell.Value) Then
        cell.Value = Val(cell.Value)
    End If
Next cell

Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True

'Crea las formulas
lastRow = Cells(Rows.Count, "A").End(xlUp).row
Sheets("Bswift-SAP").Range("K2:K" & lastRow) = "=IFERROR(RC[-10]&RC[1],"""")"
Sheets("Bswift-SAP").Range("L2:L" & lastRow) = "=IFERROR(VLOOKUP(RC[-8],WTs!C[-11]:C[-9],2,FALSE),"""")"
Sheets("Bswift-SAP").Range("M2:M" & lastRow) = "=IFERROR(VLOOKUP(RC[-9],WTs!C[-12]:C[-10],3,FALSE),"""")"
Sheets("Bswift-SAP").Range("N2:N" & lastRow) = "=IF(OR(RC[-2]=2511,RC[-2]=2995),IFERROR(VLOOKUP(RC[-13]&""2995"",'SAP-Bswift'!C1:C8,8,0),""""),"""")"
Sheets("Bswift-SAP").Range("O2:O" & lastRow) = "=IF(RC[-1]<>"""",IF(OR(IF(OR(RC[-3]=""No"",IFERROR(VLOOKUP(RC[-14]&RC[-3],'SAP-Bswift'!C1,1,0),0)=0),""No IT14"",IF(RC[-9]>0,ABS(IFERROR(VLOOKUP(RC[-14]&RC[-3],'SAP-Bswift'!C1:C8,8,0),0))-RC[-9],""No EE contribution""))=RC[-1],IF(OR(RC[-3]=""No"",IFERROR(VLOOKUP(RC[-14]&RC[-3],'SAP-Bswift'!C1,1,0),0)=0),""No IT14"",IF(RC[-9]>0,ABS(IFERROR(VLOOKUP(RC[-14]&RC[-3],'SAP" & _
        "-Bswift'!C1:C8,8,0),0))-RC[-9],""No EE contribution""))=RC[-1]+0.01,IF(OR(RC[-3]=""No"",IFERROR(VLOOKUP(RC[-14]&RC[-3],'SAP-Bswift'!C1,1,0),0)=0),""No IT14"",IF(RC[-9]>0,ABS(IFERROR(VLOOKUP(RC[-14]&RC[-3],'SAP-Bswift'!C1:C8,8,0),0))-RC[-9],""No EE contribution""))=RC[-1]-0.01),0,IF(OR(RC[-3]=""No"",IFERROR(VLOOKUP(RC[-14]&RC[-3],'SAP-Bswift'!C1,1,0),0)=0),""No IT14" & _
        """,IF(RC[-9]>0,ABS(IFERROR(VLOOKUP(RC[-14]&RC[-3],'SAP-Bswift'!C1:C8,8,0),0))-RC[-9],""No EE contribution""))),IF(OR(RC[-3]=""No"",IFERROR(VLOOKUP(RC[-14]&RC[-3],'SAP-Bswift'!C1,1,0),0)=0),""No IT14"",IF(RC[-9]>0,ABS(IFERROR(VLOOKUP(RC[-14]&RC[-3],'SAP-Bswift'!C1:C8,8,0),0))-RC[-9],""No EE contribution"")))" & _
        ""
Sheets("Bswift-SAP").Range("P2:P" & lastRow) = "=+IF(RC[-3]=""No"",""No ER contribution"",IFERROR((ABS(VLOOKUP(RC[-15]&RC[-3],'SAP-Bswift'!C1:C8,8,0))-RC[-8]),""No IT14""))"
Sheets("Bswift-SAP").Range("Q2:Q" & lastRow) = "=VLOOKUP(RC[-16],Status!C[-16]:C[-14],3,FALSE)"
Sheets("Bswift-SAP").Range("R2:R" & lastRow) = "=VLOOKUP(RC[-17],Status!C[-17]:C[-2],16,FALSE)"
Range("K:R").Copy
Range("K:R").PasteSpecial Paste:=xlPasteValues
Application.CutCopyMode = False

'****************** 7) PRUEBAS LOGICAS ******************
Call Filtros_Bswift

'****************** 8) TRABAJA EN EL DOCUMENTO DE STD ******************
Call STD_Bswift

'Pasa datos al documento de filtrado
Set wbFiltrado = Workbooks.Open(Ruta_Audi & "\" & fecha_hoy & " Filtered Bswift" & ".xlsx")
wbPlantilla_BSWIFT.Activate
Sheets("Bswift-SAP").Activate
If ActiveSheet.AutoFilterMode Then
    ActiveSheet.AutoFilterMode = False
End If
lastRow = Cells(Rows.Count, "A").End(xlUp).row
Rows("1:1").AutoFilter
Rows("1:1").AutoFilter Field:=4, Criteria1:="=Short term disability", Operator:=xlOr, Criteria2:="=State STD" 'Columna D
Rows("1:1").AutoFilter Field:=19, Criteria1:="=REV STD", Operator:=xlOr, Criteria2:="=#N/A"  'Columna S

'Pasa informacion al documento de filtrado
Range("A1:S" & lastRow).SpecialCells(xlCellTypeVisible).Copy
wbFiltrado.Sheets("STD").Activate
Sheets("STD").Range("A1").PasteSpecial Paste:=xlPasteAll
Columns("A:S").AutoFit
Application.CutCopyMode = False
wbFiltrado.Save
wbFiltrado.Close

'****************** 11) GUARDA TODO ******************
wbPlantilla_BSWIFT.Activate
Sheets("Bswift-SAP").Activate
If ActiveSheet.AutoFilterMode Then
    ActiveSheet.AutoFilterMode = False
End If
wbPlantilla_BSWIFT.Save
wbPlantilla_BSWIFT.Close

End Sub

Sub Filtros_Bswift()

'activa del documento y limpia filtros
Set wbFiltrado = Workbooks.Open(Ruta_Audi & "\" & fecha_hoy & " Filtered Bswift" & ".xlsx")
wbPlantilla_BSWIFT.Activate
Sheets("Bswift-SAP").Activate
If ActiveSheet.AutoFilterMode Then
    ActiveSheet.AutoFilterMode = False
End If
lastRow = Cells(Rows.Count, "A").End(xlUp).row

'---------------------A) MIN 16:10 - REVISADO Y OK
Rows("1:1").AutoFilter
Rows("1:1").AutoFilter Field:=17, Criteria1:="=Termination", Operator:=xlAnd 'Columna Q
Rows("1:1").AutoFilter Field:=15, Criteria1:="=No IT14", Operator:=xlAnd 'Columna O
Rows("1:1").AutoFilter Field:=16, Criteria1:="=No ER contribution", Operator:=xlOr, Criteria2:="=No IT14" 'Columna P
If Application.WorksheetFunction.Subtotal(103, Range("A2:A" & lastRow)) > 0 Then
    For Each cell In Range("A2:A" & lastRow).SpecialCells(xlCellTypeVisible)
        cell.Offset(0, 18).Value = "OK"
    Next cell
End If


'---------------------B) MIN 16:50 - REVISADO Y OK
'INFO ES LA QUE SE PEGA EN OTRA HOJA
If ActiveSheet.AutoFilterMode Then
    ActiveSheet.AutoFilterMode = False
End If
Rows("1:1").AutoFilter
Rows("1:1").AutoFilter Field:=19, Criteria1:="=" 'Columna S
Rows("1:1").AutoFilter Field:=17, Criteria1:="=Termination" 'Columna Q
If Application.WorksheetFunction.Subtotal(103, Range("A2:A" & lastRow)) > 0 Then
    For Each cell In Range("A2:A" & lastRow).SpecialCells(xlCellTypeVisible)
        cell.Offset(0, 18).Value = "REV - TERMINATION"
    Next cell
End If

    'Crea hoja de termination
    Range("A1:S" & lastRow).SpecialCells(xlCellTypeVisible).Copy
    wbFiltrado.Sheets("Terminated").Activate
    Sheets("Terminated").Range("A1").PasteSpecial Paste:=xlPasteAll
    Columns("A:S").AutoFit
    Application.CutCopyMode = False
    wbPlantilla_BSWIFT.Activate
    Sheets("Bswift-SAP").Activate

'---------------------C) MIN 18:01 - REVISADO Y OK
If ActiveSheet.AutoFilterMode Then
    ActiveSheet.AutoFilterMode = False
End If
Rows("1:1").AutoFilter
Rows("1:1").AutoFilter Field:=19, Criteria1:="=" 'Columna S
Rows("1:1").AutoFilter Field:=17, Criteria1:="=Layoff - hourly (active)", Operator:=xlOr, Criteria2:="=Layoff - Salary (Inactive)" 'Columna Q
Rows("1:1").AutoFilter Field:=12, Criteria1:="<>No", Operator:=xlAnd 'Columna L
Rows("1:1").AutoFilter Field:=13, Criteria1:="=No", Operator:=xlAnd 'Columna M
Rows("1:1").AutoFilter Field:=15, Criteria1:="=No IT14", Operator:=xlAnd 'Columna O
Rows("1:1").AutoFilter Field:=16, Criteria1:="=No ER contribution", Operator:=xlAnd 'Columna P
If Application.WorksheetFunction.Subtotal(103, Range("A2:A" & lastRow)) > 0 Then
    For Each cell In Range("A2:A" & lastRow).SpecialCells(xlCellTypeVisible)
        cell.Offset(0, 18).Value = "OK"
    Next cell
End If


'---------------------D) MIN 18:54 - REVISADO Y OK
If ActiveSheet.AutoFilterMode Then
    ActiveSheet.AutoFilterMode = False
End If
Rows("1:1").AutoFilter
Rows("1:1").AutoFilter Field:=19, Criteria1:="=" 'Columna S
Rows("1:1").AutoFilter Field:=17, Criteria1:="=Layoff - hourly (active)", Operator:=xlOr, Criteria2:="=Layoff - Salary (Inactive)" 'Columna Q
Rows("1:1").AutoFilter Field:=13, Criteria1:="<>No", Operator:=xlAnd 'Columna M
Rows("1:1").AutoFilter Field:=12, Criteria1:="=No", Operator:=xlAnd 'Columna L
Rows("1:1").AutoFilter Field:=15, Criteria1:="=No IT14", Operator:=xlAnd 'Columna O
Rows("1:1").AutoFilter Field:=16, Criteria1:="=0", Operator:=xlAnd 'Columna P
If Application.WorksheetFunction.Subtotal(103, Range("A2:A" & lastRow)) > 0 Then
    For Each cell In Range("A2:A" & lastRow).SpecialCells(xlCellTypeVisible)
        cell.Offset(0, 18).Value = "OK"
    Next cell
End If


'---------------------E) MIN 19:49 - REVISADO Y OK
If ActiveSheet.AutoFilterMode Then
    ActiveSheet.AutoFilterMode = False
End If
Rows("1:1").AutoFilter
Rows("1:1").AutoFilter Field:=19, Criteria1:="=" 'Columna S
Rows("1:1").AutoFilter Field:=17, Criteria1:="=Layoff - hourly (active)", Operator:=xlOr, Criteria2:="=Layoff - Salary (Inactive)" 'Columna Q
Rows("1:1").AutoFilter Field:=12, Criteria1:="<>No", Operator:=xlAnd 'Columna L
Rows("1:1").AutoFilter Field:=13, Criteria1:="<>No", Operator:=xlAnd 'Columna M
Rows("1:1").AutoFilter Field:=15, Criteria1:="=No IT14", Operator:=xlAnd 'Columna O
Rows("1:1").AutoFilter Field:=16, Criteria1:="=0", Operator:=xlAnd 'Columna P
If Application.WorksheetFunction.Subtotal(103, Range("A2:A" & lastRow)) > 0 Then
    For Each cell In Range("A2:A" & lastRow).SpecialCells(xlCellTypeVisible)
        cell.Offset(0, 18).Value = "OK"
    Next cell
End If


'---------------------F) MIN 20:55 - REVISADO Y OK
If ActiveSheet.AutoFilterMode Then
    ActiveSheet.AutoFilterMode = False
End If
Rows("1:1").AutoFilter
Rows("1:1").AutoFilter Field:=19, Criteria1:="=" 'Columna S
Rows("1:1").AutoFilter Field:=17, Criteria1:="=Layoff - hourly (active)", Operator:=xlOr, Criteria2:="=Layoff - Salary (Inactive)" 'Columna Q
Rows("1:1").AutoFilter Field:=6, Criteria1:="=0,00", Operator:=xlAnd 'Columna F
Rows("1:1").AutoFilter Field:=8, Criteria1:="=0,00", Operator:=xlAnd 'Columna H"
Rows("1:1").AutoFilter Field:=15, Criteria1:="=No IT14", Operator:=xlAnd 'Columna O"
Rows("1:1").AutoFilter Field:=16, Criteria1:="=No IT14", Operator:=xlAnd 'Columna P
If Application.WorksheetFunction.Subtotal(103, Range("A2:A" & lastRow)) > 0 Then
    For Each cell In Range("A2:A" & lastRow).SpecialCells(xlCellTypeVisible)
        cell.Offset(0, 18).Value = "OK"
    Next cell
End If


'---------------------G) MIN 21:47 - REVISAR
'LOS QUE TIENEN ALGUNA DIFERNCIA Y ESTAN EN LAYOFF
'SON LOS QUE ELLA DEJA EN BLANCO MIRAR QUE COMENTARIO SE LE DEJA DESPUES
If ActiveSheet.AutoFilterMode Then
    ActiveSheet.AutoFilterMode = False
End If
Rows("1:1").AutoFilter
Rows("1:1").AutoFilter Field:=19, Criteria1:="=" 'Columna S
Rows("1:1").AutoFilter Field:=17, Criteria1:="=Layoff - hourly (active)", Operator:=xlOr, Criteria2:="=Layoff - Salary (Inactive)" 'Columna Q
If Application.WorksheetFunction.Subtotal(103, Range("A2:A" & lastRow)) > 0 Then
    For Each cell In Range("A2:A" & lastRow).SpecialCells(xlCellTypeVisible)
        cell.Offset(0, 18).Value = " " 'REV - LAYOFF
    Next cell
End If


'---------------------H) MIN 22:06 - REVISADO Y OK
If ActiveSheet.AutoFilterMode Then
    ActiveSheet.AutoFilterMode = False
End If
Rows("1:1").AutoFilter
Rows("1:1").AutoFilter Field:=19, Criteria1:="=" 'Columna S
Rows("1:1").AutoFilter Field:=17, Criteria1:="=LOA (without pay)", Operator:=xlOr 'Columna Q
Rows("1:1").AutoFilter Field:=12, Criteria1:="<>No", Operator:=xlAnd 'Columna L
Rows("1:1").AutoFilter Field:=13, Criteria1:="=No", Operator:=xlAnd 'Columna M
Rows("1:1").AutoFilter Field:=15, Criteria1:="=No IT14", Operator:=xlAnd 'Columna O
Rows("1:1").AutoFilter Field:=16, Criteria1:="=No ER contribution", Operator:=xlAnd 'Columna P
If Application.WorksheetFunction.Subtotal(103, Range("A2:A" & lastRow)) > 0 Then
    For Each cell In Range("A2:A" & lastRow).SpecialCells(xlCellTypeVisible)
        cell.Offset(0, 18).Value = "OK"
    Next cell
End If


'---------------------I) MIN 22:55 - REVISADO OK
'Se separa en dos el codigo para poder quitar los que pertenecen en la columna M a 5121,5122 y 5123

If ActiveSheet.AutoFilterMode Then
    ActiveSheet.AutoFilterMode = False
End If
Rows("1:1").AutoFilter
Rows("1:1").AutoFilter Field:=19, Criteria1:="=" 'Columna S
Rows("1:1").AutoFilter Field:=17, Criteria1:="=LOA (without pay)", Operator:=xlOr 'Columna Q
Rows("1:1").AutoFilter Field:=13, Criteria1:="<>No", Operator:=xlAnd, Criteria2:="<5121" 'Columna M
Rows("1:1").AutoFilter Field:=12, Criteria1:="=No", Operator:=xlAnd 'Columna L
Rows("1:1").AutoFilter Field:=15, Criteria1:="=No IT14", Operator:=xlAnd 'Columna O
Rows("1:1").AutoFilter Field:=16, Criteria1:="=0", Operator:=xlAnd 'Columna P
If Application.WorksheetFunction.Subtotal(103, Range("A2:A" & lastRow)) > 0 Then
    For Each cell In Range("A2:A" & lastRow).SpecialCells(xlCellTypeVisible)
        cell.Offset(0, 18).Value = "OK"
    Next cell
End If


If ActiveSheet.AutoFilterMode Then
    ActiveSheet.AutoFilterMode = False
End If
Rows("1:1").AutoFilter
Rows("1:1").AutoFilter Field:=19, Criteria1:="=" 'Columna S
Rows("1:1").AutoFilter Field:=17, Criteria1:="=LOA (without pay)", Operator:=xlOr 'Columna Q
Rows("1:1").AutoFilter Field:=13, Criteria1:="<>No", Operator:=xlAnd, Criteria2:=">5123" 'Columna M
Rows("1:1").AutoFilter Field:=12, Criteria1:="=No", Operator:=xlAnd 'Columna L
Rows("1:1").AutoFilter Field:=15, Criteria1:="=No IT14", Operator:=xlAnd 'Columna O
Rows("1:1").AutoFilter Field:=16, Criteria1:="=0", Operator:=xlAnd 'Columna P
If Application.WorksheetFunction.Subtotal(103, Range("A2:A" & lastRow)) > 0 Then
    For Each cell In Range("A2:A" & lastRow).SpecialCells(xlCellTypeVisible)
        cell.Offset(0, 18).Value = "OK"
    Next cell
End If

'---------------------J) MIN 23:48 - REVISADO OK
'EN EL VIDEO NO DICE NADA DE SI QUITAR O NO EL "NO" DE LA M SOLO LO DE FILTRAR 5121,5122 O 5123
'Se separa en dos el codigo para poder quitar los que pertenecen en la columna M a 5121,5122 y 5123

If ActiveSheet.AutoFilterMode Then
    ActiveSheet.AutoFilterMode = False
End If
Rows("1:1").AutoFilter
Rows("1:1").AutoFilter Field:=19, Criteria1:="=" 'Columna S
Rows("1:1").AutoFilter Field:=17, Criteria1:="=LOA (without pay)", Operator:=xlOr 'Columna Q
Rows("1:1").AutoFilter Field:=12, Criteria1:="<>No", Operator:=xlAnd 'Columna L
Rows("1:1").AutoFilter Field:=13, Criteria1:="<5121", Operator:=xlAnd 'Columna M
Rows("1:1").AutoFilter Field:=15, Criteria1:="=No IT14", Operator:=xlAnd 'Columna O
Rows("1:1").AutoFilter Field:=16, Criteria1:="=0", Operator:=xlAnd 'Columna P
If Application.WorksheetFunction.Subtotal(103, Range("A2:A" & lastRow)) > 0 Then
    For Each cell In Range("A2:A" & lastRow).SpecialCells(xlCellTypeVisible)
        cell.Offset(0, 18).Value = "OK"
    Next cell
End If


If ActiveSheet.AutoFilterMode Then
    ActiveSheet.AutoFilterMode = False
End If
Rows("1:1").AutoFilter
Rows("1:1").AutoFilter Field:=19, Criteria1:="=" 'Columna S
Rows("1:1").AutoFilter Field:=17, Criteria1:="=LOA (without pay)", Operator:=xlOr 'Columna Q
Rows("1:1").AutoFilter Field:=12, Criteria1:="<>No", Operator:=xlAnd 'Columna L
Rows("1:1").AutoFilter Field:=13, Criteria1:=">5123", Operator:=xlAnd 'Columna M
Rows("1:1").AutoFilter Field:=15, Criteria1:="=No IT14", Operator:=xlAnd 'Columna O
Rows("1:1").AutoFilter Field:=16, Criteria1:="=0", Operator:=xlAnd 'Columna P

If Application.WorksheetFunction.Subtotal(103, Range("A2:A" & lastRow)) > 0 Then
    For Each cell In Range("A2:A" & lastRow).SpecialCells(xlCellTypeVisible)
        cell.Offset(0, 18).Value = "OK"
    Next cell
End If

'---------------------K) MIN 24:26 - REVISADO OK
If ActiveSheet.AutoFilterMode Then
    ActiveSheet.AutoFilterMode = False
End If
Rows("1:1").AutoFilter
Rows("1:1").AutoFilter Field:=19, Criteria1:="=" 'Columna S
Rows("1:1").AutoFilter Field:=17, Criteria1:="=LOA (without pay)", Operator:=xlOr 'Columna Q
Rows("1:1").AutoFilter Field:=6, Criteria1:="=0,00", Operator:=xlAnd 'Columna F
Rows("1:1").AutoFilter Field:=8, Criteria1:="=0,00", Operator:=xlAnd 'Columna H"
Rows("1:1").AutoFilter Field:=15, Criteria1:="=No IT14", Operator:=xlAnd 'Columna O"
Rows("1:1").AutoFilter Field:=16, Criteria1:="=No IT14", Operator:=xlAnd 'Columna P

If Application.WorksheetFunction.Subtotal(103, Range("A2:A" & lastRow)) > 0 Then
    For Each cell In Range("A2:A" & lastRow).SpecialCells(xlCellTypeVisible)
        cell.Offset(0, 18).Value = "OK"
    Next cell
End If

'---------------------L) MIN 25:16 - REVISAR QUE COMENTARIO SE PONE ACA, PORQUE ELLA LAS DEJA EN BLANCO
'REVISAR CON LESLIE SI SI ESTA BIEN PARA CUANDO TOCA REVISAR SI SE QUEDA EN BLANCO O TIENE COMENTARIO
If ActiveSheet.AutoFilterMode Then
    ActiveSheet.AutoFilterMode = False
End If
Rows("1:1").AutoFilter
Rows("1:1").AutoFilter Field:=19, Criteria1:="=" 'Columna S
Rows("1:1").AutoFilter Field:=17, Criteria1:="=LOA (without pay)", Operator:=xlOr 'Columna Q
Rows("1:1").AutoFilter Field:=13, Criteria1:="<>No", Operator:=xlAnd, Criteria2:="<5121" 'Columna M
If Application.WorksheetFunction.Subtotal(103, Range("A2:A" & lastRow)) > 0 Then
    For Each cell In Range("A2:A" & lastRow).SpecialCells(xlCellTypeVisible)
        cell.Offset(0, 18).Value = " " 'REV - LOA
    Next cell
End If

If ActiveSheet.AutoFilterMode Then
    ActiveSheet.AutoFilterMode = False
End If
Rows("1:1").AutoFilter
Rows("1:1").AutoFilter Field:=19, Criteria1:="=" 'Columna S
Rows("1:1").AutoFilter Field:=17, Criteria1:="=LOA (without pay)", Operator:=xlOr 'Columna Q
Rows("1:1").AutoFilter Field:=13, Criteria1:="<>No", Operator:=xlAnd, Criteria2:=">5123" 'Columna M
If Application.WorksheetFunction.Subtotal(103, Range("A2:A" & lastRow)) > 0 Then
    For Each cell In Range("A2:A" & lastRow).SpecialCells(xlCellTypeVisible)
        cell.Offset(0, 18).Value = " " 'REV - LOA
    Next cell
End If
 
'---------------------M) MIN 26:48 - REVISADO OK
'FUNCIONA BIEN PORQUE DEJA EN BLANCO
'REVISAR ESTA PORQUE SI SON 5121,5122 o 5123 SE DEJAN QUIETAS EN BLANCO, PERO SI SON DE ALGUN OTRO SE PONE OK
If ActiveSheet.AutoFilterMode Then
    ActiveSheet.AutoFilterMode = False
End If
Rows("1:1").AutoFilter
Rows("1:1").AutoFilter Field:=19, Criteria1:="=" 'Columna S
Rows("1:1").AutoFilter Field:=17, Criteria1:="=LOA (without pay)", Operator:=xlOr 'Columna Q
Rows("1:1").AutoFilter Field:=16, Criteria1:="=0", Operator:=xlAnd 'Columna P
Rows("1:1").AutoFilter Field:=15, Criteria1:="=No IT14", Operator:=xlAnd 'Columna O"
Rows("1:1").AutoFilter Field:=13, Criteria1:="<5121", Operator:=xlAnd 'Columna M
If Application.WorksheetFunction.Subtotal(103, Range("A2:A" & lastRow)) > 0 Then
    For Each cell In Range("A2:A" & lastRow).SpecialCells(xlCellTypeVisible)
        cell.Offset(0, 18).Value = "OK"
    Next cell
End If


If ActiveSheet.AutoFilterMode Then
    ActiveSheet.AutoFilterMode = False
End If
Rows("1:1").AutoFilter
Rows("1:1").AutoFilter Field:=19, Criteria1:="=" 'Columna S
Rows("1:1").AutoFilter Field:=17, Criteria1:="=LOA (without pay)", Operator:=xlOr 'Columna Q
Rows("1:1").AutoFilter Field:=16, Criteria1:="=0", Operator:=xlAnd 'Columna P
Rows("1:1").AutoFilter Field:=15, Criteria1:="=No IT14", Operator:=xlAnd 'Columna O"
Rows("1:1").AutoFilter Field:=13, Criteria1:=">5123", Operator:=xlAnd 'Columna M
If Application.WorksheetFunction.Subtotal(103, Range("A2:A" & lastRow)) > 0 Then
    For Each cell In Range("A2:A" & lastRow).SpecialCells(xlCellTypeVisible)
        cell.Offset(0, 18).Value = "OK"
    Next cell
End If

'---------------------N) MIN 27:53 - REVISADO OK
If ActiveSheet.AutoFilterMode Then
    ActiveSheet.AutoFilterMode = False
End If
Rows("1:1").AutoFilter
Rows("1:1").AutoFilter Field:=19, Criteria1:="=" 'Columna S
Rows("1:1").AutoFilter Field:=12, Criteria1:="<>No", Operator:=xlAnd 'Columna L
Rows("1:1").AutoFilter Field:=13, Criteria1:="=No", Operator:=xlAnd 'Columna M
Rows("1:1").AutoFilter Field:=6, Criteria1:="<>0", Operator:=xlAnd 'Columna F
Rows("1:1").AutoFilter Field:=15, Criteria1:="=-", Operator:=xlAnd 'Columna O"
If Application.WorksheetFunction.Subtotal(103, Range("A2:A" & lastRow)) > 0 Then
    For Each cell In Range("A2:A" & lastRow).SpecialCells(xlCellTypeVisible)
        cell.Offset(0, 18).Value = "OK"
    Next cell
End If

'---------------------O) MIN 28:44 - REVISADO OK
If ActiveSheet.AutoFilterMode Then
    ActiveSheet.AutoFilterMode = False
End If
Rows("1:1").AutoFilter
Rows("1:1").AutoFilter Field:=19, Criteria1:="=" 'Columna S
Rows("1:1").AutoFilter Field:=12, Criteria1:="<>No", Operator:=xlAnd 'Columna L
Rows("1:1").AutoFilter Field:=13, Criteria1:="=No", Operator:=xlAnd 'Columna M
Rows("1:1").AutoFilter Field:=6, Criteria1:="=0,00", Operator:=xlAnd 'Columna F
Rows("1:1").AutoFilter Field:=15, Criteria1:="=No EE contribution", Operator:=xlOr, Criteria2:="=No IT14" 'Columna O"
Rows("1:1").AutoFilter Field:=16, Criteria1:="=No ER contribution", Operator:=xlAnd 'Columna P
If Application.WorksheetFunction.Subtotal(103, Range("A2:A" & lastRow)) > 0 Then
    For Each cell In Range("A2:A" & lastRow).SpecialCells(xlCellTypeVisible)
        cell.Offset(0, 18).Value = "OK"
    Next cell
End If

'---------------------P) MIN 29:30 -REVISAR PORQUE NO LO ESTA FILTRANDO las dos al mismo tiempo
If ActiveSheet.AutoFilterMode Then
    ActiveSheet.AutoFilterMode = False
End If
Rows("1:1").AutoFilter
Rows("1:1").AutoFilter Field:=19, Criteria1:="=" 'Columna S
Rows("1:1").AutoFilter Field:=4, Criteria1:="<>Short Term Disability", Operator:=xlAnd, Criteria2:="<>State STD" 'Columna D
Rows("1:1").AutoFilter Field:=13, Criteria1:="<>No", Operator:=xlAnd 'Columna M
Rows("1:1").AutoFilter Field:=12, Criteria1:="=No", Operator:=xlAnd 'Columna L
Rows("1:1").AutoFilter Field:=8, Criteria1:="<>0", Operator:=xlAnd 'Columna H"
Rows("1:1").AutoFilter Field:=15, Criteria1:="=No IT14", Operator:=xlAnd 'Columna O"
Rows("1:1").AutoFilter Field:=16, Criteria1:="=0", Operator:=xlAnd 'Columna P
If Application.WorksheetFunction.Subtotal(103, Range("A2:A" & lastRow)) > 0 Then
    For Each cell In Range("A2:A" & lastRow).SpecialCells(xlCellTypeVisible)
        cell.Offset(0, 18).Value = "OK"
    Next cell
End If

'---------------------Q) MIN 30:37 REVISAR LO DEL SHOR TERM
'REVISAR FUNCIONAMIENTO
If ActiveSheet.AutoFilterMode Then
    ActiveSheet.AutoFilterMode = False
End If
Rows("1:1").AutoFilter
Rows("1:1").AutoFilter Field:=19, Criteria1:="=" 'Columna S
Rows("1:1").AutoFilter Field:=4, Criteria1:="<>Short Term Disability", Operator:=xlAnd, Criteria2:="<>State STD" 'Columna D
Rows("1:1").AutoFilter Field:=13, Criteria1:="<>No", Operator:=xlAnd 'Columna M
Rows("1:1").AutoFilter Field:=12, Criteria1:="=No", Operator:=xlAnd 'Columna L
Rows("1:1").AutoFilter Field:=8, Criteria1:="=0,00", Operator:=xlAnd 'Columna H"
Rows("1:1").AutoFilter Field:=15, Criteria1:="=No IT14", Operator:=xlAnd 'Columna O"
Rows("1:1").AutoFilter Field:=16, Criteria1:="=No IT14", Operator:=xlAnd 'Columna P
If Application.WorksheetFunction.Subtotal(103, Range("A2:A" & lastRow)) > 0 Then
    For Each cell In Range("A2:A" & lastRow).SpecialCells(xlCellTypeVisible)
        cell.Offset(0, 18).Value = "OK"
    Next cell
End If

'---------------------R) MIN 31:05 - REVISADO OK
If ActiveSheet.AutoFilterMode Then
    ActiveSheet.AutoFilterMode = False
End If
Rows("1:1").AutoFilter
Rows("1:1").AutoFilter Field:=19, Criteria1:="=" 'Columna S
Rows("1:1").AutoFilter Field:=4, Criteria1:="<>Short Term Disability", Operator:=xlAnd, Criteria2:="<>State STD" 'Columna D
Rows("1:1").AutoFilter Field:=12, Criteria1:="<>No", Operator:=xlAnd 'Columna L
Rows("1:1").AutoFilter Field:=13, Criteria1:="<>No", Operator:=xlAnd 'Columna M
Rows("1:1").AutoFilter Field:=6, Criteria1:="<>0", Operator:=xlAnd 'Columna F
Rows("1:1").AutoFilter Field:=8, Criteria1:="<>0", Operator:=xlAnd 'Columna H
Rows("1:1").AutoFilter Field:=15, Criteria1:="=-", Operator:=xlAnd 'Columna O
Rows("1:1").AutoFilter Field:=16, Criteria1:="=0", Operator:=xlAnd 'Columna P
If Application.WorksheetFunction.Subtotal(103, Range("A2:A" & lastRow)) > 0 Then
    For Each cell In Range("A2:A" & lastRow).SpecialCells(xlCellTypeVisible)
        cell.Offset(0, 18).Value = "OK"
    Next cell
End If

'---------------------S) MIN 32:02 - REVISADO OK
If ActiveSheet.AutoFilterMode Then
    ActiveSheet.AutoFilterMode = False
End If
Rows("1:1").AutoFilter
Rows("1:1").AutoFilter Field:=19, Criteria1:="=" 'Columna S
Rows("1:1").AutoFilter Field:=4, Criteria1:="<>Short Term Disability", Operator:=xlAnd, Criteria2:="<>State STD" 'Columna D
Rows("1:1").AutoFilter Field:=12, Criteria1:="<>No", Operator:=xlAnd 'Columna L
Rows("1:1").AutoFilter Field:=13, Criteria1:="<>No", Operator:=xlAnd 'Columna M
Rows("1:1").AutoFilter Field:=6, Criteria1:="=0,00", Operator:=xlAnd 'Columna F
Rows("1:1").AutoFilter Field:=8, Criteria1:="=0,00", Operator:=xlAnd 'Columna H
Rows("1:1").AutoFilter Field:=15, Criteria1:="=No EE contribution", Operator:=xlOr, Criteria2:="=No IT14" 'Columna O"
Rows("1:1").AutoFilter Field:=16, Criteria1:="=No IT14", Operator:=xlAnd 'Columna P
If Application.WorksheetFunction.Subtotal(103, Range("A2:A" & lastRow)) > 0 Then
    For Each cell In Range("A2:A" & lastRow).SpecialCells(xlCellTypeVisible)
        cell.Offset(0, 18).Value = "OK"
    Next cell
End If

''''''''''''''''''''''''''''''''''' VERIFICAR DE ACA PARA ABAJO

'---------------------T) MIN 33:10 - REVISADO OK
If ActiveSheet.AutoFilterMode Then
    ActiveSheet.AutoFilterMode = False
End If
Rows("1:1").AutoFilter
Rows("1:1").AutoFilter Field:=19, Criteria1:="=" 'Columna S
Rows("1:1").AutoFilter Field:=4, Criteria1:="<>Short Term Disability", Operator:=xlAnd, Criteria2:="<>State STD" 'Columna D
Rows("1:1").AutoFilter Field:=12, Criteria1:="<>No", Operator:=xlAnd 'Columna L
Rows("1:1").AutoFilter Field:=13, Criteria1:="<>No", Operator:=xlAnd 'Columna M
Rows("1:1").AutoFilter Field:=6, Criteria1:="=0,00", Operator:=xlAnd 'Columna F
Rows("1:1").AutoFilter Field:=8, Criteria1:="<>0", Operator:=xlAnd 'Columna H
Rows("1:1").AutoFilter Field:=15, Criteria1:="=No IT14", Operator:=xlAnd 'Columna O"
Rows("1:1").AutoFilter Field:=16, Criteria1:="=0", Operator:=xlAnd 'Columna P
If Application.WorksheetFunction.Subtotal(103, Range("A2:A" & lastRow)) > 0 Then
    For Each cell In Range("A2:A" & lastRow).SpecialCells(xlCellTypeVisible)
        cell.Offset(0, 18).Value = "OK"
    Next cell
End If

'---------------------U) MIN 33:53
'REVISAR ESTE CUANDO ES EL FILTRO CONTRARIO AL ANTERIOR, REVISARLO CON LESLIE
If ActiveSheet.AutoFilterMode Then
    ActiveSheet.AutoFilterMode = False
End If
Rows("1:1").AutoFilter
Rows("1:1").AutoFilter Field:=19, Criteria1:="=" 'Columna S
Rows("1:1").AutoFilter Field:=4, Criteria1:="<>Short Term Disability", Operator:=xlAnd, Criteria2:="<>State STD" 'Columna D
Rows("1:1").AutoFilter Field:=12, Criteria1:="<>No", Operator:=xlAnd 'Columna L
Rows("1:1").AutoFilter Field:=13, Criteria1:="<>No", Operator:=xlAnd 'Columna M
Rows("1:1").AutoFilter Field:=8, Criteria1:="=0,00", Operator:=xlAnd 'Columna H
Rows("1:1").AutoFilter Field:=6, Criteria1:="<>0", Operator:=xlAnd 'Columna F
Rows("1:1").AutoFilter Field:=15, Criteria1:="=-", Operator:=xlAnd 'Columna O"
Rows("1:1").AutoFilter Field:=16, Criteria1:="=No IT14", Operator:=xlAnd 'Columna P
If Application.WorksheetFunction.Subtotal(103, Range("A2:A" & lastRow)) > 0 Then
    For Each cell In Range("A2:A" & lastRow).SpecialCells(xlCellTypeVisible)
        cell.Offset(0, 18).Value = "OK"
    Next cell
End If

'---------------------V) MIN 34:13 - REVISAR LA PARTE DE COMO PONER EK OK DADO QUE CAMBIA PERO EL FILTRO ESTA BUENO
'PREGUNTARLE A LESLIE SI LOS QUE FUMAN O NO
'REVISAR ESTE PORQUE A UNOS YA LES PUSO OK ENTONCES NO LOS CUENTA
'EMPLEADOS QUE FUMAN, ESTE SE FILTRA POR LOS QUE TENGAN EL MISMO VALOR EN LA N Y EN LA O, LOS QUE SEAN LO MISMO SE LES PONE OK
If ActiveSheet.AutoFilterMode Then
    ActiveSheet.AutoFilterMode = False
End If
Rows("1:1").AutoFilter
Rows("1:1").AutoFilter Field:=19, Criteria1:="=" 'Columna S
Rows("1:1").AutoFilter Field:=12, Criteria1:="=2511", Operator:=xlOr, Criteria2:="=2995" 'Columna L
If Application.WorksheetFunction.Subtotal(103, Range("A2:A" & lastRow)) > 0 Then
    For Each cell In Range("A2:A" & lastRow).SpecialCells(xlCellTypeVisible)
        cell.Offset(0, 18).Value = "=IF(RC[-5]=RC[-4], ""OK"", "" "")" ' TABACO - REV LATER
    Next cell
End If
If ActiveSheet.AutoFilterMode Then
    ActiveSheet.AutoFilterMode = False
End If
Range("S:S").Copy
Range("S:S").PasteSpecial Paste:=xlPasteValues

'---------------------W) MIN 36:15 - REVISADO Y OK, PERO FALTA CREAR TODA LA PARTE QUE SE HACE EN OTRO ARCHIVO QUE SE LLAMA "STD"
'SOLO SE LES PONE REV A UNOS QUE NO SALEN CUANDO SE HAGA EL OTRO DOCUMENTO Y EL RESTO SE MARCAN COMO OK, LOS QUE NO TENGNA NINGUNA DIFERENCIA SE MARCAN CON OK

If ActiveSheet.AutoFilterMode Then
    ActiveSheet.AutoFilterMode = False
End If
lastRow = Cells(Rows.Count, "A").End(xlUp).row
Rows("1:1").AutoFilter
Rows("1:1").AutoFilter Field:=19, Criteria1:="=" 'Columna S
Rows("1:1").AutoFilter Field:=4, Criteria1:="=Short term disability", Operator:=xlOr, Criteria2:="=State STD" 'Columna D

    'Pasa informacion al nuevo documento de STD
    Set wbSTD = Workbooks.Open(Ruta_Audi & "\" & fecha_hoy & " STD" & ".xlsx")
    wbPlantilla_BSWIFT.Activate
    Sheets("Bswift-SAP").Activate
    Range("A1:S" & lastRow).SpecialCells(xlCellTypeVisible).Copy
    wbSTD.Sheets("Original").Activate
    Sheets("Original").Range("A1").PasteSpecial Paste:=xlPasteAll
    Columns("A:S").AutoFit
    Application.CutCopyMode = False
    wbSTD.Save
    wbSTD.Close
  
wbPlantilla_BSWIFT.Activate
Sheets("Bswift-SAP").Activate

'PENDIENTE DE VERIFICAR QUIEN SI OK Y QUIEN REV
'If Application.WorksheetFunction.Subtotal(103, Range("A2:A" & lastRow)) > 0 Then
'    For Each cell In Range("A2:A" & lastRow).SpecialCells(xlCellTypeVisible)
'        cell.Offset(0, 18).Value = "STD - REV LATER"
'    Next cell
'End If



'De aca en a delante se hacen los filtros para revisar lo malo

'---------------------X) MIN 39:00 - REVISADO OK
'ESTA SE PEGA EN OTRA HOJA DE EMPLOY MISSING SETUP
If ActiveSheet.AutoFilterMode Then
    ActiveSheet.AutoFilterMode = False
End If
Rows("1:1").AutoFilter
Rows("1:1").AutoFilter Field:=19, Criteria1:="=" 'Columna S
Rows("1:1").AutoFilter Field:=6, Criteria1:="<>0", Operator:=xlAnd 'Columna F
Rows("1:1").AutoFilter Field:=15, Criteria1:="=No IT14", Operator:=xlAnd 'Columna O"
Rows("1:1").AutoFilter Field:=12, Criteria1:="<>No", Operator:=xlAnd 'Columna L

    'Crea hoja de EE missing set up
    Range("A1:S" & lastRow).SpecialCells(xlCellTypeVisible).Copy
    wbFiltrado.Sheets("EE missing set up").Activate
    Sheets("EE missing set up").Range("A1").PasteSpecial Paste:=xlPasteAll
    Columns("A:S").AutoFit
    Application.CutCopyMode = False
    wbPlantilla_BSWIFT.Activate
    Sheets("Bswift-SAP").Activate

'If Application.WorksheetFunction.Subtotal(103, Range("A2:A" & lastRow)) > 0 Then
'    For Each cell In Range("A2:A" & lastRow).SpecialCells(xlCellTypeVisible)
'        cell.Offset(0, 18).Value = "EE MISSING SET UP"
'    Next cell
'End If

'---------------------Z) MIN 40:12 - REVISADO OK
'ESTA SE PEGA EN OTRA HOJA DE EMPLOY DISCREPANCIES
If ActiveSheet.AutoFilterMode Then
    ActiveSheet.AutoFilterMode = False
End If
Rows("1:1").AutoFilter
Rows("1:1").AutoFilter Field:=19, Criteria1:="=" 'Columna S
Rows("1:1").AutoFilter Field:=6, Criteria1:="<>0", Operator:=xlAnd 'Columna F
Rows("1:1").AutoFilter Field:=15, Criteria1:="<>No IT14", Operator:=xlAnd 'Columna O"

    'Crea hoja de EE Discrepancies
    Range("A1:S" & lastRow).SpecialCells(xlCellTypeVisible).Copy
    wbFiltrado.Sheets("EE Discrepancies").Activate
    Sheets("EE Discrepancies").Range("A1").PasteSpecial Paste:=xlPasteAll
    Columns("A:S").AutoFit
    Application.CutCopyMode = False
    wbPlantilla_BSWIFT.Activate
    Sheets("Bswift-SAP").Activate
    
'If Application.WorksheetFunction.Subtotal(103, Range("A2:A" & lastRow)) > 0 Then
'    For Each cell In Range("A2:A" & lastRow).SpecialCells(xlCellTypeVisible)
'        cell.Offset(0, 18).Value = "EE DISCREPANCIES"
'    Next cell
'End If


'---------------------AA) MIN 40:59 - REVISADO OK
'ESTA SE PEGA EN OTRA HOJA DE EMPLOYER MISSING SETUP
If ActiveSheet.AutoFilterMode Then
    ActiveSheet.AutoFilterMode = False
End If
Rows("1:1").AutoFilter
Rows("1:1").AutoFilter Field:=19, Criteria1:="=" 'Columna S
Rows("1:1").AutoFilter Field:=8, Criteria1:="<>0", Operator:=xlAnd 'Columna H
Rows("1:1").AutoFilter Field:=16, Criteria1:="=No IT14", Operator:=xlAnd 'Columna P

    'Crea hoja de ER Missing set up
    Range("A1:S" & lastRow).SpecialCells(xlCellTypeVisible).Copy
    wbFiltrado.Sheets("ER missing set up").Activate
    Sheets("ER missing set up").Range("A1").PasteSpecial Paste:=xlPasteAll
    Columns("A:S").AutoFit
    Application.CutCopyMode = False
    wbPlantilla_BSWIFT.Activate
    Sheets("Bswift-SAP").Activate

'If Application.WorksheetFunction.Subtotal(103, Range("A2:A" & lastRow)) > 0 Then
'    For Each cell In Range("A2:A" & lastRow).SpecialCells(xlCellTypeVisible)
'        cell.Offset(0, 18).Value = "ER MISSING SET UP"
'    Next cell
'End If

'---------------------AB) MIN 42:07 - REVISADO OK
'ESTA SE PEGA EN OTRA HOJA DE EMPLOYER DISCREPANCIES
If ActiveSheet.AutoFilterMode Then
    ActiveSheet.AutoFilterMode = False
End If
Rows("1:1").AutoFilter
Rows("1:1").AutoFilter Field:=19, Criteria1:="=" 'Columna S
Rows("1:1").AutoFilter Field:=13, Criteria1:="<>No", Operator:=xlAnd 'Columna M
Rows("1:1").AutoFilter Field:=16, Criteria1:="<>0", Operator:=xlOr, Criteria2:="<>No IT14" 'Columna P
Rows("1:1").AutoFilter Field:=4, Criteria1:="<>Short term disability", Operator:=xlOr, Criteria2:="<>State STD" 'Columna D

    'Crea hoja de ER discrepancies
    Range("A1:S" & lastRow).SpecialCells(xlCellTypeVisible).Copy
    wbFiltrado.Sheets("ER discrepancies").Activate
    Sheets("ER discrepancies").Range("A1").PasteSpecial Paste:=xlPasteAll
    Columns("A:S").AutoFit
    Application.CutCopyMode = False
    wbPlantilla_BSWIFT.Activate
    Sheets("Bswift-SAP").Activate

'If Application.WorksheetFunction.Subtotal(103, Range("A2:A" & lastRow)) > 0 Then
'    For Each cell In Range("A2:A" & lastRow).SpecialCells(xlCellTypeVisible)
'        cell.Offset(0, 18).Value = "ER DISCREPANCIES"
'    Next cell
'End If

'---------------------AC) MIN 43:09 - REVISADO OK
'REVISAR ESTO PORQUE NO ESTA SALIENDO NADA
'ESTA SE PEGA EN OTRA HOJA TABACO
If ActiveSheet.AutoFilterMode Then
    ActiveSheet.AutoFilterMode = False
End If
Rows("1:1").AutoFilter
Rows("1:1").AutoFilter Field:=19, Criteria1:="=" 'Columna S
Rows("1:1").AutoFilter Field:=14, Criteria1:="<>" 'Columna N

    'Crea hoja de ER Tabaco
    Range("A1:S" & lastRow).SpecialCells(xlCellTypeVisible).Copy
    wbFiltrado.Sheets("Tabaco").Activate
    Sheets("Tabaco").Range("A1").PasteSpecial Paste:=xlPasteAll
    Columns("A:S").AutoFit
    Application.CutCopyMode = False
    wbFiltrado.Save
    wbPlantilla_BSWIFT.Activate
    Sheets("Bswift-SAP").Activate
    wbFiltrado.Close
    
'If Application.WorksheetFunction.Subtotal(103, Range("A2:A" & lastRow)) > 0 Then
'    For Each cell In Range("A2:A" & lastRow).SpecialCells(xlCellTypeVisible)
'        cell.Offset(0, 18).Value = "RECARGO POR TABACO"
'    Next cell
'End If

'Quita todos los filtros y guarda
If ActiveSheet.AutoFilterMode Then
    ActiveSheet.AutoFilterMode = False
End If
Rows("1:1").AutoFilter
wbPlantilla_BSWIFT.Save

End Sub

Sub STD_Bswift()
'MIN 36:15

'Abre documento de STD y trae la base de SAP
Set wbSTD = Workbooks.Open(Ruta_Audi & "\" & fecha_hoy & " STD" & ".xlsx")
wbPlantilla_BSWIFT.Activate
Sheets.Add(After:=Sheets(Sheets.Count)).Name = "EE From STD"
Sheets("SAP-Bswift").Activate
lastRow = Cells(Rows.Count, "A").End(xlUp).row
Range("A1:H" & lastRow).Select
Selection.Copy
wbSTD.Activate
Sheets("SAP-Bswift").Activate
Sheets("SAP-Bswift").Range("A1").PasteSpecial Paste:=xlPasteAll
Columns("A:H").AutoFit

'Tabla dinamica
Sheets("Original").Activate

    Dim ult_Tabla As Long
    ult_Tabla = Cells(Rows.Count, "A").End(xlUp).row
    Dim rangoTabla1 As Range
    Set rangoTabla1 = Sheets("Original").Range("A1:S" & ult_Tabla)
    ActiveSheet.ListObjects.Add(xlSrcRange, rangoTabla1, , xlYes).Name = "Tabla1"
    
    'Crear tabla dinamica
    Dim celdaTablaDinamica1 As Range
    Set celdaTablaDinamica1 = Sheets("Tabla dinamica").Range("A1")
    Dim tablaDinamica1 As PivotTable

    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        rangoTabla1, Version:=6).CreatePivotTable TableDestination:= _
        celdaTablaDinamica1, TableName:="tablaDinamica1", DefaultVersion:=6
    Sheets("Tabla dinamica").Select
    
    With ActiveSheet.PivotTables("tablaDinamica1").PivotFields("EE ID")
        .Orientation = xlRowField
        .Position = 1
    End With
    ActiveSheet.PivotTables("tablaDinamica1").AddDataField ActiveSheet.PivotTables( _
        "tablaDinamica1").PivotFields("ER Cost Per PP"), "Suma de ER Cost Per PP", _
        xlSum
               
'Realiza formulas
Sheets("Tabla dinamica").Activate
Range("D1").Value = "SAP ID"
Range("E1").Value = "Concatenate"
Range("F1").Value = "Vlookup from sap"
Range("G1").Value = "Diff"
Range("H1").Value = "Comment"
lastRow = Cells(Rows.Count, "A").End(xlUp).row
Sheets("Tabla dinamica").Range("D2:D" & lastRow) = "=RC[-3]"
Sheets("Tabla dinamica").Range("E2:E" & lastRow) = "=RC[-4]&5122"
Sheets("Tabla dinamica").Range("F2:F" & lastRow) = "=VLOOKUP(RC[-1],'SAP-Bswift'!C[-5]:C[2],8)"
Sheets("Tabla dinamica").Range("G2:G" & lastRow) = "=RC[-1]-RC[-5]"
Sheets("Tabla dinamica").Range("H2:H" & lastRow) = "=IFERROR(IF(RC[-1]=0,""OK"",""REV STD""),""REV STD"")"
'Range("D:H").Copy
'Range("D:H").PasteSpecial Paste:=xlPasteValues
'Application.CutCopyMode = False
       
'Copia datos y los pega en la plantilla
lastRow = Cells(Rows.Count, "E").End(xlUp).row
Range("D1:H" & lastRow).Select
Selection.Copy
wbPlantilla_BSWIFT.Activate
Sheets("EE From STD").Activate
Sheets("EE From STD").Range("A1").PasteSpecial Paste:=xlPasteValues
Columns("A:E").AutoFit
Application.CutCopyMode = False
wbPlantilla_BSWIFT.Save

'Guarda y cierra STD
wbSTD.Save
wbSTD.Close

'Hace el buscar para los STD para poner los comentarios
wbPlantilla_BSWIFT.Activate
Sheets("Bswift-SAP").Activate
If ActiveSheet.AutoFilterMode Then
    ActiveSheet.AutoFilterMode = False
End If
lastRow = Cells(Rows.Count, "A").End(xlUp).row
Rows("1:1").AutoFilter
Rows("1:1").AutoFilter Field:=4, Criteria1:="=Short term disability", Operator:=xlOr, Criteria2:="=State STD" 'Columna D
Rows("1:1").AutoFilter Field:=19, Criteria1:="=" 'Columna S


If Application.WorksheetFunction.Subtotal(103, Range("A2:A" & lastRow)) > 0 Then
    For Each cell In Range("A2:A" & lastRow).SpecialCells(xlCellTypeVisible)
        cell.Offset(0, 18).Value = "=VLOOKUP(RC[-18],'EE From STD'!C[-18]:C[-14],5,0)"
    Next cell
End If
If ActiveSheet.AutoFilterMode Then
    ActiveSheet.AutoFilterMode = False
End If
Range("S:S").Copy
Range("S:S").PasteSpecial Paste:=xlPasteValues
Sheets("EE From STD").Visible = False
wbPlantilla_BSWIFT.Save
       
End Sub
