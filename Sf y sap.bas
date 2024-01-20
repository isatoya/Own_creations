Attribute VB_Name = "Desarrollo"
Option Explicit
'variables para todo el proyecto
    Dim pais As String
    Dim ReporteSF As String
    Dim ruta As String
    Dim ruta2 As String
    Dim fecha As String
        

Sub InicializarVariables()
'Definicion de las variables
    
    'Pais seleccionado
    pais = Userform1.ComboBox1.Value
    
    'Ruta de la macro
    ruta = ThisWorkbook.Path
    
    'Ruta de la carpeta del pais
    ruta2 = ruta & "\" & pais
    
    'Fecha del dia que se ejecuta la macro
    fecha = Format(Date, "mm-dd-yyyy")
    
End Sub

Sub CrearCarpetas()
    
    InicializarVariables
    
    ' Verificar si la carpeta del paํs existe, si no, la crea
    If Dir(ruta2, vbDirectory) = "" Then
        MkDir ruta2
    End If
        
End Sub

Sub Abrir_activar()

    Dim ruta_ReporteSF As String
    Dim ReporteSF As Workbook
    
    'Llama las variables
        InicializarVariables
    
    ' Solicitar al usuario que abra un archivo
    ruta_ReporteSF = Application.GetOpenFilename("Archivos Excel (*.xls; *.xlsx), *.xls; *.xlsx")

    ' Abre y activa reporte que seleccie el usuario de SF
        If ruta_ReporteSF <> "Falso" Then
        Set ReporteSF = Workbooks.Open(ruta_ReporteSF)
        ReporteSF.Activate
        
        'Crea el archivo nuevo con formato XLSX
        ActiveSheet.Cells.Select
        Selection.Copy
        Workbooks.Add
        ActiveSheet.Range("A1").PasteSpecial Paste:=xlPasteValues
        Application.CutCopyMode = False ' Desactiva el modo de copia
        
        ActiveWorkbook.SaveAs ruta2 & "\" & fecha & " SF vs SAP Address audit  " & pais & ".xlsx"
        ActiveWorkbook.Close SaveChanges:=True
        ReporteSF.Close
        
    Else
        MsgBox "Operaci๓n cancelada por el usuario.", vbInformation
    End If
End Sub

Sub CambiosFormato()


    'Llama las variables
        InicializarVariables
        Dim wb As Workbook
        
        If Dir(ruta2 & "\" & fecha & " SF vs SAP Address audit  " & pais & ".xlsx") <> "" Then ' Verificar si el archivo existe
        
        ' Abrir el archivo
        Set wb = Workbooks.Open(ruta2 & "\" & fecha & " SF vs SAP Address audit  " & pais & ".xlsx")
        wb.Activate
        
        'Elimina filas y columnas
        Rows("1:2").Delete
        Columns("F:G").Delete
        Columns("G").Delete
        Columns("R:Z").Delete
        
        'Cambios de formato
        Columns("A:A").NumberFormat = "General"
        Columns("D:D").NumberFormat = "General"
        Columns("G:G").NumberFormat = "General"
        Columns("A:A").Value = Columns("A:A").Value
        Columns("D:D").Value = Columns("D:D").Value
        Columns("G:G").Value = Columns("G:G").Value
        Columns("H:H").NumberFormat = "mm/dd/yyyy"
        Columns("I:I").Value = Columns("I:I").Value
        Columns("I:I").NumberFormat = "mm/dd/yyyy"
        Columns("N:N").NumberFormat = "General"
        Columns("N:N").Value = Columns("N:N").Value
        
        'Agregar Nuevos tutulos
        
        Range("D1").Interior.Color = RGB(255, 255, 0)
        
        With Range("S1")
            .Value = "Change through ESS?"
            .Interior.Color = RGB(0, 32, 96)
            .Font.Bold = True
            .Font.Color = RGB(255, 255, 255)
        End With
        
        With Range("T1")
            .Value = "SAP Start date"
            .Interior.Color = RGB(0, 32, 96)
            .Font.Bold = True
            .Font.Color = RGB(255, 255, 255)
        End With

        With Range("U1")
            .Value = "SAPAddress line 1"
            .Interior.Color = RGB(0, 32, 96)
            .Font.Bold = True
            .Font.Color = RGB(255, 255, 255)
        End With

        With Range("V1")
            .Value = "SAPAddress line 2"
            .Interior.Color = RGB(0, 32, 96)
            .Font.Bold = True
            .Font.Color = RGB(255, 255, 255)
        End With

        With Range("W1")
            .Value = "SAP City / County"
            .Interior.Color = RGB(0, 32, 96)
            .Font.Bold = True
            .Font.Color = RGB(255, 255, 255)
        End With

        With Range("X1")
            .Value = "SAP State"
            .Interior.Color = RGB(0, 32, 96)
            .Font.Bold = True
            .Font.Color = RGB(255, 255, 255)
        End With

        With Range("Y1")
            .Value = "SAP Zip Code"
            .Interior.Color = RGB(0, 32, 96)
            .Font.Bold = True
            .Font.Color = RGB(255, 255, 255)
        End With

        With Range("Z1")
            .Value = "SAP Country"
            .Interior.Color = RGB(0, 32, 96)
            .Font.Bold = True
            .Font.Color = RGB(255, 255, 255)
        End With
        
        With Range("AA1")
            .Value = "Start date check"
            .Interior.Color = RGB(0, 176, 80)
            .Font.Bold = True
            .Font.Color = RGB(255, 255, 255)
        End With
        
        With Range("AB1")
            .Value = "Address line 1 check"
            .Interior.Color = RGB(0, 176, 80)
            .Font.Bold = True
            .Font.Color = RGB(255, 255, 255)
        End With

        With Range("AC1")
            .Value = "Address line 2 check"
            .Interior.Color = RGB(0, 176, 80)
            .Font.Bold = True
            .Font.Color = RGB(255, 255, 255)
        End With

        With Range("AD1")
            .Value = "City / County check"
            .Interior.Color = RGB(0, 176, 80)
            .Font.Bold = True
            .Font.Color = RGB(255, 255, 255)
        End With

        With Range("AE1")
            .Value = "State check"
            .Interior.Color = RGB(0, 176, 80)
            .Font.Bold = True
            .Font.Color = RGB(255, 255, 255)
        End With

        With Range("AF1")
            .Value = "Zip Code check"
            .Interior.Color = RGB(0, 176, 80)
            .Font.Bold = True
            .Font.Color = RGB(255, 255, 255)
        End With

        With Range("AG1")
            .Value = "Country Check"
            .Interior.Color = RGB(0, 176, 80)
            .Font.Bold = True
            .Font.Color = RGB(255, 255, 255)
        End With
        
        With Range("AH1")
            .Value = "Any check failed?"
            .Interior.Color = RGB(0, 176, 80)
            .Font.Bold = True
            .Font.Color = RGB(255, 255, 255)
        End With
        
         wb.Save
        
        'Inmovilizar primera fila
        Rows("1:1").Select
        With ActiveWindow
            .SplitColumn = 0
            .SplitRow = 1
        End With
        ActiveWindow.FreezePanes = True
        
        'Evalua los datos de la columna D y se๑ala los que no esten en formato numero
        
            Dim lastRow As Long
            Dim rng As Range
            Dim cell As Range
            
            lastRow = ActiveSheet.Cells(ActiveSheet.Rows.Count, "A").End(xlUp).Row
            
            
            If lastRow >= 2 Then
                Set rng = ActiveSheet.Range("D2:D" & lastRow)
                
                For Each cell In rng
                    If Not IsNumeric(cell.Value) Or cell.Value = "" Then
                        cell.Interior.Color = RGB(255, 255, 0)
                    End If
                Next cell
            End If
        
        wb.Save
        
        'Cambios de formato
        Columns("A:AH").AutoFit
        
        With Range("J1:Q1").Interior
            .Color = RGB(173, 216, 230)
        End With
        
        With Range("H:H").Font
            .Color = RGB(0, 0, 128)
            .Bold = True
        End With
        
        'Organizar por la columna I por fechas
        ActiveSheet.Sort.SortFields.Clear
        ActiveSheet.Sort.SortFields.Add Key:=Range("I:I"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        With ActiveSheet.Sort
            .SetRange Range("A:AH")
            .Header = xlYes
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With
        
        wb.Save
        
        'Formula para la columa S (A=G)
        If lastRow >= 2 Then
            Range("S2:S" & lastRow).Formula = "=A2=G2"
        End If
        
        'Verificacion celdas en verde
        If lastRow >= 2 Then
            Range("AA2:AA" & lastRow).Formula = "=T2=I2"
        End If
        
        If lastRow >= 2 Then
            Range("AB2:AB" & lastRow).Formula = "=U2=K2"
        End If
        
        If lastRow >= 2 Then
            Range("AC2:AC" & lastRow).Formula = "=V2=L2"
        End If
        
        If lastRow >= 2 Then
            Range("AD2:AD" & lastRow).Formula = "=W2=M2"
        End If
        
        If lastRow >= 2 Then
            Range("AE2:AE" & lastRow).Formula = "=X2=Q2"
        End If
        
        If lastRow >= 2 Then
            Range("AF2:AF" & lastRow).Formula = "=Y2=N2"
        End If
        
        If lastRow >= 2 Then
            Range("AG2:AG" & lastRow).Formula = "=Z2=O2"
        End If
        
        If lastRow >= 2 Then
            Range("AH2:AH" & lastRow).Formula = "=IF(COUNTIF(AA2:AG2,FALSE)=0,""All ok"",""Review"")"
        End If
        
        'Mostrar mensaje al usuario
        MsgBox "Please enter to modify and fill data in column D", vbInformation
        
    Else
        MsgBox "El archivo no existe en la ruta especificada.", vbExclamation
    End If
 

End Sub

Sub sap()

    InicializarVariables
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

    'Entra a la transacion SQ01
    session.findById("wnd[0]/tbar[0]/okcd").Text = "/nSQ01"
    session.findById("wnd[0]").sendVKey 0
    
    'Confirmamos el Environment
    session.findById("wnd[0]/mbar/menu[5]/menu[0]").Select
    session.findById("wnd[1]/usr/radRAD1").Select
    session.findById("wnd[1]/tbar[0]/btn[2]").press
    
    'Entra al HR_ALL_SITE
    session.findById("wnd[0]/tbar[1]/btn[19]").press
    session.findById("wnd[1]/usr/cntlGRID1/shellcont/shell").selectColumn "DBGBNUM"
    session.findById("wnd[1]/tbar[0]/btn[29]").press
    session.findById("wnd[2]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN001-LOW").Text = "HR_ALL_SITE"
    session.findById("wnd[2]/tbar[0]/btn[0]").press
    session.findById("wnd[1]/usr/cntlGRID1/shellcont/shell").selectedRows = "0"
    session.findById("wnd[1]/tbar[0]/btn[0]").press
    
    'Entra el nombre del query
    session.findById("wnd[0]/usr/ctxtRS38R-QNUM").Text = "SF_IT006_AUDIT"
    session.findById("wnd[0]/usr/ctxtRS38R-QNUM").SetFocus
    session.findById("wnd[0]/usr/ctxtRS38R-QNUM").caretPosition = 14
    session.findById("wnd[0]").sendVKey 0
    
    'PONE LA VARIANTE
    session.findById("wnd[0]/tbar[1]/btn[8]").press
    
    session.findById("wnd[0]/tbar[1]/btn[17]").press
    session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").currentCellRow = -1
    session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").selectColumn "VARIANT"
    session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").pressToolbarButton "&MB_FILTER"
    session.findById("wnd[2]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN001-LOW").Text = "/" & pais & "*"
    session.findById("wnd[2]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN001-LOW").caretPosition = 5
    session.findById("wnd[2]").sendVKey 0
    session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").selectedRows = "0"
    session.findById("wnd[1]/tbar[0]/btn[2]").press
    
    'Limpia si hay datos y pega
    session.findById("wnd[0]/usr/btn%_PNPPERNR_%_APP_%-VALU_PUSH").press
    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE").Columns.elementAt(1).Selected = True
    session.findById("wnd[1]/tbar[0]/btn[16]").press
    
    'VA AL OTRO ARCHIVO PARA COPIAR
    Dim ult_fila As Long
    Workbooks(fecha & " SF vs SAP Address audit  " & pais & ".xlsx").Activate
    Range("D1").Select
    ult_fila = Cells(Rows.Count, "C").End(xlUp).Row

    If ult_fila >= 2 Then
            Columns("D").Resize(ult_fila - 1).Offset(1, 0).Copy
        Else
            MsgBox "No hay datos en la columna D.", vbExclamation
    End If
    
    session.findById("wnd[1]/tbar[0]/btn[24]").press
    session.findById("wnd[1]/tbar[0]/btn[8]").press
    session.findById("wnd[0]/tbar[1]/btn[8]").press

    'Exporta fichero
    session.findById("wnd[0]/mbar/menu[0]/menu[3]/menu[2]").Select
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").Select
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").SetFocus
    session.findById("wnd[1]/tbar[0]/btn[0]").press
    session.findById("wnd[1]/usr/ctxtDY_PATH").Text = ruta2
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = fecha & " - SAP Address audit  " & pais & ".XLS"
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 10
    session.findById("wnd[1]/tbar[0]/btn[11]").press

    'Organizar formato
    Application.CutCopyMode = False
    Workbooks.Open ruta2 & "\" & fecha & " - SAP Address audit  " & pais & ".XLS"
    ActiveSheet.Cells.Select
    Selection.Copy
    Workbooks.Add
    ActiveSheet.Paste
    ActiveWorkbook.SaveAs ruta2 & "\" & fecha & " - SAP Address audit  " & pais & ".XLSX"
    ActiveWorkbook.Close SaveChanges:=True
    Workbooks(fecha & " - SAP Address audit  " & pais & ".XLS").Close
    Kill ruta2 & "\" & fecha & " - SAP Address audit  " & pais & ".XLS"
    
    'Cambios de formato
    Workbooks.Open ruta2 & "\" & fecha & " - SAP Address audit  " & pais & ".XLSX"
    organiza_doc_sap

End Sub

Sub organiza_doc_sap()

    InicializarVariables
    
    Workbooks(fecha & " - SAP Address audit  " & pais & ".XLSX").Activate
    Rows("1:4").Delete
    Rows("2").Delete
    Columns("A").Delete
    Columns("A:J").AutoFit
    ActiveWorkbook.Save

End Sub

Sub buscarv_sap()

    InicializarVariables
    
    'Crea una hoja nueva
    Workbooks(fecha & " SF vs SAP Address audit  " & pais & ".xlsx").Activate
    Sheets.Add(After:=Sheets(Sheets.Count)).Name = "SAP"
    Sheets.Add(After:=Sheets(Sheets.Count)).Name = "REVIEW"
    
    'Trae los datos descargados
    Workbooks(fecha & " - SAP Address audit  " & pais & ".XLSX").Activate
    Sheets(1).Cells.Copy
    Workbooks(fecha & " SF vs SAP Address audit  " & pais & ".xlsx").Activate
    Sheets("SAP").Range("A1").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    
    'Cierra el reporte que se descargo de sap
    Workbooks(fecha & " - SAP Address audit  " & pais & ".XLSX").Close
    
    'Activa el reporte original
    Workbooks(fecha & " SF vs SAP Address audit  " & pais & ".xlsx").Activate
    Sheets(1).Activate
    
    Dim ultima As Long
    ultima = ActiveSheet.Cells(ActiveSheet.Rows.Count, "A").End(xlUp).Row
    
    'Realiza los buscarv entre el reporte de sap y el de sf
     
    With ActiveSheet.Range("T2:T" & ultima)
        .Formula = "=VLOOKUP(RC[-16],SAP!C[-19]:C[-17],3,0)"
    End With
    
    With ActiveSheet.Range("U2:U" & ultima)
        .Formula = "=VLOOKUP(RC[-17],SAP!C[-20]:C[-17],4,0)"
    End With
    
    With ActiveSheet.Range("V2:V" & ultima)
        .Formula = "=VLOOKUP(RC[-18],SAP!C[-21]:C[-17],5,0)"
    End With
    
    With ActiveSheet.Range("W2:W" & ultima)
        .Formula = "=VLOOKUP(RC[-19],SAP!C[-22]:C[-17],6,0)"
    End With
    
    With ActiveSheet.Range("X2:X" & ultima)
        .Formula = "=VLOOKUP(RC[-20],SAP!C[-23]:C[-17],7,0)"
    End With
    
    With ActiveSheet.Range("Y2:Y" & ultima)
        .Formula = "=VLOOKUP(RC[-21],SAP!C[-24]:C[-17],8,0)"
    End With
    
    With ActiveSheet.Range("Z2:Z" & ultima)
        .Formula = "=VLOOKUP(RC[-22],SAP!C[-25]:C[-17],9,0)"
    End With
    
    'Copia y pega como valores
    Columns("T:Z").Copy
    Columns("T:Z").PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    Columns("T:T").NumberFormat = "mm/dd/yyyy"
    
    'Hace un remplazo de 0 por vacios
    Columns("T:Z").Select
    Selection.Replace What:="0", Replacement:="", LookAt:=xlWhole, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False
        
    'Hace un reemplazo de US por Unidated States
    Columns("Z:Z").Select
    Selection.Replace What:="US", Replacement:="United States", LookAt:=xlWhole, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False
    
    'Hace un reemplazo de CA por Canada
    Columns("Z:Z").Select
    Selection.Replace What:="CA", Replacement:="Canada", LookAt:=xlWhole, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False
    
    'Guarda
    Columns("A:AH").AutoFit
    Sheets("SAP").Visible = xlSheetHidden
    ActiveWorkbook.Save
    
End Sub

Sub Hoja_revisar()

    InicializarVariables
    
    'Aplicar autofiltro en la fila 1 y copia las celdas visibles
    Rows("1:1").AutoFilter
    ActiveSheet.Range("AH1").AutoFilter Field:=34, Criteria1:="Review"
    Cells.Select
    Selection.Copy
    
    'Pega en la otra hoja
    Sheets("REVIEW").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Columns("A:AH").AutoFit
    
    'Desactivar el autofiltro
    Sheets(1).Activate
    ActiveSheet.AutoFilterMode = False
    
    'Guarda
    ActiveWorkbook.Save
    ActiveWorkbook.Close
    
    'Elimina el archivo de sap
    Kill ruta2 & "\" & fecha & " - SAP Address audit  " & pais & ".XLSX"
    
End Sub

Sub AbrirUserForm()
    'Abre el formulario
    Userform1.Show
End Sub
