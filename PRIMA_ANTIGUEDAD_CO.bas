Attribute VB_Name = "PRIMA_ANTIGUEDAD_CO"
Option Explicit
'variables para todo el proyecto

    Dim Fecha1, Fecha2, Fecha3, mes, mes2, Mes_texto, A�o, mes_2, Mes_siguiente, nombreArchivo As String
    Dim ruta, Ruta_A�o, Ruta_Mes As String
    Dim LastRow, i As Long
    Dim CelsAG As Object
    
Sub Ejecutar_ANTIGUEDAD_CO()

    ' Verificar si hay datos en las celdas I8 y M8
    If ThisWorkbook.Sheets("Reportes").Range("I8").Value = "" Or ThisWorkbook.Sheets("Reportes").Range("M8").Value = "" Then
        MsgBox "Datos incompletos, por favor ingrese los datos antes de ejecutar.", vbExclamation
        Exit Sub
    End If

    ' Llama a cada una de las funciones
    DeactivateStuff
    CrearCarpetas
    Auditoria_Antiguedad
    ReactivateStuff
    
    MsgBox "Reporte finalizado. Lo podra encontrar en carpeta de auditorias", vbInformation

End Sub
    
    
    
Sub InicializarVariables()
'Definicion de las variables
    
    '------------ Variables originales del archivo -----------
    mes = ThisWorkbook.Sheets("Reportes").Range("N8").Text
    mes_2 = ThisWorkbook.Sheets("Reportes").Range("M12").Value
    Mes_texto = ThisWorkbook.Sheets("Reportes").Range("I12").Value
    A�o = ThisWorkbook.Sheets("Reportes").Range("I10").Value
    Fecha1 = ThisWorkbook.Sheets("Reportes").Range("I8").Value
    Fecha2 = ThisWorkbook.Sheets("Reportes").Range("M8").Value
    
    '------------ Rutas -----------
    ruta = ThisWorkbook.Path & "\"
    Ruta_A�o = ruta & A�o
    Ruta_Mes = Ruta_A�o & "\" & mes & ". " & Mes_texto
    

End Sub

Sub CrearCarpetas()
    
    'Creaci�n y validacion de las carpetas
    InicializarVariables
    '''''''
    ''A�O''
    '''''''
    Ruta_A�o = ruta & A�o
    If Dir(Ruta_A�o, vbDirectory + vbHidden) = "" Then
        'Comprueba que la carpeta no exista para crear el directorio.
        If Dir(Ruta_A�o & vbDirectory + vbHidden) = "" Then MkDir Ruta_A�o
    End If
    '''''''
    ''MES''
    '''''''
    Ruta_Mes = Ruta_A�o & "\" & mes & ". " & Mes_texto
    If Dir(Ruta_Mes, vbDirectory + vbHidden) = "" Then
        'Comprueba que la carpeta no exista para crear el directorio.
        If Dir(Ruta_Mes & vbDirectory + vbHidden) = "" Then MkDir Ruta_Mes
    End If
        
End Sub


Sub Auditoria_Antiguedad()
    
    InicializarVariables

'------------------- PRIMERA DESCARGA DE SAP Y ORGANIZAR EL ARCHIVO -------------------

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

    'Entra a la transacion PC00_M99_CWTR
    session.findById("wnd[0]/tbar[0]/okcd").Text = "/nPC00_M99_CWTR"
    session.findById("wnd[0]").sendVKey 0
    
    'Coloca la sociedad
    session.findById("wnd[0]/usr/ctxtPNPBUKRS-LOW").Text = "CO*"
    
    'Cambia fechas al mes que se este realizando la auditoria
    session.findById("wnd[0]/usr/ctxtBEGD_CAL").Text = Fecha1
    session.findById("wnd[0]/usr/ctxtENDD_CAL").Text = Fecha2
    
    'Coloca cc nom 1033 y Selecciona objeto
    session.findById("wnd[0]/usr/ctxtS_LGART-LOW").Text = "1033"
    session.findById("wnd[0]/usr/ctxtS_LGART-LOW").SetFocus
    session.findById("wnd[0]/usr/ctxtS_LGART-LOW").caretPosition = 4
    session.findById("wnd[0]/usr/btnLABEL02").press
    session.findById("wnd[1]/usr/tblSAPLSLFBFIELDCONTROL").getAbsoluteRow(0).Selected = True
    session.findById("wnd[1]/usr/tblSAPLSLFBFIELDCONTROL/txtFIELDTAB-FIELDTEXT[0,0]").caretPosition = 19
    session.findById("wnd[1]/usr/btn%_AUTOTEXT012").press
    session.findById("wnd[1]/tbar[0]/btn[0]").press
    
    
    'Ejecuta
    session.findById("wnd[0]/tbar[1]/btn[8]").press
    
    'Descarga el archivo
    session.findById("wnd[0]/mbar/menu[0]/menu[1]/menu[2]").Select
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").Select
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").SetFocus
    session.findById("wnd[1]/tbar[0]/btn[0]").press
    session.findById("wnd[1]/usr/ctxtDY_PATH").Text = Ruta_Mes
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = A�o & mes & "." & "AUD PRIMA ADM" & ".XLS"
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 16
    session.findById("wnd[1]/tbar[0]/btn[11]").press
    
    
    'Copia del documento
    Application.CutCopyMode = False
    Workbooks.Open Ruta_Mes & "\" & A�o & mes & "." & "AUD PRIMA ADM" & ".XLS"
    ActiveSheet.Cells.Select
    Selection.Copy
    Workbooks.Add
    ActiveSheet.Paste
    ActiveWorkbook.SaveAs Ruta_Mes & "\" & A�o & mes & "." & "AUD PRIMA ADM" & ".XLSX"
    ActiveWorkbook.Close SaveChanges:=True
    Workbooks(A�o & mes & "." & "AUD PRIMA ADM" & ".XLS").Close
    Kill Ruta_Mes & "\" & A�o & mes & "." & "AUD PRIMA ADM" & ".XLS"
    
    
    'Organizar documento
    Workbooks.Open Ruta_Mes & "\" & A�o & mes & "." & "AUD PRIMA ADM" & ".XLSX"
    Workbooks(A�o & mes & "." & "AUD PRIMA ADM" & ".XLSX").Activate
    Rows("1").Delete
    Rows("2").Delete
    Columns("A").Delete
    Columns("A:V").AutoFit
    ActiveSheet.Name = "CWTR"
    
    'Cambios de formato para la cantidad
    LastRow = ActiveSheet.Cells(ActiveSheet.Rows.Count, "A").End(xlUp).row
    Columns("T:T").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("T1").Value = "Cantidad"
    If LastRow >= 2 Then
        Range("T2:T" & LastRow).Formula = "=VALUE(SUBSTITUTE(RC[1], CHAR(160),""""))"
    End If
    Columns("T:T").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("U:U").Select
    Selection.Delete

    'Cambios de formato para el importe
    Columns("U:U").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("U1").Value = "Importe"
    If LastRow >= 2 Then
        Range("U2:U" & LastRow).Formula = "=VALUE(SUBSTITUTE(RC[1], CHAR(160),""""))"
    End If
    Columns("U:U").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("V:V").Select
    Selection.Delete
    Columns("U:U").NumberFormat = "$#,##0"
    Columns("A:V").AutoFit
    
    'Cambiar formato de fecha
    Dim CelsD As Range
    Dim UltimaFilaD As Long
    UltimaFilaD = ActiveSheet.Cells(Rows.Count, 1).End(xlUp).row
    
    If UltimaFilaD >= 2 Then
        For Each CelsD In ActiveSheet.Range("N2:N" & UltimaFilaD)
            If Len(CelsD.Value) = 10 And Mid(CelsD.Value, 3, 1) = "." And Mid(CelsD.Value, 6, 1) = "." Then
                CelsD.Value = DateSerial(Right(CelsD.Value, 4), Mid(CelsD.Value, 4, 2), Left(CelsD.Value, 2))
                CelsD.NumberFormat = "dd/mm/yyyy"
            End If
        Next CelsD
    End If
    
    Sheets.Add(After:=Sheets(Sheets.Count)).Name = "SALARIAL"
    ActiveWorkbook.Save
    

    'Trae lo del maestro
    'Abrir maestro del mes correspondiente
        Dim ruta_maestroA As String
        Dim MAESTRO As Workbook
        
        ' Solicitar al usuario que abra un archivo
        MsgBox "Por favor selecciene el archivo del Maestro de Activos mas reciente.", vbInformation
        ruta_maestroA = Application.GetOpenFilename("Archivos Excel (*.xls; *.xlsx), *.xls; *.xlsx")
        Application.AskToUpdateLinks = False
        
        ' Abre el reporte seleccionado por el usuario y copiia la primera hoja
        If ruta_maestroA <> "Falso" Then
            
            Set MAESTRO = Workbooks.Open(Filename:=ruta_maestroA, UpdateLinks:=0)
            MAESTRO.Activate
            Worksheets("SALARIAL").Activate
            
            'Quita filtros si los tiene
            If ActiveSheet.AutoFilterMode Then
                ActiveSheet.AutoFilterMode = False
            End If
            
            Rows("1:1").AutoFilter Field:=11, Criteria1:="=Adm", Operator:=xlOr, Criteria2:="=Adm var"
            LastRow = ActiveSheet.Cells(Rows.Count, 6).End(xlUp).row
            Range("A1:AK" & LastRow).Select
            Selection.Copy
        
        Else
            MsgBox "Operaci�n cancelada por el usuario.", vbInformation
        End If
    
    Workbooks(A�o & mes & "." & "AUD PRIMA ADM" & ".XLSX").Activate
    Sheets("SALARIAL").Activate
    Range("A1").PasteSpecial Paste:=xlPasteAll
    MAESTRO.Save
    MAESTRO.Close
    
    
    
'------------------- EMPIEZA A REALIZAR LA AUDITORIA -------------------
    
    
    'Activa el libro de la auditoria
    Workbooks(A�o & mes & "." & "AUD PRIMA ADM" & ".XLSX").Activate
    Sheets("SALARIAL").Activate
    Columns("C:D").Delete
    Columns("D:G").Delete
    Columns("F:O").Delete
    Columns("G:R").Delete
    Columns("I:I").NumberFormat = "$#,##0"
    Columns("F:F").NumberFormat = "dd/mm/yyyy"
    ActiveWorkbook.Save
    
    'Extraer el mes
    LastRow = ActiveSheet.Cells(ActiveSheet.Rows.Count, "I").End(xlUp).row
    Columns("G:G").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("G1").Value = "MES"
    If LastRow >= 2 Then
        Range("G2:G" & LastRow).Formula = "=MONTH(RC[-1])"
    End If
    Columns("G:G").NumberFormat = "0"
    Columns("G:G").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    'Condicional para eliminar los que no sean del periodo
    LastRow = ActiveSheet.Cells(ActiveSheet.Rows.Count, "J").End(xlUp).row
    For i = LastRow To 2 Step -1
        If Sheets("SALARIAL").Cells(i, "G").Value <> mes_2 Then
            Sheets("SALARIAL").Rows(i).Delete ' Eliminar la fila si la condici�n se cumple
        End If
    Next i
    Columns("A:J").AutoFit
    ActiveWorkbook.Save
    
    'Condicional para eliminar los que no sean ley 50
    LastRow = ActiveSheet.Cells(ActiveSheet.Rows.Count, "J").End(xlUp).row
    For i = LastRow To 2 Step -1
        If Sheets("SALARIAL").Cells(i, "C").Value <> "Ley 50" Then
            Sheets("SALARIAL").Rows(i).Delete ' Eliminar la fila si la condici�n se cumple
        End If
    Next i
    Columns("A:J").AutoFit
    ActiveWorkbook.Save
    
    'Formulas
    LastRow = ActiveSheet.Cells(ActiveSheet.Rows.Count, "J").End(xlUp).row
    Range("K2:K" & LastRow) = "=+VLOOKUP(RC[-10],CWTR!C[-10]:C[10],21,FALSE)"
    Range("L2:L" & LastRow) = "=+RC[-2]-RC[-1]"
    
    'Formato de las celdas de los campos
    Range("K1").Value = "AUDITORIA"
    Range("L1").Value = "DIFERENCIA"
    Columns("K:K").NumberFormat = "$#,##0"
    Columns("L:L").NumberFormat = "$#,##0"
    With Range("A1:L1")
        .Interior.Color = RGB(174, 214, 241)
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
    End With
    Columns("A:L").AutoFit
    
    Sheets("CWTR").Activate
    With Range("A1:L1")
        .Interior.Color = RGB(213, 245, 227)
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
    End With
    Columns("A:V").AutoFit
    Sheets("SALARIAL").Activate
    
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



