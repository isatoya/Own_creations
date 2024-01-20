Attribute VB_Name = "M�dulo2"

Sub SAP_ZHR929()

    Dim A�o As Integer
    Dim fechaEnero As Date
    Dim fechaDiciembre As Date
    Dim ruta As String
    ruta = ThisWorkbook.Path
    A�o = UserForm1.ComboBox1.Value 'a�o seleccionado del combobox
    Dim ruta2 As String
    ruta2 = ruta & "\" & CStr(A�o) & "\1. PW PTU"
    

    'Primer y ultimo dia del a�o
    fechaEnero = DateSerial(A�o, 1, 1) 'Enero 1
    fechaDiciembre = DateSerial(A�o + 1, 1, 1) - 1 'Diciembre 31

    'fecha formato para sap
    Dim fecha1 As String
    Dim fecha2 As String

    fecha1 = Format(fechaEnero, "dd.mm.yyyy")
    fecha2 = Format(fechaDiciembre, "dd.mm.yyyy")

    'Conexion con SAP
    Application.DisplayAlerts = False
    Set SapGuiAuto = GetObject("SAPGUI")
    Set App = SapGuiAuto.GetScriptingEngine
    Set Connection = App.Children(0)
    Set session = Connection.Children(0)

    'Vuelve a la pantalla inicial
    session.findById("wnd[0]/tbar[0]/okcd").Text = "/n"
    session.findById("wnd[0]").sendVKey 0

    'Entra a la transacccion ZHR929
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]/tbar[0]/okcd").Text = "/nZHR929"
    
    'Busca la variante
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]/tbar[1]/btn[17]").press
    session.findById("wnd[1]/usr/txtENAME-LOW").Text = "LGIRALD2"
    session.findById("wnd[1]/usr/txtENAME-LOW").SetFocus
    session.findById("wnd[1]/usr/txtENAME-LOW").caretPosition = 8
    session.findById("wnd[1]/tbar[0]/btn[8]").press
    
    'Cambia fechas
    session.findById("wnd[0]/usr/ctxtPNPBEGDA").Text = fecha1
    session.findById("wnd[0]/usr/ctxtPNPENDDA").Text = fecha2
    session.findById("wnd[0]/usr/ctxtPNPENDDA").SetFocus
    
    'Ejecuta
    session.findById("wnd[0]/usr/ctxtPNPENDDA").caretPosition = 8
    session.findById("wnd[0]").sendVKey 0
    
    'Descarga como fichero
    session.findById("wnd[0]/tbar[1]/btn[8]").press
    session.findById("wnd[0]/mbar/menu[0]/menu[3]/menu[2]").Select
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").Select
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").SetFocus
    session.findById("wnd[1]/tbar[0]/btn[0]").press
    session.findById("wnd[1]/usr/ctxtDY_PATH").Text = ruta2
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = "ZHR929.xls"
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 10
    session.findById("wnd[1]/tbar[0]/btn[11]").press
    
    'Cambia el documento de formato
    Workbooks.Open ruta2 & "\ZHR929" & ".xls"
    ActiveSheet.Cells.Select
    Selection.Copy
    Workbooks.Add
    ActiveSheet.Paste
    ActiveWorkbook.SaveAs ruta2 & "\ZHR929" & ".xlsx"
    ActiveWorkbook.Close SaveChanges:=True
    ArchivoA = ActiveWorkbook.Name
    Workbooks("ZHR929" & ".xls").Close
    Kill ruta2 & "\ZHR929" & ".xls"
    
    'Cambios de formato
    Workbooks.Open ruta2 & "\ZHR929" & ".xlsx"
    Organizar_documentos_929


End Sub

Sub SAP_ZHRMX27()

    Dim A�o As Integer
    Dim fechaEnero As Date
    Dim fechaDiciembre As Date
    Dim ruta As String
    ruta = ThisWorkbook.Path
    A�o = UserForm1.ComboBox1.Value 'a�o seleccionado del combobox
    Dim ruta2 As String
    ruta2 = ruta & "\" & CStr(A�o) & "\1. PW PTU"

    'Primer y ultimo dia del a�o
    fechaEnero = DateSerial(A�o, 1, 1) 'Enero 1
    fechaDiciembre = DateSerial(A�o + 1, 1, 1) - 1 'Diciembre 31

    'fecha formato para sap
    Dim fecha1 As String
    Dim fecha2 As String

    fecha1 = Format(fechaEnero, "dd.mm.yyyy")
    fecha2 = Format(fechaDiciembre, "dd.mm.yyyy")

    'Conexion con SAP
    Application.DisplayAlerts = False
    Set SapGuiAuto = GetObject("SAPGUI")
    Set App = SapGuiAuto.GetScriptingEngine
    Set Connection = App.Children(0)
    Set session = Connection.Children(0)
    
    'Vuelve a la pantalla inicial
    session.findById("wnd[0]/tbar[0]/okcd").Text = "/n"
    session.findById("wnd[0]").sendVKey 0

    'Entra a la transaccion
    session.findById("wnd[0]/tbar[0]/okcd").Text = "/nZHRMX27"
    session.findById("wnd[0]").sendVKey 0
    
    'Busca la variante
    session.findById("wnd[0]/tbar[1]/btn[17]").press
    session.findById("wnd[1]/usr/txtV-LOW").Text = "FLEX PTU"
    session.findById("wnd[1]/usr/txtENAME-LOW").Text = ""
    session.findById("wnd[1]/usr/txtV-LOW").caretPosition = 8
    session.findById("wnd[1]/tbar[0]/btn[8]").press
    
    'Pega los numeros de los empleados
    session.findById("wnd[0]/usr/btn%_PNPPERNR_%_APP_%-VALU_PUSH").press
    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE").Columns.elementAt(1).Selected = True
    session.findById("wnd[1]/tbar[0]/btn[16]").press
    
     'Parte del codigo que copia los numeros de usuario de la ZHR929
        Application.CutCopyMode = False
        Dim lastRow As Long
        
        If Dir(ruta2 & "\ZHR929.xlsx") = "" Then
            MsgBox "El archivo ZHR929.xlsx no se encontr� en la ruta especificada.", vbExclamation
            Exit Sub
        End If
        
        Workbooks.Open Filename:=ruta2 & "\ZHR929.xlsx"
        Application.CutCopyMode = True
        Workbooks("ZHR929.xlsx").Activate
        Sheets("Hoja1").Activate
        Range("A1").Select
        lastRow = Cells(Rows.Count, "A").End(xlUp).Row
        
        If lastRow >= 2 Then
            Columns("A").Resize(lastRow - 1).Offset(1, 0).Copy
        Else
            MsgBox "No hay datos en la columna A.", vbExclamation
        End If
    
    session.findById("wnd[1]/tbar[0]/btn[24]").press
    session.findById("wnd[1]/tbar[0]/btn[8]").press
    session.findById("wnd[0]/tbar[1]/btn[8]").press
    
    'Exporta el fichero
    session.findById("wnd[0]/mbar/menu[0]/menu[3]/menu[2]").Select
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").Select
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").SetFocus
    session.findById("wnd[1]/tbar[0]/btn[0]").press
    session.findById("wnd[1]/usr/ctxtDY_PATH").Text = ruta2
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = "ZHRMX27.XLS"
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 6
    session.findById("wnd[1]/tbar[0]/btn[11]").press

    'Organiza la informaci�n extra�da de SAP en formato xlsx
    Workbooks("ZHR929.xlsx").Close SaveChanges:=False
    Workbooks.Open ruta2 & "\ZHRMX27" & ".xls"
    ActiveSheet.Cells.Select
    Selection.Copy
    Workbooks.Add
    ActiveSheet.Paste
    ActiveWorkbook.SaveAs ruta2 & "\ZHRMX27" & ".xlsx"
    ActiveWorkbook.Close SaveChanges:=True
    ArchivoA = ActiveWorkbook.Name
    Workbooks("ZHRMX27" & ".xls").Close
    Kill ruta2 & "\ZHRMX27" & ".xls"
    
    'Cambios de formato
    Workbooks.Open ruta2 & "\ZHRMX27" & ".xlsx"
    Organizar_documentos_ZHRMX27


End Sub

Sub SAP_ZPYMX025_DIASPTULG2()
'PRIEMRA VARIANTE: DIAS PTU LG 2

    Dim A�o As Integer
    Dim fechaEnero As Date
    Dim fechaDiciembre As Date
    Dim ruta As String
    ruta = ThisWorkbook.Path
    A�o = UserForm1.ComboBox1.Value 'a�o seleccionado del combobox
    Dim ruta2 As String
    ruta2 = ruta & "\" & CStr(A�o) & "\1. PW PTU"

    'Primer y ultimo dia del a�o
    fechaEnero = DateSerial(A�o, 1, 1) 'Enero 1
    fechaDiciembre = DateSerial(A�o + 1, 1, 1) - 1 'Diciembre 31

    'fecha formato para sap
    Dim fecha1 As String
    Dim fecha2 As String

    fecha1 = Format(fechaEnero, "dd.mm.yyyy")
    fecha2 = Format(fechaDiciembre, "dd.mm.yyyy")

    'Conexion con SAP
    Application.DisplayAlerts = False
    Set SapGuiAuto = GetObject("SAPGUI")
    Set App = SapGuiAuto.GetScriptingEngine
    Set Connection = App.Children(0)
    Set session = Connection.Children(0)
    
    'Vuelve a la pantalla inicial
    session.findById("wnd[0]/tbar[0]/okcd").Text = "/n"
    session.findById("wnd[0]").sendVKey 0

    'Entra a la transaccion
    session.findById("wnd[0]/tbar[0]/okcd").Text = "/nZPYMX025"
    session.findById("wnd[0]").sendVKey 0
    
    'Busca la variante
    session.findById("wnd[0]/tbar[1]/btn[17]").press
    session.findById("wnd[1]/usr/txtV-LOW").Text = "PTU DIAS LG 2"
    session.findById("wnd[1]/usr/txtENAME-LOW").Text = ""
    session.findById("wnd[1]/usr/txtV-LOW").caretPosition = 13
    session.findById("wnd[1]/tbar[0]/btn[8]").press
    
    'Cambia fechas
    session.findById("wnd[0]/usr/ctxtBEGD_CAL").Text = fecha1
    session.findById("wnd[0]/usr/ctxtENDD_CAL").Text = fecha2
    session.findById("wnd[0]/usr/ctxtENDD_CAL").SetFocus
    
    'Pone numeros de personal
    session.findById("wnd[0]/usr/ctxtENDD_CAL").caretPosition = 10
    session.findById("wnd[0]/usr/btn%_PNPPERNR_%_APP_%-VALU_PUSH").press
    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE").Columns.elementAt(1).Selected = True
    session.findById("wnd[1]/tbar[0]/btn[16]").press ' Borra
    
    'Parte del codigo que copia los numeros de usuario de la ZHR929
        Application.CutCopyMode = False
        Dim lastRow As Long
        
        If Dir(ruta2 & "\ZHR929.xlsx") = "" Then
            MsgBox "El archivo ZHR929.xlsx no se encontr� en la ruta especificada.", vbExclamation
            Exit Sub
        End If
        
        Workbooks.Open Filename:=ruta2 & "\ZHR929.xlsx"
        Application.CutCopyMode = True
        Workbooks("ZHR929.xlsx").Activate
        Sheets("Hoja1").Activate
        Range("A1").Select
        lastRow = Cells(Rows.Count, "A").End(xlUp).Row
        
        If lastRow >= 2 Then
            Columns("A").Resize(lastRow - 1).Offset(1, 0).Copy
        Else
            MsgBox "No hay datos en la columna A.", vbExclamation
        End If
    
    session.findById("wnd[1]/tbar[0]/btn[24]").press ' Pega
    session.findById("wnd[1]/tbar[0]/btn[8]").press
    session.findById("wnd[0]/tbar[1]/btn[8]").press
    
    'Exporta como fichero
    session.findById("wnd[0]/mbar/menu[0]/menu[1]/menu[2]").Select
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").Select
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").SetFocus
    session.findById("wnd[1]/tbar[0]/btn[0]").press
    session.findById("wnd[1]/usr/ctxtDY_PATH").Text = ruta2
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = "ZPYMX025.xls"
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 6
    session.findById("wnd[1]/tbar[0]/btn[11]").press

    'Organiza la informaci�n extra�da de SAP en formato xlsx
    Workbooks("ZHR929.xlsx").Close SaveChanges:=False
    Workbooks.Open ruta2 & "\ZPYMX025" & ".xls"
    ActiveSheet.Cells.Select
    Selection.Copy
    Workbooks.Add
    ActiveSheet.Paste
    ActiveWorkbook.SaveAs ruta2 & "\ZPYMX025" & ".xlsx"
    ActiveWorkbook.Close SaveChanges:=True
    ArchivoA = ActiveWorkbook.Name
    Workbooks("ZPYMX025" & ".xls").Close
    Kill ruta2 & "\ZPYMX025" & ".xls"
    
    
    'Cambios de formato
    Workbooks.Open ruta2 & "\ZPYMX025" & ".xlsx"
    Organizar_documentos_ZPYMX025
    
End Sub

Sub SAP_ZPYMX025_PESOSPTULGV2()
'VARIANTE 2 : PESOS PTU LC V2

    Dim A�o As Integer
    Dim fechaEnero As Date
    Dim fechaDiciembre As Date
    Dim ruta As String
    ruta = ThisWorkbook.Path
    A�o = UserForm1.ComboBox1.Value 'a�o seleccionado del combobox
    Dim ruta2 As String
    ruta2 = ruta & "\" & CStr(A�o) & "\1. PW PTU"

    'Primer y ultimo dia del a�o
    fechaEnero = DateSerial(A�o, 1, 1) 'Enero 1
    fechaDiciembre = DateSerial(A�o + 1, 1, 1) - 1 'Diciembre 31

    'fecha formato para sap
    Dim fecha1 As String
    Dim fecha2 As String

    fecha1 = Format(fechaEnero, "dd.mm.yyyy")
    fecha2 = Format(fechaDiciembre, "dd.mm.yyyy")

    'Conexion con SAP
    Application.DisplayAlerts = False
    Set SapGuiAuto = GetObject("SAPGUI")
    Set App = SapGuiAuto.GetScriptingEngine
    Set Connection = App.Children(0)
    Set session = Connection.Children(0)

    'Vuelve a la pantalla inicial
    session.findById("wnd[0]/tbar[0]/okcd").Text = "/n"
    session.findById("wnd[0]").sendVKey 0

    'Entra a la transaccion
    session.findById("wnd[0]/tbar[0]/okcd").Text = "/nZPYMX025"
    session.findById("wnd[0]").sendVKey 0
    
    'Busca la variante
    session.findById("wnd[0]/tbar[1]/btn[17]").press
    session.findById("wnd[1]/usr/txtV-LOW").Text = "PTU PESOS LG 2"
    session.findById("wnd[1]/usr/txtENAME-LOW").Text = ""
    session.findById("wnd[1]/usr/txtV-LOW").caretPosition = 13
    session.findById("wnd[1]/tbar[0]/btn[8]").press
    
    'Cambia fechas
    session.findById("wnd[0]/usr/ctxtBEGD_CAL").Text = fecha1
    session.findById("wnd[0]/usr/ctxtENDD_CAL").Text = fecha2
    session.findById("wnd[0]/usr/ctxtENDD_CAL").SetFocus
    
    'Pone numeros de personal
    session.findById("wnd[0]/usr/ctxtENDD_CAL").caretPosition = 10
    session.findById("wnd[0]/usr/btn%_PNPPERNR_%_APP_%-VALU_PUSH").press
    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE").Columns.elementAt(1).Selected = True
    session.findById("wnd[1]/tbar[0]/btn[16]").press ' Borra
    
    'Parte del codigo que copia los numeros de usuario de la ZHR929
        Application.CutCopyMode = False
        Dim lastRow As Long
        
        If Dir(ruta2 & "\ZHR929.xlsx") = "" Then
            MsgBox "El archivo ZHR929.xlsx no se encontr� en la ruta especificada.", vbExclamation
            Exit Sub
        End If
        
        Workbooks.Open Filename:=ruta2 & "\ZHR929.xlsx"
        Application.CutCopyMode = True
        Workbooks("ZHR929.xlsx").Activate
        Sheets("Hoja1").Activate
        Range("A1").Select
        lastRow = Cells(Rows.Count, "A").End(xlUp).Row
        
        If lastRow >= 2 Then
            Columns("A").Resize(lastRow - 1).Offset(1, 0).Copy
        Else
            MsgBox "No hay datos en la columna A.", vbExclamation
        End If
    
    session.findById("wnd[1]/tbar[0]/btn[24]").press ' Pega
    session.findById("wnd[1]/tbar[0]/btn[8]").press
    session.findById("wnd[0]/tbar[1]/btn[8]").press
    
    'Exporta como fichero
    session.findById("wnd[0]/mbar/menu[0]/menu[1]/menu[2]").Select
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").Select
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").SetFocus
    session.findById("wnd[1]/tbar[0]/btn[0]").press
    session.findById("wnd[1]/usr/ctxtDY_PATH").Text = ruta2
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = "ZPYMX025_V2.xls"
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 6
    session.findById("wnd[1]/tbar[0]/btn[11]").press
     
    'Organiza la informaci�n extra�da de SAP en formato xlsx
    Workbooks("ZHR929.xlsx").Close SaveChanges:=False
    Workbooks.Open ruta2 & "\ZPYMX025_V2" & ".xls"
    ActiveSheet.Cells.Select
    Selection.Copy
    Workbooks.Add
    ActiveSheet.Paste
    ActiveWorkbook.SaveAs ruta2 & "\ZPYMX025_V2" & ".xlsx"
    ActiveWorkbook.Close SaveChanges:=True
    ArchivoA = ActiveWorkbook.Name
    Workbooks("ZPYMX025_V2" & ".xls").Close
    Kill ruta2 & "\ZPYMX025_V2" & ".xls"
    
    'Cambios de formato
    Workbooks.Open ruta2 & "\ZPYMX025_V2" & ".xlsx"
    Organizar_documentos_ZPYMX025

End Sub

Sub SAP_ZPYMX025_PTUAUSENTISMO()
'VARIANTE 3: PTU AUSENTIMOS

    Dim A�o As Integer
    Dim fechaEnero As Date
    Dim fechaDiciembre As Date
    Dim ruta As String
    ruta = ThisWorkbook.Path
    A�o = UserForm1.ComboBox1.Value 'a�o seleccionado del combobox
    Dim ruta2 As String
    ruta2 = ruta & "\" & CStr(A�o) & "\1. PW PTU"

    'Primer y ultimo dia del a�o
    fechaEnero = DateSerial(A�o, 1, 1) 'Enero 1
    fechaDiciembre = DateSerial(A�o + 1, 1, 1) - 1 'Diciembre 31

    'fecha formato para sap
    Dim fecha1 As String
    Dim fecha2 As String

    fecha1 = Format(fechaEnero, "dd.mm.yyyy")
    fecha2 = Format(fechaDiciembre, "dd.mm.yyyy")

    'Conexion con SAP
    Application.DisplayAlerts = False
    Set SapGuiAuto = GetObject("SAPGUI")
    Set App = SapGuiAuto.GetScriptingEngine
    Set Connection = App.Children(0)
    Set session = Connection.Children(0)

    'Vuelve a la pantalla inicial
    session.findById("wnd[0]/tbar[0]/okcd").Text = "/n"
    session.findById("wnd[0]").sendVKey 0

    'Entra a la transaccion
    session.findById("wnd[0]/tbar[0]/okcd").Text = "/nZPYMX025"
    session.findById("wnd[0]").sendVKey 0
    
    'Busca la variante
    session.findById("wnd[0]/tbar[1]/btn[17]").press
    session.findById("wnd[1]/usr/txtV-LOW").Text = "PTU AUSENTISMO"
    session.findById("wnd[1]/usr/txtENAME-LOW").Text = ""
    session.findById("wnd[1]/usr/txtV-LOW").caretPosition = 13
    session.findById("wnd[1]/tbar[0]/btn[8]").press
    
    'Cambia fechas
    session.findById("wnd[0]/usr/ctxtBEGD_CAL").Text = fecha1
    session.findById("wnd[0]/usr/ctxtENDD_CAL").Text = fecha2
    session.findById("wnd[0]/usr/ctxtENDD_CAL").SetFocus
    
    'Pone numeros de personal
    session.findById("wnd[0]/usr/ctxtENDD_CAL").caretPosition = 10
    session.findById("wnd[0]/usr/btn%_PNPPERNR_%_APP_%-VALU_PUSH").press
    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE").Columns.elementAt(1).Selected = True
    session.findById("wnd[1]/tbar[0]/btn[16]").press ' Borra
    
    'Parte del codigo que copia los numeros de usuario de la ZHR929
        Application.CutCopyMode = False
        Dim lastRow As Long
        
        If Dir(ruta2 & "\ZHR929.xlsx") = "" Then
            MsgBox "El archivo ZHR929.xlsx no se encontr� en la ruta especificada.", vbExclamation
            Exit Sub
        End If
        
        Workbooks.Open Filename:=ruta2 & "\ZHR929.xlsx"
        Application.CutCopyMode = True
        Workbooks("ZHR929.xlsx").Activate
        Sheets("Hoja1").Activate
        Range("A1").Select
        lastRow = Cells(Rows.Count, "A").End(xlUp).Row
        
        If lastRow >= 2 Then
            Columns("A").Resize(lastRow - 1).Offset(1, 0).Copy
        Else
            MsgBox "No hay datos en la columna A.", vbExclamation
        End If
    
    session.findById("wnd[1]/tbar[0]/btn[24]").press ' Pega
    session.findById("wnd[1]/tbar[0]/btn[8]").press
    session.findById("wnd[0]/tbar[1]/btn[8]").press
    
    'Exporta como fichero
    session.findById("wnd[0]/mbar/menu[0]/menu[1]/menu[2]").Select
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").Select
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").SetFocus
    session.findById("wnd[1]/tbar[0]/btn[0]").press
    session.findById("wnd[1]/usr/ctxtDY_PATH").Text = ruta2
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = "ZPYMX025_AUSENTISMOS.xls"
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 6
    session.findById("wnd[1]/tbar[0]/btn[11]").press
    
    'Organiza la informaci�n extra�da de SAP en formato xlsx
    Workbooks("ZHR929.xlsx").Close SaveChanges:=False
    Workbooks.Open ruta2 & "\ZPYMX025_AUSENTISMOS" & ".xls"
    ActiveSheet.Cells.Select
    Selection.Copy
    Workbooks.Add
    ActiveSheet.Paste
    ActiveWorkbook.SaveAs ruta2 & "\ZPYMX025_AUSENTISMOS" & ".xlsx"
    ActiveWorkbook.Close SaveChanges:=True
    ArchivoA = ActiveWorkbook.Name
    Workbooks("ZPYMX025_AUSENTISMOS" & ".xls").Close
    Kill ruta2 & "\ZPYMX025_AUSENTISMOS" & ".xls"
    
    'Cambios de formato
    Workbooks.Open ruta2 & "\ZPYMX025_AUSENTISMOS" & ".xlsx"
    Organizar_documentos_ZPYMX025

End Sub

Sub Organizar_documentos_929()

    Dim A�o As Integer
    Dim ruta As String
    Dim ruta2 As String
    A�o = UserForm1.ComboBox1.Value 'a�o seleccionado del combobox
    ruta = ThisWorkbook.Path
    ruta2 = ruta & "\" & CStr(A�o) & "\1. PW PTU"
    
    Rows("1:4").Delete
    Rows("2").Delete
    Columns("A").Delete
    ActiveWorkbook.Save
    ActiveWorkbook.Close
    
End Sub

Sub Organizar_documentos_ZPYMX025()

    Dim A�o As Integer
    Dim ruta As String
    Dim ruta2 As String
    A�o = UserForm1.ComboBox1.Value 'a�o seleccionado del combobox
    ruta = ThisWorkbook.Path
    ruta2 = ruta & "\" & CStr(A�o) & "\1. PW PTU"
    
    Rows("1").Delete
    Rows("2").Delete
    Columns("A").Delete
    ActiveWorkbook.Save
    ActiveWorkbook.Close
    
End Sub

Sub Organizar_documentos_ZHRMX27()

    Dim A�o As Integer
    Dim ruta As String
    Dim ruta2 As String
    A�o = UserForm1.ComboBox1.Value 'a�o seleccionado del combobox
    ruta = ThisWorkbook.Path
    ruta2 = ruta & "\" & CStr(A�o) & "\1. PW PTU"
    
    Rows("1:7").Delete
    Rows("2").Delete
    Columns("A").Delete
    ActiveWorkbook.Save
    ActiveWorkbook.Close
    
End Sub

Sub Crear_Carpetas()
    Dim A�o As Integer
    A�o = UserForm1.ComboBox1.Value

    'ruta principal de la macro
    Dim ruta As String
    ruta = ThisWorkbook.Path

    ' Crear la carpeta del a�o si no existe
    If Len(Dir(ruta & "\" & CStr(A�o), vbDirectory)) = 0 Then
        MkDir ruta & "\" & CStr(A�o)
    End If

    ' Crear la subcarpeta "1. PW PTU" dentro de la carpeta del a�o
    If Len(Dir(ruta & "\" & CStr(A�o) & "\1. PW PTU", vbDirectory)) = 0 Then
        MkDir ruta & "\" & CStr(A�o) & "\1. PW PTU"
    End If

    ' Crear la subcarpeta "3. SIND MAL ALTO - DIAS PESOS" dentro de la carpeta del a�o
    If Len(Dir(ruta & "\" & CStr(A�o) & "\3. SIND MAL ALTO - DIAS PESOS", vbDirectory)) = 0 Then
        MkDir ruta & "\" & CStr(A�o) & "\3. SIND MAL ALTO - DIAS PESOS"
    End If
    
    ' Crear la subcarpeta "4. FACTOR DIAS PESOS" dentro de la carpeta del a�o
    If Len(Dir(ruta & "\" & CStr(A�o) & "\4. FACTOR DIAS PESOS", vbDirectory)) = 0 Then
        MkDir ruta & "\" & CStr(A�o) & "\4. FACTOR DIAS PESOS"
    End If
    
End Sub
