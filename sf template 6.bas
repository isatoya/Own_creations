Attribute VB_Name = "Template"
Sub Template()

If IsEmpty(ThisWorkbook.Sheets("JobInfoImportTemplate").Range("G6").Value) Then
MsgBox "Es necesario completar la información de los códigos en la hoja 'Position'.", vbExclamation, "Falta Información"
Else

If IsEmpty(ThisWorkbook.Sheets("Positions Report").Range("A3").Value) Or IsEmpty(ThisWorkbook.Sheets("Positions Report").Range("A4").Value) Then
MsgBox "Es necesario pegar los datos en la hoja “Positions Report”.", vbExclamation, "Falta Información"
Else

If IsEmpty(ThisWorkbook.Sheets("Total").Range("A3").Value) Or IsEmpty(ThisWorkbook.Sheets("Total").Range("A4").Value) Then
MsgBox "Es necesario pegar los datos en la hoja “Total”.", vbExclamation, "Falta Información"
Else

'entrar a los las otra funciones
Cambio_colum
BuscarV
EliminarBoton
Guardar


End If
End If
End If

End Sub

Sub Cambio_colum()

'Cambiar las columnas de la hoja de positions reports y Total de orden

Worksheets("Positions Report").Activate
Columns("A:A").Select
Selection.Insert Shift:=xlToLeft, CopyOrigin:=xlFormatFromLeftOrAbove
Columns("E:E").Select
Selection.Cut
Columns("A:A").Select
ActiveSheet.Paste
Columns("E:E").Select
Selection.Delete Shift:=xlToLeft
Columns("A:A").Select
Selection.NumberFormat = "General"

Worksheets("Total").Activate
Columns("A:A").Select
Selection.Insert Shift:=xlToLeft, CopyOrigin:=xlFormatFromLeftOrAbove
Columns("AE:AE").Select
Selection.Cut
Columns("A:A").Select
ActiveSheet.Paste
Columns("AE:AE").Select
Selection.Delete Shift:=xlToLeft
Columns("A:A").Select
Selection.NumberFormat = "General"

End Sub
Sub BuscarV()

'Encuentra la última fila en la columna G
Dim UltimaFila As Long
Sheets("JobInfoImportTemplate").Activate
UltimaFila = Sheets("JobInfoImportTemplate").Cells(Sheets("JobInfoImportTemplate").Rows.Count, "G").End(xlUp).Row

'Buscarv

Sheets("JobInfoImportTemplate").Range("D7:D" & UltimaFila).NumberFormat = "General"
ActiveSheet.Range("D7").Select
ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[3],Total!C[-3]:C[70],2,0)"
Selection.AutoFill Destination:=Range("D7:D" & UltimaFila)
ActiveSheet.Range(Selection, Selection.End(xlDown)).Select

Sheets("JobInfoImportTemplate").Range("H7:H" & UltimaFila).NumberFormat = "General"
ActiveSheet.Range("H7").Select
ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-1],Total!C[-7]:C[66],33,0)"
Selection.AutoFill Destination:=Range("H7:H" & UltimaFila)
ActiveSheet.Range(Selection, Selection.End(xlDown)).Select

Sheets("JobInfoImportTemplate").Range("I7:I" & UltimaFila).NumberFormat = "General"
ActiveSheet.Range("I7").Select
ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-2],Total!C[-8]:C[65],35,0)"
Selection.AutoFill Destination:=Range("I7:I" & UltimaFila)
ActiveSheet.Range(Selection, Selection.End(xlDown)).Select

Sheets("JobInfoImportTemplate").Range("J7:J" & UltimaFila).NumberFormat = "General"
ActiveSheet.Range("J7").Select
ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-3],'Positions Report'!C[-9]:C[-2],8,0)"
Selection.AutoFill Destination:=Range("J7:J" & UltimaFila)
ActiveSheet.Range(Selection, Selection.End(xlDown)).Select

Sheets("JobInfoImportTemplate").Range("K7:K" & UltimaFila).NumberFormat = "General"
ActiveSheet.Range("K7").Select
ActiveCell.FormulaR1C1 = "=VLOOKUP(VLOOKUP(RC[-4],Total!C[-10]:C[63],67,0),'De-Para'!C[-7]:C[-6],2,0)"
Selection.AutoFill Destination:=Range("K7:K" & UltimaFila)
ActiveSheet.Range(Selection, Selection.End(xlDown)).Select

Sheets("JobInfoImportTemplate").Range("L7:L" & UltimaFila).NumberFormat = "General"
ActiveSheet.Range("L7").Select
ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-5],Total!C[-11]:C[62],37,0)"
Selection.AutoFill Destination:=Range("L7:L" & UltimaFila)
ActiveSheet.Range(Selection, Selection.End(xlDown)).Select

Sheets("JobInfoImportTemplate").Range("M7:M" & UltimaFila).NumberFormat = "General"
ActiveSheet.Range("M7").Select
ActiveCell.FormulaR1C1 = "=MID(CONCATENATE(""L"",VLOOKUP(RC[-6],'Positions Report'!C[-12]:C[32],25,0)),1,7)"
Selection.AutoFill Destination:=Range("M7:M" & UltimaFila)
ActiveSheet.Range(Selection, Selection.End(xlDown)).Select

Sheets("JobInfoImportTemplate").Range("N7:N" & UltimaFila).NumberFormat = "General"
ActiveSheet.Range("N7").Select
ActiveCell.FormulaR1C1 = "=VLOOKUP(MID(CONCATENATE(""L"",VLOOKUP(RC[-7],'Positions Report'!C[-13]:C[31],25,0)),1,7),'De-Para'!C[5]:C[6],2,0)"
Selection.AutoFill Destination:=Range("N7:N" & UltimaFila)
ActiveSheet.Range(Selection, Selection.End(xlDown)).Select

Sheets("JobInfoImportTemplate").Range("O7:O" & UltimaFila).NumberFormat = "General"
ActiveSheet.Range("O7").Select
ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-8],Total!C[-14]:C[59],8,0)"
Selection.AutoFill Destination:=Range("O7:O" & UltimaFila)
ActiveSheet.Range(Selection, Selection.End(xlDown)).Select

Sheets("JobInfoImportTemplate").Range("P7:P" & UltimaFila).NumberFormat = "General"
ActiveSheet.Range("P7").Select
ActiveCell.FormulaR1C1 = "=VLOOKUP(VLOOKUP(RC[-9],Total!C[-15]:C[58],63,0),'De-Para'!C[-15]:C[-14],2,0)"
Selection.AutoFill Destination:=Range("P7:P" & UltimaFila)
ActiveSheet.Range(Selection, Selection.End(xlDown)).Select

Sheets("JobInfoImportTemplate").Range("Q7:Q" & UltimaFila).NumberFormat = "General"
ActiveSheet.Range("Q7").Select
ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-10],'Positions Report'!C[-16]:C[28],10,0)"
Selection.AutoFill Destination:=Range("Q7:Q" & UltimaFila)
ActiveSheet.Range(Selection, Selection.End(xlDown)).Select

Sheets("JobInfoImportTemplate").Range("R7:R" & UltimaFila).NumberFormat = "General"
ActiveSheet.Range("R7").Select
ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-11],'Positions Report'!C[-17]:C[27],29,0)"
Selection.AutoFill Destination:=Range("R7:R" & UltimaFila)
ActiveSheet.Range(Selection, Selection.End(xlDown)).Select

Sheets("JobInfoImportTemplate").Range("T7:T" & UltimaFila).NumberFormat = "General"
ActiveSheet.Range("T7").Select
ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-13],Total!C[-19]:C[54],25,0)"
Selection.AutoFill Destination:=Range("T7:T" & UltimaFila)
ActiveSheet.Range(Selection, Selection.End(xlDown)).Select

Sheets("JobInfoImportTemplate").Range("U7:U" & UltimaFila).NumberFormat = "General"
ActiveSheet.Range("U7").Select
ActiveCell.FormulaR1C1 = "=VLOOKUP(VLOOKUP(RC[-14],Total!C[-20]:C[53],66,0),'De-Para'!C[-11]:C[-10],2,0)"
Selection.AutoFill Destination:=Range("U7:U" & UltimaFila)
ActiveSheet.Range(Selection, Selection.End(xlDown)).Select

Sheets("JobInfoImportTemplate").Range("V7:V" & UltimaFila).NumberFormat = "General"
ActiveSheet.Range("V7").Select
ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-15],Total!C[-21]:C[52],65,0)"
Selection.AutoFill Destination:=Range("V7:V" & UltimaFila)
ActiveSheet.Range(Selection, Selection.End(xlDown)).Select

Sheets("JobInfoImportTemplate").Range("W7:W" & UltimaFila).NumberFormat = "General"
ActiveSheet.Range("W7").Select
ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-16],Total!C[-22]:C[51],27,0)"
Selection.AutoFill Destination:=Range("W7:W" & UltimaFila)
ActiveSheet.Range(Selection, Selection.End(xlDown)).Select

Sheets("JobInfoImportTemplate").Range("X7:X" & UltimaFila).NumberFormat = "General"
ActiveSheet.Range("X7").Select
ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-17],Total!C[-23]:C[50],61,0)"
Selection.AutoFill Destination:=Range("X7:X" & UltimaFila)
ActiveSheet.Range(Selection, Selection.End(xlDown)).Select

ActiveSheet.Range("AA7:AA" & UltimaFila).Value = "Yes"

Sheets("JobInfoImportTemplate").Range("AB7:AB" & UltimaFila).NumberFormat = "General"
ActiveSheet.Range("AB7").Select
ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-21],Total!C[-27]:C[46],50,0)"
Selection.AutoFill Destination:=Range("AB7:AB" & UltimaFila)
ActiveSheet.Range(Selection, Selection.End(xlDown)).Select

Sheets("JobInfoImportTemplate").Range("AF7:AF" & UltimaFila).NumberFormat = "General"
ActiveSheet.Range("AF7").Select
ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-25],Total!C[-31]:C[42],57,0)"
Selection.AutoFill Destination:=Range("AF7:AF" & UltimaFila)
ActiveSheet.Range(Selection, Selection.End(xlDown)).Select

Sheets("JobInfoImportTemplate").Range("AG7:AG" & UltimaFila).NumberFormat = "General"
ActiveSheet.Range("AG7").Select
ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-26],Total!C[-32]:C[41],42,0)"
Selection.AutoFill Destination:=Range("AG7:AG" & UltimaFila)
ActiveSheet.Range(Selection, Selection.End(xlDown)).Select

Sheets("JobInfoImportTemplate").Range("AH7:AH" & UltimaFila).NumberFormat = "General"
ActiveSheet.Range("AH7").Select
ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-27],'Positions Report'!C[-33]:C[11],44,0)"
Selection.AutoFill Destination:=Range("AH7:AH" & UltimaFila)
ActiveSheet.Range(Selection, Selection.End(xlDown)).Select

ActiveSheet.Range("AJ7:AJ" & UltimaFila).Value = "1"

Sheets("JobInfoImportTemplate").Range("AL7:AL" & UltimaFila).NumberFormat = "General"
ActiveSheet.Range("AL7").Select
ActiveCell.FormulaR1C1 = "=VLOOKUP(VLOOKUP(RC[-31],Total!C[-37]:C[36],35,0),'De-Para'!C[-25]:C[-24],2,0)"
Selection.AutoFill Destination:=Range("AL7:AL" & UltimaFila)
ActiveSheet.Range(Selection, Selection.End(xlDown)).Select

Sheets("JobInfoImportTemplate").Range("AM7:AM" & UltimaFila).NumberFormat = "General"
ActiveSheet.Range("AM7").Select
ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-32],Total!C[-38]:C[35],40,0)"
Selection.AutoFill Destination:=Range("AM7:AM" & UltimaFila)
ActiveSheet.Range(Selection, Selection.End(xlDown)).Select

Sheets("JobInfoImportTemplate").Range("AN7:AN" & UltimaFila).NumberFormat = "General"
ActiveSheet.Range("AN7").Select
ActiveCell.FormulaR1C1 = "=VLOOKUP(MID(VLOOKUP(RC[-33],Total!C[-39]:C[-23],17,0),FIND(""@"",VLOOKUP(RC[-33],Total!C[-39]:C[-23],17,0))+1,LEN(VLOOKUP(RC[-33],Total!C[-39]:C[-23],17,0))-FIND(""@"",VLOOKUP(RC[-33],Total!C[-39]:C[-23],17,0))),'De-Para'!C[-18]:C[-17],2,0)"
Selection.AutoFill Destination:=Range("AN7:AN" & UltimaFila)
ActiveSheet.Range(Selection, Selection.End(xlDown)).Select

Sheets("JobInfoImportTemplate").Range("AO7:AO" & UltimaFila).NumberFormat = "General"
ActiveSheet.Range("AO7").Select
ActiveCell.FormulaR1C1 = "=VLOOKUP(VLOOKUP(RC[-34],Total!C[-40]:C[33],65,0),'De-Para'!C[-34]:C[-33],2,0)"
Selection.AutoFill Destination:=Range("AO7:AO" & UltimaFila)
ActiveSheet.Range(Selection, Selection.End(xlDown)).Select

Sheets("JobInfoImportTemplate").Range("AP7:AP" & UltimaFila).NumberFormat = "General"
ActiveSheet.Range("AP7").Select
ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-35],Total!C[-41]:C[32],68,0)"
Selection.AutoFill Destination:=Range("AP7:AP" & UltimaFila)
ActiveSheet.Range(Selection, Selection.End(xlDown)).Select

Sheets("JobInfoImportTemplate").Range("AQ7:AQ" & UltimaFila).NumberFormat = "General"
ActiveSheet.Range("AQ7").Select
ActiveCell.FormulaR1C1 = "=VLOOKUP(VLOOKUP(RC[-36],Total!C[-42]:C[31],70,0),'De-Para'!C[-27]:C[-26],2,0)"
Selection.AutoFill Destination:=Range("AQ7:AQ" & UltimaFila)
ActiveSheet.Range(Selection, Selection.End(xlDown)).Select

Sheets("JobInfoImportTemplate").Range("AR7:AR" & UltimaFila).NumberFormat = "mm/dd/yyyy"
ActiveSheet.Range("AR7").Select
ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-37],Total!C[-43]:C[30],31,0)"
Selection.AutoFill Destination:=Range("AR7:AR" & UltimaFila)
ActiveSheet.Range(Selection, Selection.End(xlDown)).Select

'Copia y pega como valores
'Sheets("JobInfoImportTemplate").Range("A7:AU" & UltimaFila).Copy
'Sheets("JobInfoImportTemplate").Range("A7:AU" & UltimaFila).PasteSpecial xlPasteValues
'Application.CutCopyMode = False


End Sub

Sub Guardar()

'Mostrar hojas ocultas
Sheets("HR core values  new").Visible = True
Sheets("Business Unit List ").Visible = True
Sheets("Probation Status Picklist").Visible = True
Sheets("Time Zones").Visible = True
Sheets("Home,Host Designation Picklist").Visible = True
Sheets("Contract Type (NA region) ").Visible = True
Sheets("Location Group list ").Visible = True

Dim Ruta As String
Dim Nombre As String
Dim Fecha As String
Ruta = ThisWorkbook.Path 'Donde se encuentra ubicado el archivo de la macro
Fecha = Format(Now, "ddMMyyyy")
Nombre = "6. JobInfoImportTemplate_holcimgrouD" & "_" & Fecha


Worksheets(Array("JobInfoImportTemplate", "Positions Report", "Total", "De-Para", "HR core values  new", "Business Unit List ", "Probation Status Picklist", "Time Zones", "Home,Host Designation Picklist", "Contract Type (NA region) ", "Location Group list ")).Copy
With ActiveWorkbook
            Sheets("HR core values  new").Visible = False
            Sheets("Business Unit List ").Visible = False
            Sheets("Probation Status Picklist").Visible = False
            Sheets("Time Zones").Visible = False
            Sheets("Home,Host Designation Picklist").Visible = False
            Sheets("Contract Type (NA region) ").Visible = False
            Sheets("Location Group list ").Visible = False
        .SaveAs Filename:=Ruta & "/" & Nombre & ".xlsx", FileFormat:=xlOpenXMLWorkbook
        .Close SaveChanges:=True
End With

'Volver a ocultar Hojas
ThisWorkbook.Activate
            Sheets("HR core values  new").Visible = False
            Sheets("Business Unit List ").Visible = False
            Sheets("Probation Status Picklist").Visible = False
            Sheets("Time Zones").Visible = False
            Sheets("Home,Host Designation Picklist").Visible = False
            Sheets("Contract Type (NA region) ").Visible = False
            Sheets("Location Group list ").Visible = False

Workbooks.Open (Ruta & "/" & Nombre & ".xlsx")
MsgBox "El archivo se ha guardado en la siguiente ruta:" & vbCrLf & Ruta, vbInformation, "Archivo Guardado"
ThisWorkbook.Close SaveChanges:=False

End Sub

Sub EliminarBoton()

Sheets("JobInfoImportTemplate").Activate
Dim NombreBoton As String
NombreBoton = "Ejecutar"
On Error Resume Next
Dim Boton As Object
Set Boton = ActiveSheet.Buttons(NombreBoton)
On Error GoTo 0
Boton.Delete

End Sub

