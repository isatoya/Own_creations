Attribute VB_Name = "Template"

Sub Template()

If IsEmpty(ThisWorkbook.Sheets("Position").Range("C6").Value) Then
MsgBox "Es necesario completar la informaci�n de los c�digos en la hoja 'Position'.", vbExclamation, "Falta Informaci�n"
Else

If IsEmpty(ThisWorkbook.Sheets("Positions Report").Range("A3").Value) Or IsEmpty(ThisWorkbook.Sheets("Positions Report").Range("A4").Value) Then
MsgBox "Es necesario pegar los datos en la hoja �Positions Report�.", vbExclamation, "Falta Informaci�n"
Else

If IsEmpty(ThisWorkbook.Sheets("Total").Range("A3").Value) Or IsEmpty(ThisWorkbook.Sheets("Total").Range("A4").Value) Then
MsgBox "Es necesario pegar los datos en la hoja �Total�.", vbExclamation, "Falta Informaci�n"
Else


' Ambas condiciones se cumplieron, llamar al sub "BuscarV"
Cambio_colum
Buscarv
Borrar_boton
Guardar

End If
End If
End If

End Sub

Sub Cambio_colum()

'Cambiar las columnas de la hoja de positions reports y Total de orden

Worksheets("Positions Report").Activate
Columns("A:A").Select
Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
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
Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
Columns("AE:AE").Select
Selection.Cut
Columns("A:A").Select
ActiveSheet.Paste
Columns("AE:AE").Select
Selection.Delete Shift:=xlToLeft
Columns("A:A").Select
Selection.NumberFormat = "General"

End Sub

Sub Buscarv()

'Encuentra la �ltima fila en la columna C

Dim UltimaFila As Long
Sheets("Position").Activate
UltimaFila = Sheets("Position").Cells(Sheets("Position").Rows.Count, "C").End(xlUp).Row

'Buscarv

Sheets("Position").Range("D6:D" & UltimaFila).NumberFormat = "General"
ActiveSheet.Range("D6").Select
ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-1],'Positions Report'!C[-3]:C[41],29,0)"
Selection.AutoFill Destination:=Range("D6:D" & UltimaFila)
ActiveSheet.Range(Selection, Selection.End(xlDown)).Select

ActiveSheet.Range("Y6:Y" & UltimaFila).Value = "A"

Sheets("Position").Range("Z6:Z" & UltimaFila).NumberFormat = "mm/dd/yyyy"
ActiveSheet.Range("Z6").Select
ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-23], 'Positions Report'!C[-25]:C[19], 28, FALSE)"
Selection.AutoFill Destination:=Range("Z6:Z" & UltimaFila)
ActiveSheet.Range(Selection, Selection.End(xlDown)).Select

ActiveSheet.Range("AA6:AA" & UltimaFila).Value = "RP"

ActiveSheet.Range("AB6:AB" & UltimaFila).Value = "updatePosition"

Sheets("Position").Range("AC6:AC" & UltimaFila).NumberFormat = "General"
ActiveSheet.Range("AC6").Select
ActiveCell.FormulaR1C1 = "=VLOOKUP(VLOOKUP(RC[-26],Total!C[-28]:C[45],63,0),'De-Para'!C[-28]:C[-27],2,0)"
Selection.AutoFill Destination:=Range("AC6:AC" & UltimaFila)
ActiveSheet.Range(Selection, Selection.End(xlDown)).Select

Sheets("Position").Range("AD6:AD" & UltimaFila).NumberFormat = "General"
ActiveSheet.Range("AD6").Select
ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-27],'Positions Report'!C[-29]:C[15],11,0)"
Selection.AutoFill Destination:=Range("AD6:AD" & UltimaFila)
ActiveSheet.Range(Selection, Selection.End(xlDown)).Select

Sheets("Position").Range("AE6:AE" & UltimaFila).NumberFormat = "General"
ActiveSheet.Range("AE6").Select
ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-28],'Positions Report'!C[-30]:C[14],10,0)"
Selection.AutoFill Destination:=Range("AE6:AE" & UltimaFila)
ActiveSheet.Range(Selection, Selection.End(xlDown)).Select

Sheets("Position").Range("AF6:AF" & UltimaFila).NumberFormat = "General"
ActiveSheet.Range("AF6").Select
ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-29],'Positions Report'!C[-31]:C[13],12,0)"
Selection.AutoFill Destination:=Range("AF6:AF" & UltimaFila)
ActiveSheet.Range(Selection, Selection.End(xlDown)).Select

Sheets("Position").Range("AG6:AG" & UltimaFila).NumberFormat = "General"
ActiveSheet.Range("AG6").Select
ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-30],'Positions Report'!C[-32]:C[12],24,0)"
Selection.AutoFill Destination:=Range("AG6:AG" & UltimaFila)
ActiveSheet.Range(Selection, Selection.End(xlDown)).Select

ActiveSheet.Range("AH6:AH" & UltimaFila).Value = "1"

Sheets("Position").Range("AI6:AI" & UltimaFila).NumberFormat = "General"
ActiveSheet.Range("AI6").Select
ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-32],'Positions Report'!C[-34]:C[10],35,0)"
Selection.AutoFill Destination:=Range("AI6:AI" & UltimaFila)
ActiveSheet.Range(Selection, Selection.End(xlDown)).Select

Sheets("Position").Range("AJ6:AJ" & UltimaFila).NumberFormat = "General"
ActiveSheet.Range("AJ6").Select
ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-33],Total!C[-35]:C[38],33,0)"
Selection.AutoFill Destination:=Range("AJ6:AJ" & UltimaFila)
ActiveSheet.Range(Selection, Selection.End(xlDown)).Select

Sheets("Position").Range("AK6:AK" & UltimaFila).NumberFormat = "General"
ActiveSheet.Range("AK6").Select
ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-34],'Positions Report'!C[-36]:C[8],2,0)"
Selection.AutoFill Destination:=Range("AK6:AK" & UltimaFila)
ActiveSheet.Range(Selection, Selection.End(xlDown)).Select

Sheets("Position").Range("AL6:AL" & UltimaFila).NumberFormat = "General"
ActiveSheet.Range("AL6").Select
ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-35],'Positions Report'!C[-37]:C[7],8,0)"
Selection.AutoFill Destination:=Range("AL6:AL" & UltimaFila)
ActiveSheet.Range(Selection, Selection.End(xlDown)).Select

Sheets("Position").Range("AM6:AM" & UltimaFila).NumberFormat = "General"
ActiveSheet.Range("AM6").Select
ActiveCell.FormulaR1C1 = "=VLOOKUP(VLOOKUP(RC[-36],Total!C[-38]:C[35],67,0),'De-Para'!C[-35]:C[-34],2,0)"
Selection.AutoFill Destination:=Range("AM6:AM" & UltimaFila)
ActiveSheet.Range(Selection, Selection.End(xlDown)).Select

Sheets("Position").Range("AN6:AN" & UltimaFila).NumberFormat = "General"
ActiveSheet.Range("AN6").Select
ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-37],'Positions Report'!C[-39]:C[34],6,0)"
Selection.AutoFill Destination:=Range("AN6:AN" & UltimaFila)
ActiveSheet.Range(Selection, Selection.End(xlDown)).Select

Sheets("Position").Range("AO6:AO" & UltimaFila).NumberFormat = "General"
ActiveSheet.Range("AO6").Select
ActiveCell.FormulaR1C1 = "=MID(CONCATENATE(""L"",VLOOKUP(RC[-38],'Positions Report'!C[-40]:C[4],25,0)),1,7)"
Selection.AutoFill Destination:=Range("AO6:AO" & UltimaFila)
ActiveSheet.Range(Selection, Selection.End(xlDown)).Select

Sheets("Position").Range("AP6:AP" & UltimaFila).NumberFormat = "General"
ActiveSheet.Range("AP6").Select
ActiveCell.FormulaR1C1 = "=CONCATENATE(""LA_"",VLOOKUP(RC[-39],'Positions Report'!C[-41]:C[32],2,0),""_Cluster"")"
Selection.AutoFill Destination:=Range("AP6:AP" & UltimaFila)
ActiveSheet.Range(Selection, Selection.End(xlDown)).Select

Sheets("Position").Range("AQ6:AQ" & UltimaFila).NumberFormat = "General"
ActiveSheet.Range("AQ6").Select
ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-40],'Positions Report'!C[-42]:C[31],25,0)"
Selection.AutoFill Destination:=Range("AQ6:AQ" & UltimaFila)
ActiveSheet.Range(Selection, Selection.End(xlDown)).Select

Sheets("Position").Range("AR6:AR" & UltimaFila).NumberFormat = "General"
ActiveSheet.Range("AR6").Select
ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-41],'Positions Report'!C[-43]:C[30],20,0)"
Selection.AutoFill Destination:=Range("AR6:AR" & UltimaFila)
ActiveSheet.Range(Selection, Selection.End(xlDown)).Select

Sheets("Position").Range("AU6:AU" & UltimaFila).NumberFormat = "General"
ActiveSheet.Range("AU6").Select
ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-44],'Positions Report'!C[-46]:C[27],27,0)"
Selection.AutoFill Destination:=Range("AU6:AU" & UltimaFila)
ActiveSheet.Range(Selection, Selection.End(xlDown)).Select

Sheets("Position").Range("AV6:AV" & UltimaFila).NumberFormat = "General"
ActiveSheet.Range("AV6").Select
ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-45],'Positions Report'!C[-47]:C[26],31,0)"
Selection.AutoFill Destination:=Range("AV6:AV" & UltimaFila)
ActiveSheet.Range(Selection, Selection.End(xlDown)).Select

Sheets("Position").Range("AW6:AW" & UltimaFila).NumberFormat = "General"
ActiveSheet.Range("AW6").Select
ActiveCell.FormulaR1C1 = "=CONCATENATE(VLOOKUP(RC[-46],'Positions Report'!C[-48]:C[-4],24,0),""-"",VLOOKUP(RC[-46],'Positions Report'!C[-48]:C[25],2,0),""-"",VLOOKUP(RC[-46],Total!C[-48]:C[25],33,0))"
Selection.AutoFill Destination:=Range("AW6:AW" & UltimaFila)
ActiveSheet.Range(Selection, Selection.End(xlDown)).Select

Sheets("Position").Range("BB6:BB" & UltimaFila).NumberFormat = "General"
ActiveSheet.Range("BB6").Select
ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-51],'Positions Report'!C[-53]:C[20],19,0)"
Selection.AutoFill Destination:=Range("BB6:BB" & UltimaFila)
ActiveSheet.Range(Selection, Selection.End(xlDown)).Select

Sheets("Position").Range("BC6:BC" & UltimaFila).NumberFormat = "General"
ActiveSheet.Range("D6:D7").NumberFormat = "General"
ActiveSheet.Range("BC6").Select
ActiveCell.FormulaR1C1 = "=VLOOKUP(VLOOKUP(RC[-52],Total!C[-54]:C[19],65,0),'De-Para'!C[-48]:C[-47],2,0)"
Selection.AutoFill Destination:=Range("BC6:BC" & UltimaFila)
ActiveSheet.Range(Selection, Selection.End(xlDown)).Select

Sheets("Position").Range("BD6:BD" & UltimaFila).NumberFormat = "General"
ActiveSheet.Range("BD6").Select
ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-53],Total!C[-55]:C[18],68,0)"
Selection.AutoFill Destination:=Range("BD6:BD" & UltimaFila)
ActiveSheet.Range(Selection, Selection.End(xlDown)).Select

Sheets("Position").Range("BE6:BE" & UltimaFila).NumberFormat = "General"
ActiveSheet.Range("BE6").Select
ActiveCell.FormulaR1C1 = "=VLOOKUP(VLOOKUP(RC[-54],Total!C[-56]:C[13],70,0),'De-Para'!C[-47]:C[-46],2,0)"
Selection.AutoFill Destination:=Range("BE6:BE" & UltimaFila)
ActiveSheet.Range(Selection, Selection.End(xlDown)).Select

Sheets("Position").Range("BF6:BF" & UltimaFila).NumberFormat = "General"
ActiveSheet.Range("BF6").Select
ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-55],'Positions Report'!C[-57]:C[-13],40,0)"
Selection.AutoFill Destination:=Range("BF6:BF" & UltimaFila)
ActiveSheet.Range(Selection, Selection.End(xlDown)).Select

Sheets("Position").Range("BH6:BH" & UltimaFila).NumberFormat = "General"
ActiveSheet.Range("BH6").Select
ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-57],Total!C[-59]:C[14],59,0)"
Selection.AutoFill Destination:=Range("BH6:BH" & UltimaFila)
ActiveSheet.Range(Selection, Selection.End(xlDown)).Select

'Copia y pega como valores
'Sheets("Position").Range("A6:BI" & UltimaFila).Copy
'Sheets("Position").Range("A6:BI" & UltimaFila).PasteSpecial xlPasteValues
'Application.CutCopyMode = False

End Sub

Sub Guardar()

'Mostrar hojas ocultas
Sheets("HR core values  picklists").Visible = True
Sheets("Business Unit List").Visible = True
Sheets("Location Group List").Visible = True

Dim Ruta As String
Dim Nombre As String
Dim Fecha As String
Ruta = ThisWorkbook.Path 'Donde se encuentra ubicado el archivo de la macro
Fecha = Format(Now, "ddMMyyyy")
Nombre = "5.Position" & "_" & Fecha


Worksheets(Array("Position", "Positions Report", "Total", "De-Para", "HR core values  picklists", "Business Unit List", "Location Group List")).Copy
With ActiveWorkbook
        Sheets("HR core values  picklists").Visible = False
        Sheets("Business Unit List").Visible = False
        Sheets("Location Group List").Visible = False
        .SaveAs Filename:=Ruta & "/" & Nombre & ".xlsx", FileFormat:=xlOpenXMLWorkbook
        .Close SaveChanges:=False
End With

'Volver a ocultar Hojas
ThisWorkbook.Activate
Sheets("HR core values  picklists").Visible = False
Sheets("Business Unit List").Visible = False
Sheets("Location Group List").Visible = False


Workbooks.Open (Ruta & "/" & Nombre & ".xlsx")
MsgBox "El archivo se ha guardado en la siguiente ruta:" & vbCrLf & Ruta, vbInformation, "Archivo Guardado"
ThisWorkbook.Close SaveChanges:=False

End Sub

Sub Borrar_boton()

Sheets("Position").Activate
Dim NombreBoton As String
NombreBoton = "Ejecutar"
On Error Resume Next
Dim Boton As Object
Set Boton = ActiveSheet.Buttons(NombreBoton)
On Error GoTo 0
Boton.Delete

End Sub
