Attribute VB_Name = "M�dulo1"
Sub AbrirFormulario()
    UserForm1.Show ' Abre el formulario cuando se le da clic a la forma
End Sub


Private Sub UserForm_Initialize()

    ComboBox1.AddItem "Altas"
    ComboBox1.AddItem "Ausentismos"
    ComboBox1.AddItem "Bajas"
    ComboBox1.AddItem "Cambios organizacionales (Cambio Jefe)"
    ComboBox1.AddItem "Cambios organizacionales(Puesto)"
    ComboBox1.AddItem "Contingente de Vacaciones 2.0"
    ComboBox1.AddItem "Contratos Y Finiquitos"
'    ComboBox1.AddItem "Datos Familiares"
    ComboBox1.AddItem "Horas Extras"
    ComboBox1.AddItem "Incrementos Salariales"
    ComboBox1.AddItem "Informe de p�lizas"
'    ComboBox1.AddItem "Maestro de Estructura"
'    ComboBox1.AddItem "Maestro de Personal"
'    ComboBox1.AddItem "Mestro IT"

End Sub


Sub CommandButton1_Click()

'La macro se encarga de acumular los reportes que el equipo tenga en las carpetas de cada
'mes por medio del formulario. El reporte que se escoja en el combobox es el que la macro
'va a buscar en cada carpeta de cada mes seleccionado del checkbox y los va a unir en un
'documento nuevo xlsx, el cual va a ubicar en una carpeta llamada 'Reportes acumulados'.
'Adem�s, la macro elimina los valores duplicados de ese nuevo archivo si la persona lo desea.

'NOTA: La macro debe guardarse en la misma carpeta donde se encuentran las carpetas de cada mes; de lo contrario, no funcionar� correctamente
'FECHA: Noviembre 2023
'DESARROLLADOR: Isabel Montoya


    Application.ScreenUpdating = False  'Desactivar activacion de cambio de libro
    
    'Variables
    Dim selectedReport As String  'nombre el reporte
    Dim selectedMonths() As String 'meses
    Dim i As Integer

    selectedReport = UserForm1.ComboBox1.Value 'Reporte seleccionado en el combobox


    'Veificar meses seleccionados en el checkbox --> 1=enero, 2=febrero, 3=marzo...
    
    For i = 1 To 12
        If UserForm1.Controls("CheckBox" & i).Value = True Then
            ReDim Preserve selectedMonths(1 To i)
            selectedMonths(i) = UserForm1.Controls("CheckBox" & i).Caption
        End If
    Next i


    'Ruta
    Dim basePath As String
    basePath = ThisWorkbook.Path & "\" 'donde se encuentre este archivo de la macro


    'Crear la carpeta, primero verifica si ya existe
    If Dir(basePath & "Reportes acumulados", vbDirectory) = "" Then
        MkDir basePath & "Reportes acumulados"
    End If


    'Crear libro nuevo en excel
    Dim newWorkbook As Workbook
    Set newWorkbook = Workbooks.Add


    ' Nombre del nuevo archivo
    Dim newFileName As String
    newFileName = selectedReport & " - " & Format(Now, "yyyy-mm-dd")


    'Formato celda A2 para colocar los meses
    newWorkbook.Sheets(1).Range("A2").Value = "Mes del reporte"
    With newWorkbook.Sheets(1).Range("A2").Font
        .Bold = True
        .Color = RGB(255, 255, 255)
    End With
    With newWorkbook.Sheets(1).Range("A2").Interior
        .Color = RGB(255, 0, 0)
    End With

    'Nombre del archivo y guardar
    newWorkbook.SaveAs basePath & "Reportes acumulados\" & newFileName & ".xlsx"



    'Proceso para buscar los reportes dentro de las carpetas
    
        For i = LBound(selectedMonths) To UBound(selectedMonths)
            
            Dim monthFolder As String 'variable
            monthFolder = basePath & selectedMonths(i) & "\"  'Construir la ruta de la carpeta del mes
    
            'Buscar la carpeta por el nombre del mes
            If Dir(monthFolder, vbDirectory) <> "" Then
                
                Dim fileName As String 'Busca el reporte para construir la ruta
                fileName = Dir(monthFolder & "*" & selectedReport & "*.xlsx")
    
                'Buscar el reporte dentro de cada carpeta
                Do While fileName <> ""
                    
                    Dim fullpath As String 'Construir ruta completa
                    fullpath = monthFolder & fileName
    
                    Workbooks.Open fullpath 'Abrir reporte
         
                    Dim lastRowA As Long 'Ultima fila de la columna A de la hoja del archivo en blanco
                    lastRowA = newWorkbook.Sheets(1).Cells(newWorkbook.Sheets(1).Rows.Count, "A").End(xlUp).row
    
                    'Copiar y pegar datos
                    ActiveSheet.UsedRange.Copy newWorkbook.Sheets(1).Cells(newWorkbook.Sheets(1).Rows.Count, "B").End(xlUp).Offset(1, 0)
    
                    'Rellenar la columna A con el nombre del mes
                    newWorkbook.Sheets(1).Range("A" & lastRowA + 1 & ":A" & newWorkbook.Sheets(1).Cells(newWorkbook.Sheets(1).Rows.Count, "B").End(xlUp).row).Value = selectedMonths(i)
    
                    ActiveWorkbook.Close SaveChanges:=False 'Cerrar archivo del reporte despues de pegar los datos
    
                    'Buscar el siguiente reporte
                    fileName = Dir
                    
                Loop
            Else
                MsgBox "No se encontro la carpeta de: " & monthFolder, vbExclamation
            End If
        Next i

    
    newWorkbook.Save 'Guardar
    
    
    
    'Eliminar filas de color (menos la fila 2)
        Dim lastRow_color As Long
        lastRow_color = newWorkbook.Sheets(1).Cells(newWorkbook.Sheets(1).Rows.Count, "B").End(xlUp).row
        
        Dim x As Long
        For x = lastRow_color To 3 Step -1 ' Recorremos desde la �ltima fila hacia arriba
            If newWorkbook.Sheets(1).Cells(x, 2).Interior.Color <> RGB(255, 255, 255) Then  'verifica el color de fondo
                newWorkbook.Sheets(1).Rows(x).Delete
            End If
        Next x
        
        newWorkbook.Save 'Guardar

    
    'Convertir a numero lo de las advertencias
    Dim celda_num As Range
    Dim hoja_num As Worksheet
    Dim rango_num As Range

    ' Definir la hoja de trabajo en la que deseas buscar
    Set hoja_num = newWorkbook.Sheets(1)

    ' Buscar todas las celdas con advertencia en la hoja
    On Error Resume Next
    Set rango_num = hoja_num.Cells.SpecialCells(xlCellTypeConstants, xlTextValues)
    On Error GoTo 0

    ' Si se encuentra un rango con advertencias, convertir a formato n�mero
    If Not rango_num Is Nothing Then
        For Each celda_num In rango_num
            ' Verificar si la celda no est� vac�a
            If Not IsEmpty(celda_num.Value) Then
                ' Convertir el valor de la celda a n�mero
                If IsNumeric(celda_num.Value) Then
                    celda_num.Value = CDbl(celda_num.Value)
                End If
            End If
        Next celda_num
    Else
        MsgBox "No se encontraron celdas con advertencias en formato de texto.", vbInformation
    End If
    
    

    
    'Preguntar si de deben de quitar los dopliados
    
        Dim respuesta As VbMsgBoxResult
        
        ' Pregunta al usuario si desea eliminar duplicados
        respuesta = MsgBox("�Desea eliminar filas duplicadas?", vbYesNo + vbQuestion, "Confirmaci�n")
        
        ' Verifica la respuesta del usuario
        If respuesta = vbYes Then
        
            'Inicia el proceso de la eliminacion de duplicados
            
                        Dim lastRow_duplicado As Long
                        lastRow_duplicado = ActiveSheet.Cells(ActiveSheet.Rows.Count, "B").End(xlUp).row
                        
                        Dim rngData As Range
                        Set rngData = ActiveSheet.Range("B3", ActiveSheet.Cells(lastRow_duplicado, ActiveSheet.Columns.Count).End(xlToLeft))
                        
                        ' Identificar las filas duplicadas
                        Dim dict As Object
                        Set dict = CreateObject("Scripting.Dictionary")
                        
                        Dim row As Range
                        For Each row In rngData.Rows
                            Dim keyString As String
                            keyString = Join(Application.Transpose(Application.Transpose(row.Value)), "|") ' Concatenar los valores de la fila
                            
                            On Error Resume Next
                            dict.Add keyString, row
                            On Error GoTo 0
                        Next row
                        
                        'Las que encuentre duplicadas las pone de color morado
                        Dim dictKey As Variant
                        For Each dictKey In dict.Keys
                            dict(dictKey).Interior.Color = RGB(216, 191, 216)
                        Next dictKey
                    
                    
                        'eliminar las filas en blanco
                        Dim hoja As Worksheet
                        Dim ultimaFila As Long
                        Dim fila As Long
                        Dim celda As Range
                        
                        Set hoja = newWorkbook.Sheets(1) 'especifica la hoja en la que se van a hacer los cambios
                        ultimaFila = hoja.Cells(hoja.Rows.Count, "B").End(xlUp).row 'Encuentra la �ltima fila con datos en la columna B
                         
                        ' Itera desde la �ltima fila hasta la fila 3
                        For fila = ultimaFila To 3 Step -1
                            If hoja.Cells(fila, 2).Interior.ColorIndex = xlNone Then ' Verifica si la celda en la columna B no tiene color de fondo (morado)
                                hoja.Rows(fila).Delete ' Elimina la fila si no tiene color de fondo (morado)
                            End If
                        Next fila
                    
                        'Le pone de nuevo el formato para que quede todo sin color morado (para que el archivo quede normal)
                        For Each celda In hoja.UsedRange
                            If celda.Interior.Color = RGB(216, 191, 216) Then
                                celda.Interior.ColorIndex = xlNone
                            End If
                        Next celda

        Else

        End If


    'Guardar y cerrar
    newWorkbook.Save
    newWorkbook.Close SaveChanges:=False
    
    Application.ScreenUpdating = True
    
    ' y mostrar un mensaje indicando que el reporte est� listo
    MsgBox "El reporte est� listo y se ha guardado en:" & vbNewLine & basePath & "Reportes acumulados\" & newFileName & ".xlsx"
    
    UserForm1.Hide


End Sub
