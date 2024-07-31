Attribute VB_Name = "AUT_ENEC_2024_02"
' Variables globales
Global Grafica_vertical
Global Grafica_Sect
Global Grafica_vertical_nacional
Global Grafica_vertical_nacional_men
Global Grafica_horizontal
Global Grafica_his
Global Grafica_Obra
Global Grafica_var


Sub Macro_inicio()
Attribute Macro_inicio.VB_ProcData.VB_Invoke_Func = "a\n14"
'Ctrl a
Call WorksheetLoop
Call Generación_Word

End Sub

Sub WorksheetLoop()

        Dim WS_Count As Integer
        Dim I As Integer

         ' Define WS_Count como el número de hojas en el libro activo
        WS_Count = ActiveWorkbook.Worksheets.Count

         ' Empieza el loop
        For I = 1 To WS_Count
            
            Worksheets(I).Select
            nombre = Worksheets(I).Name
            
            If InStr(1, nombre, "HIS", vbBinaryCompare) = 1 Then
                Call Macro_linea_promedio
            ElseIf InStr(1, nombre, "RANKDIS", vbBinaryCompare) = 1 Then
                Call Macro_Graficas
            ElseIf InStr(1, nombre, "RANKVPP", vbBinaryCompare) = 1 Then
                Call Macro_Graficas_Nacional
            ElseIf InStr(1, nombre, "RANKVMEN", vbBinaryCompare) = 1 Then
                Call Macro_Graficas_Nacional_Mensual
            ElseIf InStr(1, nombre, "MES", vbBinaryCompare) = 1 Then
                Call Macro_grafica_horizontal
            ElseIf InStr(1, nombre, "SECTOR", vbBinaryCompare) = 1 Then
                Call Macro_Sect
            ElseIf InStr(1, nombre, "OBRA", vbBinaryCompare) = 1 Then
                Call Macro_Obra
            End If

        Next I

End Sub


Sub Macro_Graficas()
'RANK
' Ranking por entidad federativa
' Ctrl + g

Dim Grafica As ChartObject

fila = 4

Do While Cells(fila, 1) = ""
    fila = fila + 1
Loop

ultimo = fila
Do While Cells(ultimo, 1) <> ""
    ultimo = ultimo + 1
Loop
'

'
Range("B" & (fila + 1) & ":B" & (ultimo - 1)).NumberFormat = "0.0"
hoja = ActiveSheet.Name

'
Range("A" & fila & ":B" & (ultimo - 1)).AutoFilter
'

Range("A" & (fila + 1) & ":B" & (ultimo - 1)).Sort Key1:=Range("B" & (fila + 1)), Order1:=xlAscending

'
jalisco = fila

Do While Cells(jalisco, 1) <> "Jalisco"
    jalisco = jalisco + 1
Loop
jalisco = jalisco - fila
'
Range("A" & (fila + 1) & ":B" & (ultimo - 1)).Select
'
Set Grafica_vertical = ActiveSheet.ChartObjects.Add(Left:=4 * 48, Width:=468.1, Top:=60, Height:=448.5)

With Grafica_vertical.Chart
    .ApplyChartTemplate ("C:\Users\arturo.carrillo\AppData\Roaming\Microsoft\Plantillas\Charts\ENEC RANK AUT.crtx") ' UBICACIÓN PERSONAL
    .SetSourceData Source:=Range("A" & (fila + 1) & ":B" & (ultimo - 1))
    .SeriesCollection(1).Points(jalisco).Format.Fill.ForeColor.RGB = RGB(251, 187, 39)
End With
'
Range("A" & fila & ":B" & (ultimo - 1)).AutoFilter
'
End Sub

Sub Macro_Graficas_Nacional()
'RANK
' Ranking por entidad federativa con valor nacional, anual
'

Dim Grafica As ChartObject

fila = 4

Do While Cells(fila, 1) = ""
    fila = fila + 1
Loop

ultimo = fila
Do While Cells(ultimo, 1) <> ""
    ultimo = ultimo + 1
Loop
'

'
Range("B" & (fila + 1) & ":B" & (ultimo - 1)).NumberFormat = "0.0"
hoja = ActiveSheet.Name

'
Range("A" & fila & ":B" & (ultimo - 1)).AutoFilter
'

Range("A" & (fila + 1) & ":B" & (ultimo - 1)).Sort Key1:=Range("B" & (fila + 1)), Order1:=xlAscending

'
jalisco = fila
nacional = fila

Do While Cells(jalisco, 1) <> "Jalisco"
    jalisco = jalisco + 1
Loop
jalisco = jalisco - fila

Do While Cells(nacional, 1) <> "Nacional"
    nacional = nacional + 1
Loop
nacional = nacional - fila
'
Range("A" & (fila + 1) & ":B" & (ultimo - 1)).Select
'
Set Grafica_vertical_nacional = ActiveSheet.ChartObjects.Add(Left:=4 * 48, Width:=468.1, Top:=60, Height:=448.5)

With Grafica_vertical_nacional.Chart
    .ApplyChartTemplate ("C:\Users\arturo.carrillo\AppData\Roaming\Microsoft\Plantillas\Charts\ENEC RANK AUT.crtx") ' UBICACIÓN PERSONAL
    .SetSourceData Source:=Range("A" & (fila + 1) & ":B" & (ultimo - 1))
    .SeriesCollection(1).Points(jalisco).Format.Fill.ForeColor.RGB = RGB(251, 187, 39)
    .SeriesCollection(1).Points(nacional).Format.Fill.ForeColor.RGB = RGB(149, 104, 43)
End With
'
Range("A" & fila & ":B" & (ultimo - 1)).AutoFilter
'
End Sub

Sub Macro_Graficas_Nacional_Mensual()
'RANK
' Ranking por entidad federativa con valor nacional, mensual
'

Dim Grafica As ChartObject

fila = 4

Do While Cells(fila, 1) = ""
    fila = fila + 1
Loop

ultimo = fila
Do While Cells(ultimo, 1) <> ""
    ultimo = ultimo + 1
Loop
'

'
Range("B" & (fila + 1) & ":B" & (ultimo - 1)).NumberFormat = "0.0"
hoja = ActiveSheet.Name

'
Range("A" & fila & ":B" & (ultimo - 1)).AutoFilter
'

Range("A" & (fila + 1) & ":B" & (ultimo - 1)).Sort Key1:=Range("B" & (fila + 1)), Order1:=xlAscending

'
jalisco = fila
nacional = fila

Do While Cells(jalisco, 1) <> "Jalisco"
    jalisco = jalisco + 1
Loop
jalisco = jalisco - fila

Do While Cells(nacional, 1) <> "Nacional"
    nacional = nacional + 1
Loop
nacional = nacional - fila
'
Range("A" & (fila + 1) & ":B" & (ultimo - 1)).Select
'
Set Grafica_vertical_nacional_men = ActiveSheet.ChartObjects.Add(Left:=4 * 48, Width:=468.1, Top:=60, Height:=448.5)

With Grafica_vertical_nacional_men.Chart
    .ApplyChartTemplate ("C:\Users\arturo.carrillo\AppData\Roaming\Microsoft\Plantillas\Charts\ENEC RANK AUT.crtx") ' UBICACIÓN PERSONAL
    .SetSourceData Source:=Range("A" & (fila + 1) & ":B" & (ultimo - 1))
    .SeriesCollection(1).Points(jalisco).Format.Fill.ForeColor.RGB = RGB(251, 187, 39)
    .SeriesCollection(1).Points(nacional).Format.Fill.ForeColor.RGB = RGB(149, 104, 43)
End With
'
Range("A" & fila & ":B" & (ultimo - 1)).AutoFilter
'
End Sub

Sub Macro_grafica_horizontal()
'MES
'Cifras Mismo Mes Distintos Años
'Ctrl + h

Range("B6").Select
Range(Selection, Selection.End(xlDown)).Select
Selection.NumberFormat = "#,##0"

inicio = 5
fin = inicio

Do While Cells(fin, 1) <> ""
    fin = fin + 1
Loop

Range("A" & (inicio + 1) & ":B" & (fin - 1)).Select

Set Grafica_horizontal = ActiveSheet.ChartObjects.Add(Left:=4 * 48, Width:=468.1, Top:=60, Height:=250)

With Grafica_horizontal.Chart
    .ApplyChartTemplate ("C:\Users\arturo.carrillo\AppData\Roaming\Microsoft\Plantillas\Charts\ENEC MES AUT.crtx") ' UBICACIÓN PERSONAL
    .SetSourceData Source:=Range("A" & (inicio + 1) & ":B" & (fin - 1))
    .SeriesCollection(1).Points(fin - 1 - inicio).Format.Fill.ForeColor.RGB = RGB(251, 187, 39)
    .SetElement (msoElementDataLabelOutSideEnd)
    For j = 1 To (fin - 2 - inicio)
        .SeriesCollection(1).Points(j).DataLabel.Delete
    Next j
End With

End Sub

Sub Macro_linea_promedio()
'HIS & VAR
'Gráfica de Barras de Históricos Mensuales  con Linea de Promedio de Últimos 12 Meses
'Ctrl + p

nombre = ActiveSheet.Name

If InStr(1, nombre, "HIS", vbBinaryCompare) = 1 Then
    Range("C6:D6").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.NumberFormat = "#,##0"

    inicio = 5
    fin = inicio
    
    Do While Cells(fin, 2) <> ""
        fin = fin + 1
    Loop
    
    Range("A" & (inicio) & ":D" & (fin - 1)).Select

    Set Grafica_his = ActiveSheet.ChartObjects.Add(Left:=5 * 48, Width:=468.1, Top:=60, Height:=250)
    
    With Grafica_his.Chart
        .ApplyChartTemplate ("C:\Users\arturo.carrillo\AppData\Roaming\Microsoft\Plantillas\Charts\ENEC HIS AUT.crtx") ' UBICACIÓN PERSONAL
        .SetSourceData Source:=Range("A" & (inicio) & ":D" & (fin - 1))
        For k = 1 To (fin - 1)
            If (k Mod 12) = ((fin - 1 - inicio) Mod 12) Then
                .SeriesCollection(1).Points(k).Format.Fill.ForeColor.RGB = RGB(124, 135, 142)
            End If
        Next k
        .SeriesCollection(1).Points(fin - 1 - inicio).Format.Fill.ForeColor.RGB = RGB(251, 187, 39)
    End With
End If

End Sub
Sub Macro_Sect()


inicio1 = 5
fin1 = inicio1
    
    Do While Cells(fin1, 3) <> ""
        fin1 = fin1 + 1
    Loop
    
Range("C" & (inicio1) & ":D" & (fin1 - 1)).Select
Selection.NumberFormat = "#,##0"

Range("E" & (inicio1) & ":E" & (fin1 - 1)).Select
Selection.NumberFormat = "0.0%"

Set Grafica_Sect = ActiveSheet.ChartObjects.Add(Left:=8 * 48, Width:=468.1, Top:=60, Height:=250)

With Grafica_Sect.Chart
        .ApplyChartTemplate ("C:\Users\arturo.carrillo\AppData\Roaming\Microsoft\Plantillas\Charts\ENEC SECT AUT.crtx") ' UBICACIÓN PERSONAL
        .SetSourceData Source:=Range("A" & (inicio1) & ":E" & (fin1 - 1))
End With
End Sub
Sub Macro_Obra()


inicio1 = 5
fin1 = inicio1
    
    Do While Cells(fin1, 3) <> ""
        fin1 = fin1 + 1
    Loop
    
Range("A" & (inicio1) & ":E" & (fin1 - 1)).Select
Selection.NumberFormat = "#,##0"

Set Grafica_Obra = ActiveSheet.ChartObjects.Add(Left:=7 * 48, Width:=468.1, Top:=60, Height:=250)

With Grafica_Obra.Chart
        .ApplyChartTemplate ("C:\Users\arturo.carrillo\AppData\Roaming\Microsoft\Plantillas\Charts\ENEC OBRA AUT.crtx") ' UBICACIÓN PERSONAL
        .SetSourceData Source:=Range("A" & (inicio1) & ":E" & (fin1 - 1))
End With
End Sub
Sub Generación_Word()

' Nombre  y ubicación de la plantilla
plantilla = "C:\Users\arturo.carrillo\Documents\ENEC\AUT\PLANTILLA ENEC 2024.dotx" ' UBICACIÓN PERSONAL

' Creamos el nuevo archivo word usando la plantilla
Set aplicacion = CreateObject("Word.Application")
aplicacion.Visible = True

Set documento = aplicacion.Documents.Add(Template:=plantilla, NewTemplate:=False, DocumentType:=0)

' Cambiamos la fecha del encabezado
diahoy = Format(Day(Now), "00")
meshoy = Format(Month(Now), "00")
añohoy = Year(Now)
If Month(Now) = 1 Then
    meshoypal = "enero"
    mesbas = Format(11, "00")
    mesbaspal = "noviembre"
    añobas = Year(Now) - 1
ElseIf Month(Now) = 2 Then
    meshoypal = "febrero"
    mesbas = Format(12, "00")
    mesbaspal = "diciembre"
    añobas = Year(Now) - 1
ElseIf Month(Now) = 3 Then
    meshoypal = "marzo"
    mesbas = Format(1, "00")
    mesbaspal = "enero"
    añobas = Year(Now)
ElseIf Month(Now) = 4 Then
    meshoypal = "abril"
    mesbas = Format(2, "00")
    mesbaspal = "febrero"
    añobas = Year(Now)
ElseIf Month(Now) = 5 Then
    meshoypal = "mayo"
    mesbas = Format(3, "00")
    mesbaspal = "marzo"
    añobas = Year(Now)
ElseIf Month(Now) = 6 Then
    meshoypal = "junio"
    mesbas = Format(4, "00")
    mesbaspal = "abril"
    añobas = Year(Now)
ElseIf Month(Now) = 7 Then
    meshoypal = "julio"
    mesbas = Format(5, "00")
    mesbaspal = "mayo"
    añobas = Year(Now)
ElseIf Month(Now) = 8 Then
    meshoypal = "agosto"
    mesbas = Format(6, "00")
    mesbaspal = "junio"
    añobas = Year(Now)
ElseIf Month(Now) = 9 Then
    meshoypal = "septiembre"
    mesbas = Format(7, "00")
    mesbaspal = "julio"
    añobas = Year(Now)
ElseIf Month(Now) = 10 Then
    meshoypal = "octubre"
    mesbas = Format(8, "00")
    mesbaspal = "agosto"
    añobas = Year(Now)
ElseIf Month(Now) = 11 Then
    meshoypal = "noviembre"
    mesbas = Format(9, "00")
    mesbaspal = "septiembre"
    añobas = Year(Now)
ElseIf Month(Now) = 12 Then
    meshoypal = "diciembre"
    mesbas = Format(10, "00")
    mesbaspal = "octubre"
    añobas = Year(Now)
End If

'FECHAS MANUALES
'diahoy = InputBox("Ingresa el día de hoy en formato de número a dos dígitos (ej. 23):")'
'meshoy = InputBox("Ingresa el mes de hoy en formato de número a dos dígitos (ej. 10):")
'añohoy = InputBox("Ingresa el año de hoy en formato de número a cuatro dígitos (ej. 2019):")
'meshoypal = InputBox("Ingresa el mes de hoy en formato de palabra en minúsculas (ej. octubre):")
'mesbas = InputBox("Ingresa el mes de la última base de datos del INEGI (dos meses atrás) en formato de número a dos dígitos (ej. 08):")
'mesbaspal = InputBox("Ingresa el mes de la última base de datos del INEGI (dos meses atrás) en formato de palabra en minúsculas (ej. agosto):")
'añobas = InputBox("Ingresa el año de la última base de datos del INEGI (dos meses atrás) en formato de número a cuatro dígitos (ej. 2019):")


' Cambiamos los espaciados del boletín
With documento.Content
    .Style = "Espaciado principal"
End With

' Insertar título del boletín
documento.Content.insertparagraphafter

With documento.Content
    .InsertAfter Hoja8.Cells(2, 1).Value ' Título del boletín [Paragraph(2)]
    .insertparagraphafter
End With

With documento.Paragraphs(2).Range
    .Style = "Título 1"
End With

' Insertar párrafo de texto MES
With documento.Content
    .InsertAfter Hoja8.Cells(5, 2).Value ' Texto MES [Paragraph(4)]
    .insertparagraphafter
End With

With documento.Paragraphs(3).Range
    .Style = "Normal"
End With

' Insertar título de gráfica MES
With documento.Content
    .InsertAfter Hoja8.Cells(6, 2).Value ' Título de gráfica MES [Paragraph(5)]
    .insertparagraphafter
End With

With documento.Paragraphs(4).Range
    .Style = "Figura - titulos"
End With

' Pasar gráfica MES
Grafica_horizontal.Chart.ChartArea.Copy
documento.Paragraphs(5).Range.Paste
documento.Content.insertparagraphafter

' Insertar fuente
With documento.Content
    .InsertAfter "Fuente: IIEG con información de INEGI." ' Nota [Paragraph(7)]
    .insertparagraphafter
End With

With documento.Paragraphs(6).Range
    .Style = "Fuentes"
End With
' Insertar salto de página
documento.Paragraphs(7).Range.InsertBreak Type:=7 'wdSectionBreakNextPage
documento.Content.insertparagraphafter

' Insertar texto de la gráfica HIS
With documento.Content
    .InsertAfter Hoja8.Cells(9, 2).Value ' Texto de la gráfica HIS [Paragraph(10)]
    .insertparagraphafter
End With

With documento.Paragraphs(9).Range
    .Style = "Normal"
End With

' Insertar título de gráfica HIS
With documento.Content
    .InsertAfter Hoja8.Cells(10, 2).Value ' Título de gráfica HIS [Paragraph(11)]
    .insertparagraphafter
End With

With documento.Paragraphs(10).Range
    .Style = "Figura - titulos"
End With

' Pasar gráfica HIS
Grafica_his.Chart.ChartArea.Copy
documento.Paragraphs(11).Range.Paste
documento.Content.insertparagraphafter

' Insertar fuente
With documento.Content
    .InsertAfter "Fuente: IIEG con información de INEGI." ' Nota [Paragraph(13)]
    .insertparagraphafter
End With

With documento.Paragraphs(12).Range
    .Style = "Fuentes"
End With

' Insertar nota
With documento.Content
    .InsertAfter Hoja8.Cells(11, 2).Value ' Nota [Paragraph(14)]
    .insertparagraphafter
End With

With documento.Paragraphs(13).Range
    .Style = "Fuentes"
End With

' Insertar salto de página
documento.Paragraphs(14).Range.InsertBreak Type:=7
documento.Content.insertparagraphafter




' Insertar texto de la gráfica RANKDIS
With documento.Content
    .InsertAfter Hoja8.Cells(14, 2).Value ' Texto de la gráfica RANKDIS [Paragraph(17)]
    .insertparagraphafter
End With

With documento.Paragraphs(16).Range
    .Style = "Normal"
End With

' Insertar título de gráfica RANKDIS
With documento.Content
    .InsertAfter Hoja8.Cells(15, 2).Value ' Título de gráfica RANKDIS [Paragraph(18)]
    .insertparagraphafter
End With

With documento.Paragraphs(17).Range
    .Style = "Figura - titulos"
End With

' Pasar gráfica RANKDIS
Grafica_vertical.Chart.ChartArea.Copy
documento.Paragraphs(18).Range.Paste
documento.Content.insertparagraphafter

' Insertar fuente
With documento.Content
    .InsertAfter "Fuente: IIEG con información de INEGI." ' Nota [Paragraph(20)]
    .insertparagraphafter
End With

With documento.Paragraphs(19).Range
    .Style = "Fuentes"
End With
' Insertar salto de página
documento.Paragraphs(20).Range.InsertBreak Type:=7
documento.Content.insertparagraphafter

' Insertar texto de la gráfica RANKVPP
With documento.Content
    .InsertAfter Hoja8.Cells(18, 2).Value ' Texto de la gráfica RANKVPP [Paragraph(23)]
    .insertparagraphafter
End With

With documento.Paragraphs(22).Range
    .Style = "Normal"
End With

' Insertar título de gráfica RANKVPP
With documento.Content
    .InsertAfter Hoja8.Cells(19, 2).Value ' Título de la gráfica RANKVPP [Paragraph(24)]
    .insertparagraphafter
End With

With documento.Paragraphs(23).Range
    .Style = "Figura - titulos"
End With

' Pasar gráfica RANKVPP
Grafica_vertical_nacional.Chart.ChartArea.Copy
documento.Paragraphs(24).Range.Paste
documento.Content.insertparagraphafter

' Insertar fuente
With documento.Content
    .InsertAfter "Fuente: IIEG con información de INEGI." ' Nota [Paragraph(26)]
    .insertparagraphafter
End With

With documento.Paragraphs(25).Range
    .Style = "Fuentes"
End With

' Insertar nota
With documento.Content
    .InsertAfter Hoja8.Cells(20, 2).Value ' Nota [Paragraph(27)]
    .insertparagraphafter
End With

With documento.Paragraphs(26).Range
    .Style = "Fuentes"
End With
' Insertar salto de página
documento.Paragraphs(27).Range.InsertBreak Type:=7
documento.Content.insertparagraphafter

' Insertar texto de la gráfica RANKVMEN
With documento.Content
    .InsertAfter Hoja8.Cells(23, 2).Value ' Texto de la gráfica RANKVMEN [Paragraph(30)]
    .insertparagraphafter
End With

With documento.Paragraphs(29).Range
    .Style = "Normal"
End With

' Insertar título de gráfica RANKVMEN
With documento.Content
    .InsertAfter Hoja8.Cells(24, 2).Value ' Título de la gráfica RANKVMEN [Paragraph(31)]
    .insertparagraphafter
End With

With documento.Paragraphs(30).Range
    .Style = "Figura - titulos"
End With

' Pasar gráfica RANKVPP
Grafica_vertical_nacional_men.Chart.ChartArea.Copy
documento.Paragraphs(31).Range.Paste
documento.Content.insertparagraphafter

' Insertar fuente
With documento.Content
    .InsertAfter "Fuente: IIEG con información de INEGI." ' Nota [Paragraph(33)]
    .insertparagraphafter
End With

With documento.Paragraphs(32).Range
    .Style = "Fuentes"
End With

' Insertar nota
With documento.Content
    .InsertAfter Hoja8.Cells(25, 2).Value ' Nota [Paragraph(34)]
    .insertparagraphafter
End With

With documento.Paragraphs(33).Range
    .Bold = False
    .Font.Size = 10
    .ParagraphFormat.Alignment = 3
End With
' Insertar salto de página
documento.Paragraphs(34).Range.InsertBreak Type:=7
documento.Content.insertparagraphafter




' Insertar texto de la gráfica SECTOR
With documento.Content
    .InsertAfter Hoja8.Cells(28, 2).Value ' Texto de la gráfica SECTOR [Paragraph(37)]
    .insertparagraphafter
End With

With documento.Paragraphs(36).Range
    .Style = "Normal"
End With

' Insertar título de gráfica SECTOR
With documento.Content
    .InsertAfter Hoja8.Cells(29, 2).Value ' Título de la gráfica SECTOR [Paragraph(38)]
    .insertparagraphafter
End With

With documento.Paragraphs(37).Range
    .Style = "Figura - titulos"
End With

' Pasar gráfica SECTOR
Grafica_Sect.Chart.ChartArea.Copy
documento.Paragraphs(38).Range.Paste
documento.Content.insertparagraphafter

' Insertar fuente
With documento.Content
    .InsertAfter "Fuente: IIEG con información de INEGI." ' Nota [Paragraph(40)]
    .insertparagraphafter
End With

With documento.Paragraphs(39).Range
    .Style = "Fuentes"
End With

' Insertar nota
With documento.Content
    .InsertAfter Hoja8.Cells(30, 2).Value ' Nota [Paragraph(41)]
    .insertparagraphafter
End With

With documento.Paragraphs(40).Range
    .Style = "Fuentes"
End With
' Insertar salto de página
documento.Paragraphs(41).Range.InsertBreak Type:=7
documento.Content.insertparagraphafter



' Insertar texto de la gráfica OBRA
With documento.Content
    .InsertAfter Hoja8.Cells(33, 2).Value ' Texto de la gráfica OBRA [Paragraph(44)]
    .insertparagraphafter
End With

With documento.Paragraphs(43).Range
    .Style = "Normal"
End With

' Insertar título de gráfica OBRA
With documento.Content
    .InsertAfter Hoja8.Cells(34, 2).Value ' Título de la gráfica OBRA [Paragraph(45)]
    .insertparagraphafter
End With

With documento.Paragraphs(44).Range
    .Style = "Figura - titulos"
End With

' Pasar gráfica OBRA
Grafica_Obra.Chart.ChartArea.Copy
documento.Paragraphs(45).Range.Paste
documento.Content.insertparagraphafter

' Insertar fuente
With documento.Content
    .InsertAfter "Fuente: IIEG con información de INEGI." ' Nota [Paragraph(47)]
    .insertparagraphafter
End With

With documento.Paragraphs(46).Range
    .Style = "Fuentes"
End With

' Insertar nota
With documento.Content
    .InsertAfter Hoja8.Cells(35, 2).Value ' Nota [Paragraph(48)]
    .insertparagraphafter
End With

With documento.Paragraphs(47).Range
    .Style = "Fuentes"
End With





' Cambiar la fecha de realización
Set cuadrofecha = documento.Sections(1).Headers(1).Shapes.AddTextbox(msoTextOrientationHorizontal, _
                  340, 35 - 7, 240, 70 / 2)
                  ' wdHeaderFooterPrimary = 1
cuadrofecha.TextFrame.TextRange.Text = "Ficha informativa, " & diahoy & " de " & meshoypal & " de " & añohoy
cuadrofecha.TextFrame.TextRange.Font.Color = RGB(98, 113, 120)
cuadrofecha.TextFrame.TextRange.Font.Underline = wdUnderlineSingle
cuadrofecha.TextFrame.TextRange.ParagraphFormat.Alignment = wdAlignParagraphRight
cuadrofecha.Fill.ForeColor = RGB(255, 255, 255)
cuadrofecha.Line.ForeColor = RGB(255, 255, 255)

' Guardar el documento
documento.SaveAs "C:\Users\arturo.carrillo\Documents\ENEC\" & añobas & " " & mesbas & "\Ficha informativa Industria de la Construccion (ENEC), " & mesbaspal & " " & añobas & "-" & añohoy & meshoy & diahoy & ".docx" ' UBICACIÓN PERSONAL

End Sub


