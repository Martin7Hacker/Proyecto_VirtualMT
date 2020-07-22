Attribute VB_Name = "Imprimir"
'***************************************************************************
'*
'*
'* Imprimir Guardar Ficheros con Virtual Martin temporize v1.0
'*
'*
'***************************************************************************
'A esta función se le envía el control LV a imprimir
Public Sub Imprimir_ListView()
 Dim i As Integer, AnchoCol As Single, espacio As Integer, x As Integer
 AnchoCol = 0
 'Recorremos desde la primer columna hasta la última para almacenar el ancho total
 For i = 1 To frmVisorEventos.ListView1.ColumnHeaders.Count
 AnchoCol = AnchoCol + frmVisorEventos.ListView1.ColumnHeaders(i).Width
 Next
 espacio = 0
 'Encabezado de ejemplo
 Printer.Print "-------------------------------------------------------------"
 Printer.Print " Comienzo de la Impresión - Timbres    "
 Printer.Print "-------------------------------------------------------------"
 Printer.Print
 'Imprime una línea
 Printer.Line (Printer.CurrentX, Printer.CurrentY)-(Printer.ScaleWidth, Printer.CurrentY)
 With frmVisorEventos.ListView1
 Printer.Print "ID"
 'Acá se imprimen los encabezados del ListView
 For i = 1 To .ColumnHeaders.Count
 espacio = espacio + CInt(.ColumnHeaders(i).Width * Printer.ScaleWidth / AnchoCol)
 Printer.Print frmVisorEventos.ListView1.ColumnHeaders(i).Text;
 Printer.CurrentX = espacio
 Next
 'Printer.Print frmVisorEventos.ListView1.ColumnHeaders(1).Text & frmVisorEventos.ListView1.ColumnHeaders(2).Text & frmVisorEventos.ListView1.ColumnHeaders(3).Text & frmVisorEventos.ListView1.ColumnHeaders(4).Text & frmVisorEventos.ListView1.ColumnHeaders(5).Text;
 Printer.Print
 'Imprime una línea
 Printer.Line (Printer.CurrentX, Printer.CurrentY)-(Printer.ScaleWidth, Printer.CurrentY)
 'Imprime Línea en blanco
 Printer.Print
 'Este bucle recorre los items y subitems del ListView  y los imprime
 For i = 1 To .ListItems.Count
 espacio = 0
 Printer.Print frmVisorEventos.ListView1.ListItems.Item(i).Text;
 'Recorremos las columnas
 For x = 1 To .ColumnHeaders.Count - 1
 espacio = espacio + CInt(.ColumnHeaders(x).Width * _
 Printer.ScaleWidth / AnchoCol)
 Printer.CurrentX = espacio
 Printer.Print frmVisorEventos.ListView1.ListItems.Item(i).SubItems(x);
 Next
 'Otro espacio en blanco
 Printer.Print
 Next
 End With
 Printer.Print
 'Imprime la línea de final de impresión
 Printer.Line (Printer.CurrentX, Printer.CurrentY)-(Printer.ScaleWidth, Printer.CurrentY)
 Printer.Print
 'Texto del pie
 Printer.Print "Software creado por: Martin Grasso Castrillo - Martinsoft 2012  "
 Printer.Print "fecha: " & Date & " - " & " hora: " & Time
 Printer.Line (Printer.CurrentX, Printer.CurrentY)-(Printer.ScaleWidth, Printer.CurrentY)
 Printer.Print
 Printer.Print "++++++++++++++++++++++++ Fin de la impresión ++++++++++++++++++++++++ "
 'Comenzamos la impresión
 Printer.EndDoc
End Sub

