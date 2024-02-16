Attribute VB_Name = "OrnedenesCompra"
Sub convertirTXT()
    
    ' Definir variables
    Dim rutaTXT As String
    Dim ws As Worksheet
    Dim fila As Integer
    
    ' Especificar la ruta del archivo de texto
    rutaTXT = "C:\Macros\Macro Recaudos Autonal\Documentos entrada\Ordenes de Compra txt\datos_ordenes_compra.txt"
    
    ' Definir la hoja de trabajo donde se importarán los datos
    Set ws = ThisWorkbook.Sheets("datosTxt")  ' Cambia "Sheet1" al nombre de tu hoja
    
    ' Abrir el archivo de texto para lectura
    Open rutaTXT For Input As #1
    
    ' Inicializar variable de fila
    fila = 1
    
    ' Leer datos desde el archivo de texto
    Do Until EOF(1)
        Dim linea As String
        Line Input #1, linea
        Dim datos() As String
        datos = Split(linea, ";")
        
        ' Escribir datos en la hoja de Excel
        For i = LBound(datos) To UBound(datos)
            ws.Cells(fila, i + 1).Value = Trim(datos(i))
        Next i
        
        ' Mover a la siguiente fila
        fila = fila + 1
    Loop
    
    ' Cerrar el archivo de texto
    Close #1
    
    'MsgBox "Datos importados correctamente."

End Sub
