Attribute VB_Name = "Borrar"
Sub borrarDatosMacro()

    Dim ultimaFilaTXT As Long
    
    ultimaFilaTXT = ThisWorkbook.Sheets("datosTxt").Range("A" & Rows.Count).End(xlUp).Row
    
    ThisWorkbook.Sheets("datosTxt").Range("A1:G" & ultimaFilaTXT).ClearContents
    
End Sub
