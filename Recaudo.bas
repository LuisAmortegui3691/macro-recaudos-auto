Attribute VB_Name = "Recaudo"
Sub recaudosAutonal()

    ' Meidcion del flujo de trabajo
    Dim tiempoInicio As Double
    Dim tiempoFin As Double
    Dim duracionSegundos As Double
    Dim duracionMinutos As Double
    
    Dim documentosEntrada As String, documentosSalida As String
    Dim archiPagosGenalse As String
    Dim valorCeldaPagosGenal As String, valorCeldaPlantilla As String
    Dim wsPlantilla As Worksheet
    Dim filasAEliminar As Range
    Dim ultimaFilaPlantilla As Long
    Dim filaEliminar As Long, ultimaFilaDatosTxt As Long
    Dim valorPlacaRecaudo, valorPlaca, valorDocuPlacas
    Dim TR_soat
    Dim valorHojaRecaudo As String
    Dim nombreLibroPlacas As String
    Dim validarCeldaConsolidado
    
    ' Registra el tiempo de inicio
    tiempoInicio = Timer
    
    valorHojaRecaudo = ThisWorkbook.Sheets("main").Range("F2").Value
    nombreLibroPlacas = ThisWorkbook.Sheets("main").Range("C4").Value
    
    i = 0
    
    documentosEntrada = ThisWorkbook.Sheets("main").Range("C2").Value
    documentosSalida = ThisWorkbook.Sheets("main").Range("C3").Value
    
    archiPagosGenalse = documentosEntrada & "Pagos Genalse\"
    archiPagosGenalse = Dir(archiPagosGenalse)
    
    Application.DisplayAlerts = False
    Workbooks.OpenText Filename:=documentosEntrada & "Pagos Genalse\" & archiPagosGenalse
    Application.DisplayAlerts = True
    
    Application.DisplayAlerts = False
    Workbooks.OpenText Filename:=documentosEntrada & "\Plantilla\plantilla_recaudo.xlsx"
    Application.DisplayAlerts = True
    
    ' Pegar valores de pagos generales a plantillla
    For i = 1 To 100
        valorCeldaPagosGenal = Workbooks(archiPagosGenalse).Sheets(valorHojaRecaudo).Range("A" & i).Value
        
        If (valorCeldaPagosGenal <> "Area" And valorCeldaPagosGenal <> "") Then
            Workbooks(archiPagosGenalse).Sheets(valorHojaRecaudo).Range("A" & i & ":J" & i).Copy
            Workbooks("plantilla_recaudo.xlsx").Sheets("Plantilla").Range("A" & i).PasteSpecial xlPasteValues
        ElseIf valorCeldaPlantilla = "ORDENES DEVUELTAS" Then
            Exit For
        End If

    Next i
    
    ' Definir la hoja de trabajo
    Set wsPlantilla = Workbooks("plantilla_recaudo.xlsx").Sheets("Plantilla")
    ' Inicializar la variable de filas a eliminar
    Set filasAEliminar = Nothing
    
    ' Iterar sobre las filas para identificar las que deben eliminarse
    For i = 1 To wsPlantilla.Rows.Count
        If Trim(wsPlantilla.Range("A" & i).Value) = "" Then
            If filasAEliminar Is Nothing Then
                Set filasAEliminar = wsPlantilla.Rows(i)
            Else
                Set filasAEliminar = Union(filasAEliminar, wsPlantilla.Rows(i))
            End If
        End If
    Next i
    
    ' Eliminar todas las filas seleccionadas
    If Not filasAEliminar Is Nothing Then
        filasAEliminar.Delete
    End If
    
   ' Idntificar la celda con valor ORDENE DEVUELTAS para eliminar las filas hacia abajo incluyendola
    For i = 1 To 100
        valorCeldaPlantilla = Trim(Workbooks("plantilla_recaudo.xlsx").Sheets("Plantilla").Range("A" & i).Value)
        
        If valorCeldaPlantilla = "ORDENES DEVUELTAS" Then
            filaEliminar = i
            Exit For  ' Salir del bucle una vez encontrada la fila
        End If
    Next i
    
    ' Eliminar la fila y las subsiguientes
    If filaEliminar > 0 Then
        Workbooks("plantilla_recaudo.xlsx").Sheets("Plantilla").Rows(filaEliminar & ":" & wsPlantilla.Rows.Count).Delete
    End If
    
    Set wsPlantilla = Workbooks("plantilla_recaudo.xlsx").Sheets("Plantilla")
    
    ' Encontrar la última fila con datos en la columna A
    ultimaFilaPlantilla = wsPlantilla.Cells(wsPlantilla.Rows.Count, "A").End(xlUp).Row
    
    
    ' Rango de datos para ordenar (asumiendo que los datos comienzan en la fila 1)
    Dim rangoDatos As Range
    Set rangoDatos = wsPlantilla.Range("A1:J" & ultimaFilaPlantilla)
    
    ' Ordenar por la columna "Tipo" (columna J)
    With wsPlantilla.Sort
        .SortFields.Clear
        .SortFields.Add Key:=rangoDatos.Columns(10), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SetRange rangoDatos
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    ' Pega las columnas para conciliar los datos
    wsPlantilla.Range("C2:E" & ultimaFilaPlantilla).Copy
    Workbooks("plantilla_recaudo.xlsx").Sheets("cruce_recaudos").Range("A2").PasteSpecial xlPasteValues
    wsPlantilla.Range("G2:G" & ultimaFilaPlantilla).Copy
    Workbooks("plantilla_recaudo.xlsx").Sheets("cruce_recaudos").Range("D2").PasteSpecial xlPasteValues
    wsPlantilla.Range("J2:J" & ultimaFilaPlantilla).Copy
    Workbooks("plantilla_recaudo.xlsx").Sheets("cruce_recaudos").Range("E2").PasteSpecial xlPasteValues
    
    ' Identifica la cantidad de datos en la hoja a conciliar
    ultimaFilaDatosTxt = ThisWorkbook.Sheets("datosTxt").Range("A" & Rows.Count).End(xlUp).Row
    
    ' Pega los datos de la macro extraida txt para pegarlos en la plantilla a conciliar
    Workbooks("Macro Recaudo.xlsm").Sheets("datosTxt").Range("C1:D" & ultimaFilaDatosTxt).Copy
    Workbooks("plantilla_recaudo.xlsx").Sheets("cruce_recaudos").Range("B" & ultimaFilaPlantilla + 1).PasteSpecial xlPasteValues
    Workbooks("Macro Recaudo.xlsm").Sheets("datosTxt").Range("E1:E" & ultimaFilaDatosTxt).Copy
    Workbooks("plantilla_recaudo.xlsx").Sheets("cruce_recaudos").Range("A" & ultimaFilaPlantilla + 1).PasteSpecial xlPasteValues
    Workbooks("Macro Recaudo.xlsm").Sheets("datosTxt").Range("G1:G" & ultimaFilaDatosTxt).Copy
    Workbooks("plantilla_recaudo.xlsx").Sheets("cruce_recaudos").Range("D" & ultimaFilaPlantilla + 1).PasteSpecial xlPasteValues
    Workbooks("Macro Recaudo.xlsm").Sheets("datosTxt").Range("H1:H" & ultimaFilaDatosTxt).Copy
    Workbooks("plantilla_recaudo.xlsx").Sheets("cruce_recaudos").Range("E" & ultimaFilaPlantilla + 1).PasteSpecial xlPasteValues

    archivoRecaudo = documentosEntrada & "Recaudos\"
    archivoRecaudo = Dir(archivoRecaudo)
    
    Application.DisplayAlerts = False
    Workbooks.OpenText Filename:=documentosEntrada & "Recaudos\" & archivoRecaudo
    Application.DisplayAlerts = True
    
    ' Cerrar archivo Pagos generales
    Windows(archiPagosGenalse).Activate
    ActiveWorkbook.Close SaveChanges:=False
    
    ultimaFilaRecaudo = Workbooks(archivoRecaudo).Sheets("GeneralRecaudo").Range("A" & Rows.Count).End(xlUp).Row
    
    ultimaFilaConsolidado = Workbooks("plantilla_recaudo.xlsx").Sheets("cruce_recaudos").Range("B" & Rows.Count).End(xlUp).Row

    For i = 1 To ultimaFilaConsolidado
        valorCeldaCruceCliente = Workbooks("plantilla_recaudo.xlsx").Sheets("cruce_recaudos").Range("B" & i).Value
        valorCeldaCruceTipo = Workbooks("plantilla_recaudo.xlsx").Sheets("cruce_recaudos").Range("E" & i).Value
        valorCeldaCrucePlacaRecaudo = Workbooks("plantilla_recaudo.xlsx").Sheets("cruce_recaudos").Range("A" & i).Value
        validarCeldaConsolidado = Workbooks("plantilla_recaudo.xlsx").Sheets("cruce_recaudos").Range("G" & i).Value
        
        
        If valorCeldaCruceTipo = "POLIZA" Then
            If validarCeldaConsolidado = "" Then
                For j = 1 To ultimaFilaRecaudo
                    valorAO = Workbooks(archivoRecaudo).Sheets("GeneralRecaudo").Range("AO" & j).Value
                    valorIdRecaudo = Workbooks(archivoRecaudo).Sheets("GeneralRecaudo").Range("L" & j).Value
                    valorTR = Workbooks(archivoRecaudo).Sheets("GeneralRecaudo").Range("A" & j).Value
                    
                    If valorAO <> "ok" Then
                        If valorIdRecaudo = valorCeldaCruceCliente And valorTR = "DS" Then
                            TR = Workbooks(archivoRecaudo).Sheets("GeneralRecaudo").Range("A" & j).Value
                            numeroRecaudo = Workbooks(archivoRecaudo).Sheets("GeneralRecaudo").Range("B" & j).Value
                            placa = Workbooks(archivoRecaudo).Sheets("GeneralRecaudo").Range("K" & j).Value
                            totalRecaudo = Workbooks(archivoRecaudo).Sheets("GeneralRecaudo").Range("N" & j).Value
                            Workbooks(archivoRecaudo).Activate
                            Workbooks(archivoRecaudo).Sheets("GeneralRecaudo").Rows(j).Select
                            ' Cambia el color de fondo de la selección al color #ffb3ff
                            With Selection.Interior
                                .Color = RGB(102, 255, 102) ' Código RGB para el color #ffb3ff
                            End With
                            ' Quita la selección
                            Application.CutCopyMode = False
                            Workbooks(archivoRecaudo).Sheets("GeneralRecaudo").Range("AO" & j).Value = "ok"
                        End If
                    End If
                Next j
            End If
            
            Workbooks("plantilla_recaudo.xlsx").Sheets("cruce_recaudos").Range("G" & i).Value = TR
            Workbooks("plantilla_recaudo.xlsx").Sheets("cruce_recaudos").Range("H" & i).Value = numeroRecaudo
            Workbooks("plantilla_recaudo.xlsx").Sheets("cruce_recaudos").Range("I" & i).Value = placa
            Workbooks("plantilla_recaudo.xlsx").Sheets("cruce_recaudos").Range("J" & i).Value = totalRecaudo
        End If
        
        ' Validar SOAT para traer placa de los SOAT
        If valorCeldaCruceTipo = "SOAT" Then
            Application.DisplayAlerts = False
            Workbooks.OpenText Filename:=documentosEntrada & "Placas\PAGO SOAT AUTONAL 01 FEBRERO 2024.xlsx"
            Application.DisplayAlerts = True
            
            For k = 1 To 50
                valorDocuPlacas = Workbooks(nombreLibroPlacas).Sheets("Hoja1").Range("B" & k).Value
                
                If valorDocuPlacas = valorCeldaCruceCliente Then
                    valorPlaca = Workbooks(nombreLibroPlacas).Sheets("Hoja1").Range("A" & k).Value
                    Workbooks("plantilla_recaudo.xlsx").Sheets("cruce_recaudos").Range("A" & i).Value = valorPlaca
                End If
            Next k
        End If
        
        ' Conciliacion de recaudos
        If valorCeldaCruceTipo = "SOAT" Then
            If validarCeldaConsolidado = "" Then
                ' Captura valores e iteracion
                For l = 1 To ultimaFilaRecaudo
                    valorAO = Workbooks(archivoRecaudo).Sheets("GeneralRecaudo").Range("AO" & j).Value
                    valorPlacaRecaudo = Workbooks(archivoRecaudo).Sheets("GeneralRecaudo").Range("K" & l).Value
                    valorCeldaCrucePlacaRecaudo = Workbooks("plantilla_recaudo.xlsx").Sheets("cruce_recaudos").Range("A" & i).Value
                    valorTR = Workbooks(archivoRecaudo).Sheets("GeneralRecaudo").Range("A" & l).Value
                    
                    If valorAO <> "ok" Then
                        ' Condicion para atrapar los datos requeridos
                        If valorPlacaRecaudo = valorCeldaCrucePlacaRecaudo And valorTR = "RS" Then
                            TR_soat = Workbooks(archivoRecaudo).Sheets("GeneralRecaudo").Range("A" & l).Value
                            numeroRecaudo_soat = Workbooks(archivoRecaudo).Sheets("GeneralRecaudo").Range("B" & l).Value
                            placa_soat = Workbooks(archivoRecaudo).Sheets("GeneralRecaudo").Range("K" & l).Value
                            totalRecaudo_soat = Workbooks(archivoRecaudo).Sheets("GeneralRecaudo").Range("N" & l).Value
                            Workbooks(archivoRecaudo).Activate
                            Workbooks(archivoRecaudo).Sheets("GeneralRecaudo").Rows(l).Select
                            ' Cambia el color de fondo de la selección al color #ffb3ff
                            With Selection.Interior
                                .Color = RGB(204, 0, 204) ' Código RGB para el color #ffb3ff
                            End With
                            ' Quita la selección
                            Application.CutCopyMode = False
                            
                            Workbooks("plantilla_recaudo.xlsx").Sheets("cruce_recaudos").Range("G" & i).Value = TR_soat
                            Workbooks("plantilla_recaudo.xlsx").Sheets("cruce_recaudos").Range("H" & i).Value = numeroRecaudo_soat
                            Workbooks("plantilla_recaudo.xlsx").Sheets("cruce_recaudos").Range("I" & i).Value = placa_soat
                            Workbooks("plantilla_recaudo.xlsx").Sheets("cruce_recaudos").Range("J" & i).Value = totalRecaudo_soat
                            Workbooks(archivoRecaudo).Sheets("GeneralRecaudo").Range("AO" & l).Value = "ok"
                        End If
                    End If
                Next l
            End If
        End If
    Next i

     ' Registra el tiempo de finalización
    tiempoFin = Timer
    ' Calcula la duración en segundos
    duracionSegundos = tiempoFin - tiempoInicio
    
    ' Muestra los resultados en la ventana inmediata (puedes ajustar esto según tus necesidades)
    Debug.Print "Duración en segundos: " & duracionSegundos

End Sub
