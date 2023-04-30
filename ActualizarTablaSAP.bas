Attribute VB_Name = "Módulo1"
Sub Actualizar_tablaSAP()
    Application.ScreenUpdating = False      'Evitar el Parpadeo cuando se seleccionan y actualizan valores

    Dim compra_sap, almacen, tabla_dinamica As String           'Variables que guardan nombres necesarios (como de la hoja a usar)
    Dim rango_caducidad, rango_celda_tablaSAP, rango_total_tarimas, rango_eliminar, _
        rango_material, rango_a_copiar, rango_a_pegar, rango_tabla_almacen As String    'Variables que almacenan ubicaciones de celda necesarias
    compra_sap = "Emb SAP"
    almacen = "Edo Almacén"
    tabla_dinamica = "TablaDinámica8"
    
    Dim vida_producto, suma_tarimas, cantidad_sumar, registros_actualizados As Integer
    Dim filaActual_almacen, filaActual_tablaSAP, columnaActual_tablaSAP  As Integer
    filaActual_tablaSAP = 2         'Fila actual donde se empezara a llenar la informacion
    columnaActual_tablaSAP = 13     'Columna actual donde empezara a llenar la informacion
    registros_actualizados = 0
    
    Dim hayTarimas As Boolean
    hayTarimas = True
    
    Worksheets(almacen).PivotTables(tabla_dinamica).ClearAllFilters     'Limpia los filtros si los tiene activados
    Worksheets(almacen).PivotTables(tabla_dinamica).PivotCache.Refresh   'Actualiza la tabla dinamica
    
    'Borra la tabla provisional anterior
    Dim ultima_fila As Integer
    ultima_fila = Worksheets(almacen).Cells(Rows.Count, 16).End(xlUp).Row        'Ubica la ultima fila
    Worksheets(almacen).Range("M2").AutoFilter
    rango_eliminar = "M2:Q" + CStr(ultima_fila)
    Worksheets(almacen).Range(rango_eliminar).ClearContents

    
    'Llena una tabla provisional para usarla en la busqueda de tarimas, con origen en la tabla dinamica
    ultima_fila = Worksheets(almacen).Cells(Rows.Count, 5).End(xlUp).Row        'Ubica la ultima fila
    rango_tabla_almacen = "A2:E" + CStr(ultima_fila)
    Worksheets(almacen).Range(rango_tabla_almacen).Copy                         'Copia la tabla dinamica
    Worksheets(almacen).Range("M2").PasteSpecial xlPasteValues
    'Llena valores vacios de la tabla, ya que no se copian como son necesarios
    filaActual_almacen = 3
    While Worksheets(almacen).Cells(filaActual_almacen + 1, 16) <> ""
        If Worksheets(almacen).Cells(filaActual_almacen + 1, 13) = "" Then
            Worksheets(almacen).Cells(filaActual_almacen + 1, 13) = Worksheets(almacen).Cells(filaActual_almacen, 13)
        End If
        If Worksheets(almacen).Cells(filaActual_almacen + 1, 14) = "" Then
            Worksheets(almacen).Cells(filaActual_almacen + 1, 14) = Worksheets(almacen).Cells(filaActual_almacen, 14)
        End If
        filaActual_almacen = filaActual_almacen + 1
    Wend
    'Ordena la tabla
    Worksheets(almacen).Range("M1").Sort Key1:=Worksheets(almacen).Range("M2"), Order1:=xlAscending, Header:=xlYes
    
    'Elimina datos existentes de las tarimas obtenidas anteriormente en Emb SAP
    ultima_fila = Worksheets(compra_sap).Cells(Rows.Count, 4).End(xlUp).Row
    rango_eliminar = "M" + CStr(filaActual_tablaSAP) + ":AM" + CStr(ultima_fila)
    Worksheets(compra_sap).Range(rango_eliminar).ClearContents
        
    'Se realiza el llenado de campos de las tarimas necesarias para cumplir con la concesion
    While Worksheets(compra_sap).Cells(filaActual_tablaSAP + 1, columnaActual_tablaSAP - 1) <> ""
        filaActual_almacen = 3              'Reseteo de la fila donde empezamos a buscar tarimas
        
        suma_tarimas = 0                    'Variable para determinar total de tarimas actual para el llenado de la fila
        vida_producto = 0
        hayTarimas = True
        
        rango_caducidad = "L" + CStr(filaActual_tablaSAP)       'Celdas donde se ubica valores necesarios para calcular el llenado de registros (en la tabla Emb SAP)
        rango_total_tarimas = "H" + CStr(filaActual_tablaSAP)
        rango_material = "G" + CStr(filaActual_tablaSAP)
        rango_celda_tablaSAP = "M" + CStr(filaActual_tablaSAP)

        'Funcion que garantiza el llenado de registros mientras se tenga concesion y el uso de otros criterios
        If Worksheets(compra_sap).Range(rango_celda_tablaSAP) = "" And Worksheets(compra_sap).Range(rango_caducidad).Value <> "NO" And Worksheets(compra_sap).Range(rango_total_tarimas).Value <> 0 Then
            If Worksheets(compra_sap).Range(rango_caducidad).Value < 90 Then
                'Filtro de buscar producto con DDV mayor a la vida del producto (DDV) solicitada
                Worksheets(almacen).Range("M2:Q2").AutoFilter _
                Field:=1, _
                Criteria1:=">=" + CStr(Worksheets(compra_sap).Range(rango_caducidad))
                'Filtro del material a buscar
                Worksheets(almacen).Range("M2:Q2").AutoFilter _
                Field:=2, _
                Criteria1:=Worksheets(compra_sap).Range(rango_material)
                
                'Funcion que realiza el calculo de las tarimas que se van teniendo en la operacion del llenado de registros
                While suma_tarimas < Worksheets(compra_sap).Range(rango_total_tarimas) And hayTarimas
                    rango_a_copiar = "P" + CStr(filaActual_almacen)
                    
                    'Funcion que busca las filas que conciden con los criterios utilizados en los filtros
                    While (Worksheets(almacen).Range(rango_a_copiar).EntireRow.Hidden = True Or Worksheets(almacen).Range("Q" + CStr(filaActual_almacen)) = 0) And hayTarimas
                        'Esta funcion se encarga de observar si hay suficientes tarimas para utilizar en el llenado
                        If Worksheets(almacen).Cells(filaActual_almacen, 16) = "" Then
                            hayTarimas = False
                            MsgBox ("Faltan tarimas en campo " + CStr(filaActual_tablaSAP) + ", Revise su almacen")
                            If (suma_tarimas = 0) Then
                                registros_actualizados = registros_actualizados - 1
                            End If
                        End If
                        filaActual_almacen = filaActual_almacen + 1
                        rango_a_copiar = "P" + CStr(filaActual_almacen)
                    Wend
                    
                    'Llena la informacion aqui
                    If hayTarimas Then
                        '//Copia la ubicacion
                        Worksheets(almacen).Range(rango_a_copiar).Copy
                        Worksheets(compra_sap).Cells(filaActual_tablaSAP, columnaActual_tablaSAP).PasteSpecial xlPasteValues
                        '//Copia las tarimas
                        rango_a_copiar = "Q" + CStr(filaActual_almacen)
                        cantidad_sumar = Worksheets(almacen).Range(rango_a_copiar)
                        suma_tarimas = suma_tarimas + cantidad_sumar
                        If suma_tarimas > Worksheets(compra_sap).Range(rango_total_tarimas) Then
                            cantidad_sumar = cantidad_sumar - (suma_tarimas - Worksheets(compra_sap).Range(rango_total_tarimas))
                            Worksheets(compra_sap).Cells(filaActual_tablaSAP, columnaActual_tablaSAP + 1) = cantidad_sumar
                            Worksheets(almacen).Range(rango_a_copiar) = Worksheets(almacen).Range(rango_a_copiar) - cantidad_sumar
                        Else
                            Worksheets(almacen).Range(rango_a_copiar).Copy
                            Worksheets(compra_sap).Cells(filaActual_tablaSAP, columnaActual_tablaSAP + 1).PasteSpecial xlPasteValues
                            Worksheets(almacen).Range(rango_a_copiar) = 0
                        End If
                        '//Copia el DDV
                        rango_a_copiar = "M" + CStr(filaActual_almacen)
                        vida_producto = Worksheets(almacen).Range(rango_a_copiar)
                        Worksheets(compra_sap).Cells(filaActual_tablaSAP, columnaActual_tablaSAP + 2).Value = vida_producto
                    
                        'Valores para buscar y llenar mas tarimas (si es necesario)
                        filaActual_almacen = filaActual_almacen + 1
                        columnaActual_tablaSAP = columnaActual_tablaSAP + 3
                    End If
                Wend
            Else
                'Si el DDV es mayor a 90, se coloca "De acuerdo a HH"
                Worksheets(compra_sap).Cells(filaActual_tablaSAP, columnaActual_tablaSAP).Value = "De acuerdo a HH"
            End If
            registros_actualizados = registros_actualizados + 1             'Contador de registros agregados
        End If
        'A llenar la siguiente fila de Emb SAP
        filaActual_tablaSAP = filaActual_tablaSAP + 1
        columnaActual_tablaSAP = 13
    Wend
    Application.CutCopyMode = False
    MsgBox ("Registros actualizados: " + CStr(registros_actualizados))
    Application.ScreenUpdating = True          'Regresarlo a verdadero por si acaso lo usaran
End Sub



