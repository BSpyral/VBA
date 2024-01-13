Attribute VB_Name = "AcomodaUbicaciones"
Sub acomodar_tarimas()
    Application.ScreenUpdating = True      'Evitar el Parpadeo cuando se seleccionan y actualizan valores

    Dim cambiarUbicaciones, almacen, tabla_dinamica As String     'Nombres necesarios (como de la hoja o la tabla dinamica a usar)
    cambiarUbicaciones = "UbicaciCambiar"
    almacen = "TablaDin"
    tabla_dinamica = "TablaDinámica2"
    
    Worksheets(almacen).PivotTables(tabla_dinamica).ClearAllFilters     'Limpia los filtros si los tuviera activados
    Worksheets(almacen).PivotTables(tabla_dinamica).PivotCache.Refresh   'Actualiza la tabla dinamica
    Worksheets(cambiarUbicaciones).Range("A1:F1").AutoFilter               'LimpiarFiltros si los tuviera
 
    Dim filaActual_almacen, filaActual_tablaUbicaciones, columnaActual_tablaUbicaciones  As Integer
    columnaActual_tablaUbicaciones = 1
     
    'Borra la tabla anterior
    Dim ultima_fila As Integer
    Dim rango_eliminar As String
    ultima_fila = Cells(Rows.Count, columnaActual_tablaUbicaciones).End(xlUp).Row
    rango_eliminar = "A2:F" + CStr(ultima_fila)
    Range(rango_eliminar).ClearContents
    
    'Llena la tabla para usarla en la reubicacion de tarimas, con origen en la tabla dinamica
    Dim rango_tabla_almacen As String
    ultima_fila = Worksheets(almacen).Cells(Rows.Count, 5).End(xlUp).Row
    rango_tabla_almacen = "A5:E" + CStr(ultima_fila - 1)
    Worksheets(almacen).Range(rango_tabla_almacen).Copy                         'Copia la tabla dinamica
    Worksheets(cambiarUbicaciones).Range("A2").PasteSpecial xlPasteValues
    
    'Llena y elimina valores vacios de la tabla, ya que no se copian como se necesita en la reubicacion
    filaActual_tablaUbicaciones = 2         'Fila actual donde se empezo a llenar la informacion anterior
    columnaActual_tablaUbicaciones = 1      'Columna actual donde empezo a llenar la informacion anterior
    While Worksheets(cambiarUbicaciones).Cells(filaActual_tablaUbicaciones + 1, columnaActual_tablaUbicaciones + 4) <> ""
        While Worksheets(cambiarUbicaciones).Cells(filaActual_tablaUbicaciones, columnaActual_tablaUbicaciones + 3) = ""
            Worksheets(cambiarUbicaciones).Cells(filaActual_tablaUbicaciones, columnaActual_tablaUbicaciones).EntireRow.Delete
            filaActual_tablaUbicaciones = filaActual_tablaUbicaciones - 1
        Wend
        If Worksheets(cambiarUbicaciones).Cells(filaActual_tablaUbicaciones + 1, columnaActual_tablaUbicaciones) = "" Then
            Worksheets(cambiarUbicaciones).Cells(filaActual_tablaUbicaciones + 1, columnaActual_tablaUbicaciones) = _
                        Worksheets(cambiarUbicaciones).Cells(filaActual_tablaUbicaciones, columnaActual_tablaUbicaciones)
        End If
        If Worksheets(cambiarUbicaciones).Cells(filaActual_tablaUbicaciones + 1, columnaActual_tablaUbicaciones + 1) = "" Then
            Worksheets(cambiarUbicaciones).Cells(filaActual_tablaUbicaciones + 1, columnaActual_tablaUbicaciones + 1) = _
                        Worksheets(cambiarUbicaciones).Cells(filaActual_tablaUbicaciones, columnaActual_tablaUbicaciones + 1)
        End If
        filaActual_tablaUbicaciones = filaActual_tablaUbicaciones + 1
    Wend
    Dim rangoFecha As String
    rangoFecha = "C2:C" + CStr(ultima_fila)
    Range(rangoFecha).NumberFormat = "dd/mm/yyyy"
    
    'Ordena la tabla
    Worksheets(cambiarUbicaciones).Range("A1").Sort _
            Key1:=Worksheets(cambiarUbicaciones).Range("A1"), _
            Key2:=Worksheets(cambiarUbicaciones).Range("B1"), _
            Key3:=Worksheets(cambiarUbicaciones).Range("D1"), _
            Order1:=xlAscending, Header:=xlYes
            
    
    
    'Obtener materiales con tarimas maximas diferentes de 20
    filaActual_tablaUbicaciones = 2
    columnaActual_tablaUbicaciones = 1
    Dim materialSaborizadas(50) As String
    Dim rango_saborizadas As String
    Dim cantidadSaborizadas As Integer
    
    rango_saborizadas = "L" + CStr(filaActual_tablaUbicaciones)
    cantidadSaborizadas = 0
    While Worksheets(cambiarUbicaciones).Range(rango_saborizadas) <> ""
        materialSaborizadas(cantidadSaborizadas) = Worksheets(cambiarUbicaciones).Range(rango_saborizadas).Value
        cantidadSaborizadas = cantidadSaborizadas + 1
        filaActual_tablaUbicaciones = filaActual_tablaUbicaciones + 1
        rango_saborizadas = "L" + CStr(filaActual_tablaUbicaciones)
    Wend
            
    'Aqui se guarda la informacion de tarimas no llenas para su posterior uso
    Dim ubicac_noLlenas(60) As String
    Dim tarim_noLlenas(60) As Integer
    Dim ddv_noLlenas(60) As Integer
    Dim valorActualArreglo As Integer
    valorActualArreglo = 0
        
    Dim rango_material, rango_ubicacion, rango_tarimas, _
        rango_materialSiguiente, rango_ubicacionSiguiente, rangoDDV As String
    Dim ubicacionActual, ubicacionSiguiente, material, materialSiguiente As String
    Dim total_tarimas, ddvMasBajo, tarimaMax, contador As Integer
    Dim esMismoMaterial, esSaborizada As Boolean
    filaActual_tablaUbicaciones = 2
    columnaActual_tablaUbicaciones = 1
    total_tarimas = 0
    ddvMasBajo = 0
    
    'Busca las ubicaciones de tarimas no llenas,por material y las guarda para poder reubicarlas despues
    While (Worksheets(cambiarUbicaciones).Cells(filaActual_tablaUbicaciones, columnaActual_tablaUbicaciones) <> "")
        
        esMismoMaterial = True
        
        
        While (esMismoMaterial)
            esSaborizada = False
            contador = 0
        
            rango_material = "A" + CStr(filaActual_tablaUbicaciones)
            rango_ubicacion = "B" + CStr(filaActual_tablaUbicaciones)
            rangoDDV = "D" + CStr(filaActual_tablaUbicaciones)
            rango_tarimas = "E" + CStr(filaActual_tablaUbicaciones)
            
            rango_materialSiguiente = "A" + CStr(filaActual_tablaUbicaciones + 1)
            rango_ubicacionSiguiente = "B" + CStr(filaActual_tablaUbicaciones + 1)
            
            
            material = Worksheets(cambiarUbicaciones).Range(rango_material)
            materialSiguiente = Worksheets(cambiarUbicaciones).Range(rango_materialSiguiente)
            ubicacionActual = Worksheets(cambiarUbicaciones).Range(rango_ubicacion)
            ubicacionSiguiente = Worksheets(cambiarUbicaciones).Range(rango_ubicacionSiguiente)
            
            If Worksheets(cambiarUbicaciones).Range(rangoDDV) > 1000 Then
                ddvMasBajo = 1000
            ElseIf ddvMasBajo = 0 Then
                ddvMasBajo = Worksheets(cambiarUbicaciones).Range(rangoDDV)
            End If
            
            Do While cantidadSaborizadas > contador
                If materialSaborizadas(contador) = material Then
                    esSaborizada = True
                    Exit Do
                End If
                contador = contador + 1
            Loop
            If esSaborizada Then
                tarimaMax = 19
            Else
                tarimaMax = 20
            End If
            
            total_tarimas = total_tarimas + Worksheets(cambiarUbicaciones).Range(rango_tarimas)
            
            If (material = materialSiguiente) Then
                filaActual_tablaUbicaciones = filaActual_tablaUbicaciones + 1
                If (ubicacionActual <> ubicacionSiguiente) Then
                    If total_tarimas < tarimaMax Then
                        'Se realiza lo necesario para guardar las tarimas no vacias
                        valorActualArreglo = guardarTarimas_yUbicacion(ubicacionActual, total_tarimas, _
                            ddvMasBajo, ubicac_noLlenas, tarim_noLlenas, ddv_noLlenas, valorActualArreglo)
                    End If
                    total_tarimas = 0
                    ddvMasBajo = 0
                End If
            Else
                esMismoMaterial = False
                ddvMasBajo = 0
            End If
        Wend
    If (valorActualArreglo > 1) Then
        'Aqui se reacomodan las tarimas no llenas, solo un material a la vez
        Call reacomodarTarimas(material, ubicac_noLlenas, tarim_noLlenas, ddv_noLlenas, valorActualArreglo, tarimaMax)
    End If
    
    Erase ubicac_noLlenas
    Erase tarim_noLlenas
    Erase ddv_noLlenas
    valorActualArreglo = 0
    
    filaActual_tablaUbicaciones = filaActual_tablaUbicaciones + 1
    Wend
    
    
    Worksheets(cambiarUbicaciones).Range("A1:F1").AutoFilter       'LimpiarFiltros
    
    'Quizas Agregar filtro para ver las modificaciones (Obtener de: <> de vacio)
    
    Application.CutCopyMode = False
    Application.ScreenUpdating = True          'Regresarlo a verdadero por si acaso se usara
End Sub

Function guardarTarimas_yUbicacion(ByVal ubicacionActual As String, ByVal total_tarimas As Integer, _
            ByVal ddvMasBajo As Integer, ubicac_noLlenas() As String, tarim_noLlenas() As Integer, _
            ddv_noLlenas() As Integer, ByVal valorActualArreglo As Integer) As Integer
    guardarTarimas_yUbicacion = valorActualArreglo
        
    'Restricciones de ubicaciones que no se usan en el llenado de tarimas
    Dim filaActual As Integer
    Dim cambiarUbicaciones As String
    filaActual = 2
    cambiarUbicaciones = "UbicaciCambiar"
    
    Dim restricciones(50) As String
    restricciones(0) = "h"
    restricciones(1) = "p"
    restricciones(2) = "calidad"
    restricciones(3) = "picking"
    Dim rango_ubicacionesDanadas As String
    Dim cantidadRestricciones As Integer
    cantidadRestricciones = 4
    
    rango_ubicacionesDanadas = "N" + CStr(filaActual)
    While Worksheets(cambiarUbicaciones).Range(rango_ubicacionesDanadas) <> ""
        restricciones(cantidadRestricciones) = Worksheets(cambiarUbicaciones).Range(rango_ubicacionesDanadas).Value
        cantidadRestricciones = cantidadRestricciones + 1
        filaActual = filaActual + 1
        rango_ubicacionesDanadas = "N" + CStr(filaActual)
    Wend
    
    
    Dim indiceUbicacion As Double
    Dim indiceBuscar As Integer
    Dim contadorArreglo As Integer
    contadorArreglo = 0
    indiceBuscar = 3
    
    While contadorArreglo < cantidadRestricciones
        If contadorArreglo > 1 Then
            indiceBuscar = 1
        End If
        On Error GoTo ErrorHandler
        indiceUbicacion = Application.WorksheetFunction.Search(restricciones(contadorArreglo), LCase(ubicacionActual), indiceBuscar)
        
        If indiceUbicacion > 0 Then Exit Function
        
        contadorArreglo = contadorArreglo + 1
    Wend
    
    'Almacenamiento de informacion de tarimas no llenas
    ubicac_noLlenas(valorActualArreglo) = ubicacionActual
    tarim_noLlenas(valorActualArreglo) = total_tarimas
    ddv_noLlenas(valorActualArreglo) = ddvMasBajo
    valorActualArreglo = valorActualArreglo + 1
    guardarTarimas_yUbicacion = valorActualArreglo
    
ErrorHandler:
    indiceUbicacion = -1
Resume Next
        
End Function

Function reacomodarTarimas(ByVal material As String, ubicac_noLlenas() As String, _
                    tarim_noLlenas() As Integer, ddv_masBajo() As Integer, _
                    ByVal tamañoArreglo As Integer, ByVal tarimaMaxima As Integer)
    Dim auxiliar, contador, filaActual As Integer
    Dim registros_noLlenos, ubicacionRegistros(100) As Integer
    Dim tarimasReubicadas, indiceArregloMayorTarima, tarimasJuntar, _
        tamArreglo, rangoMaxDdv, tarimaMax As Integer
    Dim esFilaLlenada As Boolean
    Dim rangoActual, ubicacionCambiar, ubicacionMayorTarima As String
    Dim cambiarUbicaciones, valorCelda As String
    cambiarUbicaciones = "UbicaciCambiar"
    
    'Filtro del material a buscar
    Worksheets(cambiarUbicaciones).Range("A1:F1").AutoFilter _
    Field:=1, _
    Criteria1:=material
    
    'Ubicacion de donde se empieza el material en la hoja de excel
    filaActual = 2
    rangoActual = "A" + CStr(filaActual)
    While Worksheets(cambiarUbicaciones).Range(rangoActual).EntireRow.Hidden
        filaActual = filaActual + 1
        rangoActual = "A" + CStr(filaActual)
    Wend
    
    'Filtro de las ubicaciones a buscar
    Worksheets(cambiarUbicaciones).Range("A1:F1").AutoFilter _
    Field:=2, _
    Criteria1:=Array(ubicac_noLlenas), _
    Operator:=xlFilterValues
    
    'Ubicacion de las filas con tarimas no llenas
    tamArreglo = 0
    contador = filaActual
    While Worksheets(cambiarUbicaciones).Range(rangoActual) <> ""
        If Worksheets(cambiarUbicaciones).Range(rangoActual).EntireRow.Hidden = False Then
            ubicacionRegistros(tamArreglo) = contador
            tamArreglo = tamArreglo + 1
        End If
        contador = contador + 1
        rangoActual = "A" + CStr(contador)
    Wend
    registros_noLlenos = tamArreglo
    
    tarimasReubicadas = 0
    auxiliar = 0
    esFilaLlenada = False
    tarimasJuntar = 0
    
    'Llenado de las reubicaciones de tarimas
    While registros_noLlenos > tarimasReubicadas
        
        contador = 0
        auxiliar = 0
        hayfilaVacia = True
        ubicacionCambiar = ""
        valorCelda = ""
        If tarimasJuntar = 20 Then
            tarimasJuntar = 0
        End If
        
        If tarimasJuntar = 0 Then
            While contador < registros_noLlenos 'busca mayor tarima
                If tarim_noLlenas(contador) > auxiliar Then
                    auxiliar = tarim_noLlenas(contador)
                    indiceArregloMayorTarima = contador
                End If
                contador = contador + 1
            Wend
            tarimasJuntar = tarim_noLlenas(indiceArregloMayorTarima)
            ubicacionMayorTarima = ubicac_noLlenas(indiceArregloMayorTarima)
            tarim_noLlenas(indiceArregloMayorTarima) = 0
        End If
        
        contador = 0
        rangoMaxDdv = 10
        tarimaMax = tarimaMaxima
        Do While contador < registros_noLlenos 'busca tarima que iguale las 20 tarimas
            If contador <> indiceArregloMayorTarima Then
                If Abs(ddv_masBajo(contador) - ddv_masBajo(indiceArregloMayorTarima)) <= rangoMaxDdv Then
                    If tarimasJuntar + tarim_noLlenas(contador) = tarimaMax And tarim_noLlenas(contador) <> 0 Then
                        tarimasJuntar = tarimasJuntar + tarim_noLlenas(contador)
                        tarim_noLlenas(contador) = 0
                        ubicacionCambiar = ubicac_noLlenas(contador)
                        Exit Do
                    End If
                End If
            End If
            contador = contador + 1
            
            If contador = registros_noLlenos Then
                contador = 0
                rangoMaxDdv = rangoMaxDdv + 10
                If rangoMaxDdv > 30 Then
                    tarimaMax = tarimaMax - 1
                    rangoMaxDdv = 10
                    If tarimaMax <= tarimasJuntar Then
                        valorCelda = "NoReubicable"
                        Exit Do
                    End If
                End If
            End If
        Loop
        
        contador = 0
        If valorCelda = "" Then
            While contador < registros_noLlenos
                valorCelda = ""
                
                If Worksheets(cambiarUbicaciones).Cells(ubicacionRegistros(contador), 2) = ubicacionMayorTarima Then
                    valorCelda = ubicacionCambiar
                ElseIf Worksheets(cambiarUbicaciones).Cells(ubicacionRegistros(contador), 2) = ubicacionCambiar Then
                    valorCelda = "REUBICANDO"
                End If
                
                If (valorCelda <> "") Then
                    tarimasReubicadas = tarimasReubicadas + 1
                    If Worksheets(cambiarUbicaciones).Cells(ubicacionRegistros(contador), 6).Value = "" Then
                        Worksheets(cambiarUbicaciones).Cells(ubicacionRegistros(contador), 6).Value = valorCelda
                    Else
                        valorCelda = Worksheets(cambiarUbicaciones).Cells(ubicacionRegistros(contador), 6).Value & "," & valorCelda
                        Worksheets(cambiarUbicaciones).Cells(ubicacionRegistros(contador), 6).Value = valorCelda
                    End If
                End If
                contador = contador + 1
            Wend
        Else
            If esFilaLlenada = True Then
                tarimasJuntar = 0
            Else
                While contador < registros_noLlenos
                    If Worksheets(cambiarUbicaciones).Cells(ubicacionRegistros(contador), 2) = ubicacionMayorTarima Then
                        valorCelda = "No Reubicable"
                        Worksheets(cambiarUbicaciones).Cells(ubicacionRegistros(contador), 6).Value = valorCelda
                        tarimasReubicadas = tarimasReubicadas + 1
                        tarimasJuntar = 0
                    End If
                    contador = contador + 1
                Wend
            End If
        End If
        
        If tarimasJuntar > 0 And tarimasJuntar < 20 Then
            esFilaLlenada = True
        Else
            esFilaLlenada = False
        End If
    Wend
     
    Worksheets(cambiarUbicaciones).Range("A1:F1").AutoFilter       'LimpiarFiltros
End Function
