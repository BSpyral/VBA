Attribute VB_Name = "Módulo2"
Sub Borrar_datos()
Attribute Borrar_datos.VB_ProcData.VB_Invoke_Func = " \n14"
    
    Dim compra_sap, program_ruta, resumen, info_tarimas As String        'Nombre de las hojas del libro
    
    compra_sap = "Emb SAP"
    program_ruta = "Programa"
    resumen = "Resumen"
    info_tarimas = "WMS"
    
    ultima_fila = Worksheets(program_ruta).Cells(Rows.Count, 1).End(xlUp).Row        'Ubica la ultima fila
    rango_eliminar = "A2:U" + CStr(ultima_fila)
    If ultima_fila > 1 Then                            'Garantiza que no borre los titulos de las columnas
    Worksheets(program_ruta).Range(rango_eliminar).ClearContents        'Elimina datos de la hoja "Programa", en este caso
    End If
    
    ultima_fila = Worksheets(compra_sap).Cells(Rows.Count, 4).End(xlUp).Row         'Ubica la ultima fila
    rango_eliminar = "A2:J" + CStr(ultima_fila) + ",M2:AM" + CStr(ultima_fila)
    If ultima_fila > 1 Then
    Worksheets(compra_sap).Range(rango_eliminar).ClearContents
    End If
    
    ultima_fila = Worksheets(info_tarimas).Cells(Rows.Count, 6).End(xlUp).Row        'Ubica la ultima fila
    rango_eliminar = "A2:S" + CStr(ultima_fila)
    If ultima_fila > 1 Then
        Worksheets(info_tarimas).Range(rango_eliminar).ClearContents
    End If
    
    ultima_fila = Worksheets(resumen).Cells(Rows.Count, 4).End(xlUp).Row        'Ubica la ultima fila
    rango_eliminar = "A2:AM" + CStr(ultima_fila)
    If ultima_fila > 1 Then
        Worksheets(resumen).Range(rango_eliminar).ClearContents
    End If
        
    MsgBox ("Datos borrados correctamente")
End Sub
