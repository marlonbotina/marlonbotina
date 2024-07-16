Private Sub Comando5_Click()
    Dim ctl As Control
    
    ' Recorrer todos los controles en el formulario
    For Each ctl In Me.Controls
        ' Verificar el tipo de cada control y realizar la acción correspondiente
        If TypeOf ctl Is TextBox Then
            ' Para cuadros de texto, establecer el valor como una cadena vacía
            ctl.Value = ""
        
        ElseIf TypeOf ctl Is ComboBox Then
            ' Para cuadros combinados, establecer el valor como NullString
            ctl.Value = NullString
        
        ElseIf TypeOf ctl Is ListBox Then
            ' Para listas, limpiar el origen de fila (RowSource)
            ctl.RowSource = ""
        
        ElseIf TypeOf ctl Is CheckBox Then
            ' Para casillas de verificación, desmarcar estableciendo el valor como False
            ctl.Value = False
        End If
    Next ctl
End Sub
