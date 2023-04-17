



Private Sub ACTUALIZAR_Click()


Hoja11.PivotTables("R_APROVECHAMIENTO").PivotCache.Refresh
Hoja11.PivotTables("R_REUTILIZADOS").PivotCache.Refresh
Hoja11.PivotTables("R_REDUCIDOS").PivotCache.Refresh



End Sub



Private Sub GUARDAR_Click()

If MES.Value = Empty Or RESIDUO.Value = Empty Or INICIATIVA.Value = Empty Or TIPO.Value = Empty Or CANTIDAD.Value = Empty Then
MsgBox ("DATOS VACIOS")
Exit Sub
End If

If lb1_info = "edicion" Then
    Call ACUATLIZAR
Exit Sub
End If


endd = Hoja9.Cells(Rows.Count, "A").End(xlUp).Row + 1


With Hoja9
        .Range("A" & endd) = MES
        .Range("B" & endd) = "Cartagena"
        .Range("C" & endd) = RESIDUO
        .Range("E" & endd) = INICIATIVA
        .Range("F" & endd) = TIPO
        .Range("G" & endd) = CANTIDAD
    
End With

'ESTE ODIGO SOLUCIONA EL PROBLEMA DE QUE LOS NUMEROS NOS SALGAN COMO TEXTO, EN RESUMEN LO UNICO QUE HACEMOS ES SUMARLE Y RESTARLE 1
If Not IsNumeric(Hoja9.Range("G" & endd).Value) Then
    Hoja9.Range("G" & endd).Value = WorksheetFunction.Trim(Hoja9.Range("G" & endd).Value)
ElseIf TypeName(Hoja9.Range("G" & endd).Value) = "String" Then
    Hoja9.Range("G" & endd).Value = Hoja9.Range("G" & endd).Value + 1 - 1
End If


If separacion_en_mesa.Value = True Then
    Hoja9.Range("H" & endd) = "Caracterizacion de ordinarios"
Else
    Hoja9.Range("H" & endd) = "Remisiones gestores ambientales"
     
End If




Call LIMPIAR


End Sub



Private Sub SALILR_Click()


Unload Me


End Sub



Private Sub UserForm_Activate()
'LISTA RESIDUO
Dim sh As Worksheet
    Set sh = ThisWorkbook.Sheets("CLASIFICACION CTG")
    
Dim i As Integer
    For i = 2 To sh.Range("A" & Rows.Count).End(xlUp).Row
        Me.RESIDUO.AddItem sh.Range("A" & i)
    Next i
    
 'LISTA MES
Dim ah As Worksheet
    Set ah = ThisWorkbook.Sheets("LISTAS")
    
Dim F As Integer
    For F = 2 To ah.Range("D" & Rows.Count).End(xlUp).Row
        Me.MES.AddItem ah.Range("D" & F)
    Next F
    

  'LISTA INICIATIVA
Dim dh As Worksheet
    Set dh = ThisWorkbook.Sheets("LISTAS")
    
Dim G As Integer
    For G = 2 To ah.Range("B" & Rows.Count).End(xlUp).Row
        Me.INICIATIVA.AddItem ah.Range("B" & G)
    Next G
    
    
  ' LISTA TIPO
  
Dim Fh As Worksheet
    Set Fh = ThisWorkbook.Sheets("LISTAS")
    
Dim R As Integer
    For R = 2 To ah.Range("C" & Rows.Count).End(xlUp).Row
        Me.TIPO.AddItem ah.Range("C" & R)
    Next R
        
End Sub

Sub LIMPIAR()
Dim objeto As Control
For Each objeto In Me.Controls
    If TypeName(objeto) = "TextBox" Then
    objeto.Text = ""
    End If
Next objeto

Dim com As Object
For Each com In Me.Controls
If TypeOf com Is ComboBox Then
com = ""
End If
Next com

End Sub


Sub ACUATLIZAR()
Dim fila As Integer
    With Hoja9
        fila = .Range("A:A").Find(MES).Row
        .Cells(2, 1) = MES
        .Cells(2, 2) = "CARTAGENA"
        .Cells(2, 3) = RESIDUO
        .Cells(2, 4) = INICIATIVA
        .Cells(2, 5) = TIPO
        .Cells(2, 6) = CANTIDAD
     End With
ActiveWorkbook.Save
MsgBox ("CAMBIO REALIZADO CON EXITO"), vbInformation
Call LIMPIAR
Call SALILR_Click
Call LISTA.mostrar

End Sub




