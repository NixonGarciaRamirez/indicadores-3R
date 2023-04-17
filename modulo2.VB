Sub REGISTRO_MENU()
'ESTO ES PARA MOSTRAR EL REGISTRO O LA CONSOLA
Hoja7.Shapes("9 Grupo").Visible = msoFalse
REGISTRO.Show

End Sub

Sub LISTA_MENU()
'CON ESTE CODIGO MOSTRAMOS UNA LISTA PROVISIONAL PARA VER LOS DATOS EN LA BASE DE DATOS
Hoja7.Shapes("9 Grupo").Visible = msoFalse
LISTA.Show

End Sub


Sub MOSTRAR_MENU1()
'CON ESTE CODIGO PODEMOS MOSTRAR Y OCULTAR EL MENU
If Hoja7.Shapes("9 Grupo").Visible = True Then
    Hoja7.Shapes("9 Grupo").Visible = msoFalse
    Else
    Hoja7.Shapes("9 Grupo").Visible = msoCTrue
    End If



End Sub

Sub MOSTRAR_MENU2()
'CON ESTE CODIGO PODEMOS MOSTRAR Y OCULTAR EL MENU
If Hoja7.Shapes("15 Grupo").Visible = True Then
    Hoja7.Shapes("15 Grupo").Visible = msoFalse
    Else
    Hoja7.Shapes("15 Grupo").Visible = msoCTrue
    End If



End Sub

Sub MOSTRAR_MENU3()
'CON ESTE CODIGO PODEMOS MOSTRAR Y OCULTAR EL MENU
If Hoja7.Shapes("25 Grupo").Visible = True Then
    Hoja7.Shapes("25 Grupo").Visible = msoFalse
    Else
    Hoja7.Shapes("25 Grupo").Visible = msoCTrue
    End If



End Sub

Sub modelr()
   REGISTRO.lb1_info.Caption = "registro"
   
End Sub

Sub modeE()
    REGISTRO.lb1_info.Caption = "EDICION"
End Sub
