




Private Sub BD1_Click()

End Sub

Private Sub CommandButton1_Click()
'con este codigo comenzamos la edision deleditor de nuestra tabla
If BD1.ListIndex = -1 Then
   MsgBox ("SELECCIONE DAOS POR FAVOR"), vdExclamation
   Exit Sub
   End If
Call cargar_campos
Call modeE
Call REGISTRO_MENU

End Sub

Private Sub CommandButton2_Click()
'eliminar columna
Dim fila As Integer

If BD1.ListIndex = -1 Then
       MsgBox ("seleccione datos por favor"), vbExclamation
       Exit Sub
End If


Me.ID = Me.BD1.Column(0)

Dim mensaje As String
mensaje = MsgBox("El registro sera eliminado, corfirma la eliminacion?", vbYesNo)
If emnsaje = vbsi Then
With Hoja9
    fila = .Range("A:A").Find(ID).EntireRow.Delete
End With
    
Call mostrar

MsgBox ("REGISTRO ELIMINADO CON EXITO"), vbInformation
Else

End If


End Sub

Private Sub CommandButton3_Click()
'para salir de la ventana
Unload Me
End Sub



Private Sub CommandButton4_Click()

Sheets("BD").Select


End Sub

Private Sub Label1_Click()

End Sub

Private Sub UserForm_Activate()
'con este codigo llamamos a nuestra lista o base de datos
Call mostrar
End Sub


Private Sub UserForm_Initialize()
''aqui creamos nuestra tabla o base de datos
With BD1

     .ColumnCount = 7
     .ColumnWidths = "100 PT ; 100 PT ; 200 PT ; 100 PT ; 200 PT ; 100 PT ; 100 PT ;"
     
End With

End Sub

Sub mostrar()
'aqui seleccionamos los datos que rellenaran nuestra base de datos
BD1.List = Sheets("BD").Range("A2:G" & Sheets("BD").Cells(Rows.Count, 1).End(xlUp).Row).Value


End Sub

Private Sub cargar_campos()
'con esta funcion llamaremos a un editor de campos de nuestra base de datos
'CABE RESALTAR QUE LOS CAMPOS A EDITAR SOLO PUEDEN SER LOS MISMOS QUE LOS QUE
'NOSOTROS INGRESAMOS AL REGUISTRO
S1 = Me.BD1.List(BD1.ListIndex, 0)
S3 = Me.BD1.List(BD1.ListIndex, 2)
S4 = Me.BD1.List(BD1.ListIndex, 4)
S5 = Me.BD1.List(BD1.ListIndex, 5)
S6 = Me.BD1.List(BD1.ListIndex, 6)


REGISTRO.MES = S1
REGISTRO.RESIDUO = S3
REGISTRO.INICIATIVA = S4
REGISTRO.TIPO = S5
REGISTRO.CANTIDAD = S6
  

End Sub