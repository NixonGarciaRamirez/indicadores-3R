Sub IR_BD()
'
' IR_BD Macro
' esta macros lo que hace es que nos dirige a la hoja de base de datos
'

'
    Sheets("BD").Select
End Sub

Sub grabar()


If Range("B7").Value = Empty Or Range("B9").Value = Empty Or Range("E7").Value = Empty Or Range("E9").Value = Empty Or Range("E11").Value < 0 Or Range("B11").Value < 0 Then
MsgBox ("DATOS VACIOS")
Exit Sub
End If



Sheets("BD").Select
Range("A2").EntireRow.Insert
Sheets("REGISTRO").Select




Range("C7").Copy
Sheets("BD").Select
Range("C2").PasteSpecial (xlPasteValues)
Sheets("REGISTRO").Select

Range("C9").Copy
Sheets("BD").Select
Range("A2").PasteSpecial (xlPasteValues)
Sheets("REGISTRO").Select

Range("F7").Copy
Sheets("BD").Select
Range("D2").PasteSpecial (xlPasteValues)
Sheets("REGISTRO").Select

Range("F9").Copy
Sheets("BD").Select
Range("E2").PasteSpecial (xlPasteValues)
Sheets("REGISTRO").Select

Range("F11").Copy
Sheets("BD").Select
Range("L2").PasteSpecial (xlPasteValues)
Sheets("REGISTRO").Select


Range("C11").Copy
Sheets("BD").Select
Range("B2").PasteSpecial (xlPasteValues)
Sheets("REGISTRO").Select

If Range("F9") = "ORGANICO" Then
     Dim a As String
     Dim b As Integer
     a = InputBox("escribe el nuemro de platos servidos")
     b = Val(a)
     Hoja9.Cells(2, 9) = b
     MsgBox ("proceso terminado")
     
Else
     Hoja9.Cells(2, 9) = "NA"
     MsgBox ("proceso terminado")
End If

End Sub



Sub LIMPIAR()

Range("C7").Value = Empty
Range("C9").Value = Empty

Range("c11").Value = Empty
Range("F7").Value = Empty

Range("F9").Value = Empty
Range("F11").Value = Empty




Range("J7").Value = Empty
Range("J9").Value = Empty

Range("L8").Value = Empty
Range("J11").Value = Empty





End Sub



Sub eliminar()

Sheets("BD").Select
Range("A2").EntireRow.Delete
Sheets("REGISTRO").Select




End Sub


Sub volver()
Sheets("MENU").Select

End Sub

Sub TABLADINAMICA()
Sheets("TD").Select

End Sub


Sub eliminar2()
Sheets("BD").Select
Range("A2").EntireRow.Delete

End Sub

Sub aprovechamiento()

Sheets("APROVECHAMIENTO").Select


End Sub

Sub reutilizacion()
Sheets("REUTILIZACION").Select

End Sub

Sub reduccion()

Sheets("REDUCCION").Select

End Sub

Sub GRABAR2()

Sheets("BD (PM)").Select
Range("A2").EntireRow.Insert
Sheets("REGISTRO").Select


Range("J7").Copy
Sheets("BD (PM)").Select
Range("B2").PasteSpecial (xlPasteValues)
Sheets("REGISTRO").Select

Range("J9").Copy
Sheets("BD (PM)").Select
Range("C2").PasteSpecial (xlPasteValues)
Sheets("REGISTRO").Select


Range("J11").Copy
Sheets("BD (PM)").Select
Range("D2").PasteSpecial (xlPasteValues)
Sheets("REGISTRO").Select

Range("L8").Copy
Sheets("BD (PM)").Select
Range("A2").PasteSpecial (xlPasteValues)
Sheets("REGISTRO").Select

End Sub

Sub ir_bd2()
Sheets("BD (PM)").Select
End Sub

Sub eliminar3()

Range("A2").EntireRow.Delete
End Sub

Sub IR_RECI()
Sheets("RECICLABLES").Select
End Sub

Sub IR_ORDI()
Sheets("ORDINARIOS").Select
End Sub
