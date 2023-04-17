Sub hide()
    Application.ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"", False)"
    Application.DisplayFormulaBar = False
    Application.DisplayStatusBar = Not Application.DisplayStatusBar
    ActiveWindow.DisplayWorkbookTabs = False
    ActiveWindow.DisplayHeadings = False
End Sub

Sub Auto_open()
    Application.ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"", False)"
    Application.DisplayFormulaBar = False
    Application.DisplayStatusBar = Not Application.DisplayStatusBar
    ActiveWindow.DisplayWorkbookTabs = False
    ActiveWindow.DisplayHeadings = False
End Sub

Sub Unhide()
    Application.ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"", True)"
    Application.DisplayFormulaBar = True
    Application.DisplayStatusBar = True
    ActiveWindow.DisplayWorkbookTabs = True
    ActiveWindow.DisplayHeadings = True
End Sub


Sub dinamica_cantidad_estrategia()

Sheets("DN_CANT").Select

End Sub

Sub dinamica_cantidad_RESIDUO()

Sheets("DN_RESI_GENER").Select

End Sub


Sub dinamica_RESIUDOS()

Sheets("RESIDUOS").Select

End Sub
Sub IR_APROVECHAMIENTO()

Sheets("aprovechamiento").Select

End Sub
Sub IR_REUTILIZACION()

Sheets("reutilizacion").Select

End Sub


Sub IR_REDUCCION()

Sheets("% REDUCCION").Select

End Sub


Sub IR_REDUCCION2()

Sheets("REDUCCION 2").Select

End Sub
