Sub cleanData()
'
' Apagar informações
'

'
Dim i As Variant

Application.ScreenUpdating = False
Application.EnableEvents = False

Range("B11:E23").Select
Selection.ClearContents
Range("H11:K23").Select
Selection.ClearContents
Range("B11").Select

Sheets("Análise de Composição").Activate

Range("B8:U81").Select
Selection.ClearContents
Range("B8").Select

Sheets("CV 300-345 STi").Activate
    
For i = 11 To 23

    Range("B" & i).Select
    ActiveCell.FormulaR1C1 = _
        "=IF('Análise de Composição'!R[-3]C="""","""",'Análise de Composição'!R[-3]C)"
    Range("C" & i).Select
    ActiveCell.FormulaR1C1 = _
        "=IF('Análise de Composição'!R[-3]C[-1]="""","""",(XLOOKUP(RC[-1],'Análise de Composição'!C[-1],'Análise de Composição'!C)))"
    Range("D" & i).Select
    ActiveCell.FormulaR1C1 = _
        "=IF('Análise de Composição'!R[-3]C[-2]="""","""",(XLOOKUP(RC[-2],'Análise de Composição'!C[-2],'Análise de Composição'!C)))"
    Range("E" & i).Select
    ActiveCell.FormulaR1C1 = _
        "=IF('Análise de Composição'!R[-3]C[-3]="""","""",(XLOOKUP(RC[-3],'Análise de Composição'!C[-3],'Análise de Composição'!C)))"
    Range("H" & i).Select
    ActiveCell.FormulaR1C1 = _
        "=IF('Análise de Composição'!R[-3]C[-6]="""","""",(XLOOKUP(RC[-6],'Análise de Composição'!C[-6],'Análise de Composição'!C[12])))"
    Range("I" & i).Select
    ActiveCell.FormulaR1C1 = _
        "=IF('Análise de Composição'!R[-3]C[-7]="""","""",(XLOOKUP(RC[-7],'Análise de Composição'!C[-7],'Análise de Composição'!C[12])))"
    Range("J" & i).Select
    ActiveCell.FormulaR1C1 = _
        "=IF('Análise de Composição'!R[-3]C[-8]="""","""",(XLOOKUP(RC[-8],'Análise de Composição'!C[-8],'Análise de Composição'!C[-3])))"
    Range("K" & i).Select
    ActiveCell.FormulaR1C1 = _
        "=IF('Análise de Composição'!R[-3]C[-9]="""","""",(XLOOKUP(RC[-9],'Análise de Composição'!C[-9],'Análise de Composição'!C[-2])))"
    
Next

Application.ScreenUpdating = True
Application.EnableEvents = True
    
End Sub
