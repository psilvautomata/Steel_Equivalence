Sub analise_Comp()

Application.ScreenUpdating = False 'Disable screen updating
Application.EnableEvents = False 'Disable excel events


Workbooks.Open Filename:="...\BD_Certificados.xlsm" 'Variable declarations

Workbooks("Atender Material").Activate 'Activate Excel Workbook

max = Worksheets("Análise de Composição").Range("V1").Value

Range("C8:U81").Select
Selection.ClearContents

For i = 8 To max + 7

    varLote = Worksheets("Análise de Composição").Range("B" & i).Value
    
    Workbooks("BD_Certificados.xlsm").Activate
    Worksheets("Dados_Galv").Activate
    Worksheets("Dados_Galv").Range("A2").Value = varLote
    Worksheets("Dados_Galv").Range("B2:O2").Select
    Selection.Copy
    
    'Copies core values from DB Workbook
    Mat = Worksheets("Dados_Galv").Range("S2").Value
    LE = Worksheets("Dados_Galv").Range("Q2").Value
    LR = Worksheets("Dados_Galv").Range("R2").Value
    Along = Worksheets("Dados_Galv").Range("P2").Value
    Acab = Worksheets("Dados_Galv").Range("T2").Value
    
    Workbooks("Atender Material.xlsm").Activate
    Worksheets("Análise de Composição").Range("F" & i).Select
    
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    'Pastes core values from DB Workbook
    Worksheets("Análise de Composição").Range("C" & i).Value = Along
    Worksheets("Análise de Composição").Range("D" & i).Value = LE
    Worksheets("Análise de Composição").Range("E" & i).Value = LR
    Worksheets("Análise de Composição").Range("T" & i).Value = Acab
    Worksheets("Análise de Composição").Range("U" & i).Value = Mat
    
Next

Workbooks("BD_Certificados.xlsm").Close SaveChanges:=False 'Close BD workbook withou saving

Application.ScreenUpdating = True 'Enable screen updating
Application.EnableEvents = True 'Enable excel events

MsgBox ("Dados importados com sucesso!")

Range("B8").Select

End Sub
