VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Planilha2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    ' Só ativa se clicar nas células A2 ou A3
    If Target.Cells.Count > 1 Then Exit Sub
    If Intersect(Target, Range("A2:A3")) Is Nothing Then Exit Sub
    
    Application.ScreenUpdating = False
    
    ' Pinta o fundo das células do menu
    Range("A2:A3").Interior.Color = RGB(15, 108, 189)
    Target.Interior.Color = RGB(0, 70, 140)
    
    ' Esconde os dois blocos usando os seus nomes novos
    On Error Resume Next
    Me.Shapes("BotAgenda").Visible = msoFalse
    Me.Shapes("BotFinaliza").Visible = msoFalse
    On Error GoTo 0
    
    ' Mostra apenas o bloco correto
    If Target.Address = "$A$2" Then
        Me.Shapes("BotAgenda").Visible = msoTrue
    ElseIf Target.Address = "$A$3" Then
        Me.Shapes("BotFinaliza").Visible = msoTrue
    End If
    
    Application.ScreenUpdating = True
End Sub

