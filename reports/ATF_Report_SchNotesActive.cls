Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Sub MakeVertLine(LineHeight As Double, _
                         Optional LineTop As Double, _
                         Optional LineLeft As Double, _
                         Optional LineWidth As Double, _
                         Optional MakeFrame As Boolean)
    Const acbcSMTwips = 1
    Const acbcDSSolid = 0
    Me.ScaleMode = acbcSMTwips
    Me.DrawStyle = acbcDSSolid
    ' Set the coordinates for the line.
    If Nz(LineTop, 0) = 0 Then LineTop = 0
    If Nz(LineLeft, 0) = 0 Then LineLeft = 200
    If Nz(LineWidth, 0) = 0 Then LineWidth = 4
    ' Draw the line.
    If MakeFrame = True Then
    Me.Line (LineLeft, LineTop)-Step(LineWidth, LineHeight), , B
    Else
    Me.Line (LineLeft, LineTop)-Step(LineWidth, LineHeight), , BF
    End If
    'BF will make a filled box
End Sub

Private Sub Report_Load()
'Me.ActionDesc.Height = Me.RegText.Height
End Sub
