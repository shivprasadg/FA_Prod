Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
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



Private Sub GroupFooter0_Print(Cancel As Integer, PrintCount As Integer)
'Call MakeVertLine(Me.TotalOEC.Height)
'If Len(TotalOEC) > 0 Then
'Call MakeVertLine(4, 0, Me.TotalOEC.Left - 125, Me.TotalOEC.Width + 50)
'Call MakeVertLine(4, Me.TotalOEC.Top, Me.TotalOEC.Left - 125, Me.TotalOEC.Width + 50)
'End If
End Sub



Private Sub GroupFooter1_Print(Cancel As Integer, PrintCount As Integer)
Call MakeVertLine(2, (Me.Units.Top - 40), 20, (1440 * 8) - 40)
Call MakeVertLine(2, Me.Units.Top, 20, (1440 * 8) - 40)
End Sub
