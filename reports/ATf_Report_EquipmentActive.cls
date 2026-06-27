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
If Len(TotalOEC) > 0 Then
Call MakeVertLine(4, Me.VendorOec.Top, Me.VendorOec.Left - 115, (Me.VendorOec.Width + Me.TotalOEC.Width) + 40)
End If
End Sub

Private Sub GroupFooter1_Print(Cancel As Integer, PrintCount As Integer)
If Me.GroupsIn > 1 Then
Call MakeVertLine(2, Me.Units.Top, Me.Units.Left, (1440 * 8))
'Call MakeVertLine(4, Me.Units.Height + 50, 0, (1440 * 8))
End If
End Sub

Private Sub GroupHeader0_Print(Cancel As Integer, PrintCount As Integer)
'UnitDesc Header
'Me.UnitCnt
'Call MakeVertLine(Me.GroupHeader0.Height)
End Sub

Private Sub GroupHeader1_Print(Cancel As Integer, PrintCount As Integer)
'ChildPartDesc Header
'ChildDescLine
'Call MakeVertLine(Me.ChildDescLine.Height + 20)
End Sub

Private Sub GroupHeader2_Print(Cancel As Integer, PrintCount As Integer)
'UnitGroup Header
'Call MakeVertLine(Me.GroupHeader2.Height)
'Call MakeVertLine(100, 90, 170, 100, True)
'Call MakeVertLine(100, 90, 85, 100)
End Sub
