Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Public Sub MakeVertLine(LineHeight As Double _
                         , Optional LineTop As Double _
                         , Optional LineLeft As Double _
                         , Optional LineWidth As Double _
                         , Optional MakeFrame As Boolean _
                         , Optional Rgb1 As Integer _
                         , Optional Rgb2 As Integer _
                         , Optional Rgb3 As Integer)
    Const acbcSMTwips = 1
    Const acbcDSSolid = 0
    Me.ScaleMode = acbcSMTwips
    Me.DrawStyle = acbcDSSolid
    ' Set the coordinates for the line.
    If Nz(LineTop, 0) = 0 Then LineTop = 0
    If Nz(LineLeft, 0) = 0 Then LineLeft = 200
    If Nz(LineWidth, 0) = 0 Then LineWidth = 4
    If Nz(Rgb1, 0) = 0 Then Rgb1 = 127
    If Nz(Rgb2, 0) = 0 Then Rgb2 = 127
    If Nz(Rgb3, 0) = 0 Then Rgb3 = 127

    ' Draw the line.
        If MakeFrame = True Then
    Me.Line (LineLeft, LineTop)-Step(LineWidth, LineHeight), RGB(Rgb1, Rgb2, Rgb3), B
    Else
    Me.Line (LineLeft, LineTop)-Step(LineWidth, LineHeight), RGB(Rgb1, Rgb2, Rgb3), BF
    End If
    'BF will make a filled box
End Sub

Private Sub GroupHeader0_Print(Cancel As Integer, PrintCount As Integer)
'UnitDesc Header
'Me.UnitCnt
'Call MakeVertLine(Me.GroupHeader0.Height)
End Sub

Private Sub Detail_Print(Cancel As Integer, PrintCount As Integer)
Call MakeVertLine(80, 50, 750, 50, False, 0, 0, 0)
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
