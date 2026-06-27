Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Detail_Print(Cancel As Integer, PrintCount As Integer)
If Len(CONDESC) > 0 Then
Call MakeVertLine(CONDESC.Height, 0, Me.[Type].Left + 50, 2, False, 0, 0, 0)
Call MakeVertLine(CONDESC.Height, 0, Me.[clientCompanyName].Left + 10, 2, False, 0, 114, 188)
Call MakeVertLine(CONDESC.Height, 0, Me.[conYear].Left + 10, 2, False, 117, 140, 72)
End If
End Sub

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
