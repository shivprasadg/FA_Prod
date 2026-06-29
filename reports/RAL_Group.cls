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
Private Sub GroupHeader0_Print(Cancel As Integer, PrintCount As Integer)

Call MakeVertLine(Me.GroupHeaderLine.Height, Me.GroupHeaderLine.Top, Me.GroupHeaderLine.Left, GroupHeaderLine.Width, True)

Call MakeVertLine(Me.Desc.Height, Me.Desc.Top, Me.Desc.Left, Me.Desc.Width, True)
Call MakeVertLine(Me.hereto.Height, Me.hereto.Top, Me.hereto.Left, Me.hereto.Width, True)
Call MakeVertLine(Me.Lessorcostx.Height, Me.Lessorcostx.Top, Me.Lessorcostx.Left, Me.Lessorcostx.Width, True)
Call MakeVertLine(Me.SyndRent.Height, Me.SyndRent.Top, Me.SyndRent.Left, Me.SyndRent.Width, True)
Call MakeVertLine(Me.BLCD.Height, Me.BLCD.Top, Me.BLCD.Left, Me.BLCD.Width, True)
Call MakeVertLine(Me.LeaseTerm.Height, Me.LeaseTerm.Top, Me.LeaseTerm.Left, Me.LeaseTerm.Width, True)
Call MakeVertLine(Me.AssignedRents.Height, Me.AssignedRents.Top, Me.AssignedRents.Left, Me.AssignedRents.Width, True)

Call MakeVertLine(Me.Desc.Height, Me.Desc.Top, Me.DescLabel.Left, Me.DescLabel.Width, True)
Call MakeVertLine(Me.hereto.Height, Me.hereto.Top, Me.HereToLabel.Left, Me.HereToLabel.Width, True)
Call MakeVertLine(Me.Lessorcostx.Height, Me.Lessorcostx.Top, Me.LessorCostLabel.Left, Me.LessorCostLabel.Width, True)
Call MakeVertLine(Me.SyndRent.Height, Me.SyndRent.Top, Me.SyndRentLabel.Left, Me.SyndRentLabel.Width, True)
Call MakeVertLine(Me.BLCD.Height, Me.BLCD.Top, Me.BLCDLabel.Left, Me.BLCDLabel.Width, True)
Call MakeVertLine(Me.LeaseTerm.Height, Me.LeaseTerm.Top, Me.LeaseTermLabel.Left, Me.LeaseTermLabel.Width, True)
Call MakeVertLine(Me.AssignedRents.Height, Me.AssignedRents.Top, Me.AssignedLabel.Left, Me.AssignedLabel.Width, True)
End Sub
