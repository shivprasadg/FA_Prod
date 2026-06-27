Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Dim SQLz As String
Private Sub Report_Load()
'If Me.clientID = 4209 Then Me.FeeLabel.Caption = "Ancillary"
End Sub

Public Sub Report_Open(Cancel As Integer)

If Not IsNull(Me.OpenArgs) Then
SQLz = "Select * From ExitReportBank " & Me.OpenArgs & ""
Else
SQLz = "Select * From ExitReportBank;"
End If
Me.Report.RecordSource = SQLz
End Sub

Private Sub Send2Excel_Click()
Dim XLS, Headerx As String
Headerx = "EXIT Report " & Me.UnitsShown & " Units " & Format(Date, "mm-dd-yy")
If SQLz Like "*Sale*" Then Headerx = Headerx & " BDEX " & Me.BDEX
'_________________________
'XLS = Col letter,Format Option
' first set = number of columns
'Format Options Type Command
'Types:     T = Text | B = Bold | $ = Currency
'Commands:  C = Center | L=Lexf | R=Right | E = Sum col | T= Total Col Count | X = COUNT

'XLS = "|C,TX|Q,$E|R,$E|S,$E|T,$E|U,$E|V,$E|W,$E|Z,$E|AE,$E|AI,$E|AJ,$E|AK,$E|AL,$E|AM,$E|AN,$E|AO,$E|AP,$E|AQ,$E|AR,$E|AS,$E|AT,$E|AU,$E|AV,$E|AW,$E|AX,$E|AY,$E|AZ,$E|BA,$E|BB,$E|BC,$E|BD,$E|BE,$E|BF,$E|BF,$E|BG,$E|BH,$E|BI,$E|"
XLS = "|G,BC|L,$E|N,BC|N,BC|O,$E|U,BC|V,BC|AC,BC|"
SendReport2Excel SQLz, "EXIT Report Units", Headerx, True, XLS, Me.UnitsShown

End Sub
