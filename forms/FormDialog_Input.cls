Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Open(Cancel As Integer)
    Dim NumArgs, ii As Integer
    Dim StrArg As String
    'PromptUser (Prompt,Style,Title,Input,button1,button2,button3)
    If Len(Me.OpenArgs) > 0 Then NumArgs = InStr(Me.OpenArgs, "|")

    For ii = NumArgs To 1 Step -1
        StrArg = SplitString(Me.OpenArgs, "|", ii - 1)
    Select Case ii

        Case Is = 7
            Me.bt3.Visible = False
      '      Me.bt3.caption = StrArg
        Case Is = 6
            Me.bt2.Visible = False
      '      Me.bt2.caption = StrArg
        Case Is = 5
            Me.bt1.Visible = True
            Me.bt1.caption = StrArg
        Case Is = 4
        Me.InputText.Visible = True
        Me.InputLabel.caption = StrArg
        Case Is = 3
        Me.Form.caption = StrArg
        Case Is = 2
            Select Case StrArg
            Case "Critical"
                Me.imgCritical.Visible = True
            Case "Question"
                Me.imgQuestion.Visible = True
            Case "Information"
                Me.imgInformation.Visible = True
            Case "Warning"
                Me.imgExclamation.Visible = True
            End Select
        Case Is = 1
           Me.Message = StrArg
        End Select
        Next ii
End Sub
' msgbox  Prompt,Style-buttons,Title,help,Context as VbMsgBoxResult

'PromptUser (Prompt,Style,Title,Input,button1,button2,button3)
