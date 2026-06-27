Attribute VB_Name = "FAModule1"
Option Compare Database
Public Function getValueFromPopUp(formName As String, PopUpControlName As String, Optional MyOpenArgs As String = "None") As Variant
  'FormName: Name of the popup form
  'PopupControlName: Name of the control on the pop up/dialog that you want the value
  'MyOpenArgs: If you want to pass something to your pop up.
  Dim frm As Access.Form
  Dim ctrl As Access.Control
  'Code execution stops after next line of code
  DoCmd.close acForm, formName
  DoCmd.OpenForm formName, , , , acFormEdit, acDialog, MyOpenArgs
  'Wait until form is closed or hidden
  'The popup needs an OK button that hides the popup(me.visible = false), and a Cancel button that just closes it
  If CurrentProject.AllForms(formName).IsLoaded Then
    Set frm = Forms(formName)
    Set ctrl = frm.Controls(PopUpControlName)
    If ctrl.ControlType = acLabel Then
      getValueFromPopUp = frm.Controls(PopUpControlName).caption
    Else
      getValueFromPopUp = frm.Controls(PopUpControlName).Value
    End If
    DoCmd.close acForm, frm.Name
  End If
End Function

Public Function SetAdobePDFAsDefault()
    Dim objWMIService As Object
    Dim colPrinters As Object
    Dim objPrinter As Object
    Dim strPrinterName As String
    strPrinterName = "Microsoft Print to PDF"
  '  strPrinterName = "Adobe PDF"
    ' Connect to WMI
    Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
    Set colPrinters = objWMIService.ExecQuery("Select * From Win32_Printer")

    ' Loop through printers and set Adobe PDF as default
    For Each objPrinter In colPrinters
        If objPrinter.Name = strPrinterName Then
            objPrinter.SetDefaultPrinter
   '         MsgBox "Adobe PDF has been set as the default printer.", vbInformation
            Exit Function
        End If
    Next
    MsgBox "Adobe PDF printer not found.", vbExclamation
End Function

Public Function IsOpen(ByVal StrFormName As String) As Boolean
    'Returns True if the specified form is open in form view.

    Const ConDesignView = 0
    Const conObjStateClosed = 0
    On Error GoTo FixMe
    IsOpen = False
    '  If SysCmd(acSysCmdGetObjectState, acForm, StrFormName) <> conObjStateClosed Then

    If Forms(StrFormName).CurrentView <> ConDesignView Then IsOpen = True
    '      End If

    ' End If
FixMe:
End Function

'-----------------------------------------------------------------------------'
' General Clipboard context menu, the basis for all forms                     '
'-----------------------------------------------------------------------------'
Public Function CreateShortCutMenus()
    On Error Resume Next
    CommandBars("SortClip").Delete
    CommandBars("PrintMe").Delete
    CommandBars("FormsCmd").Delete

    Dim cmb As CommandBar
    Dim cmbControl As Office.CommandBarControl
    Set cmb = CommandBars.Add("SortClip", msoBarPopup, False, False)

    With cmb
        If LCase(Environ("username")) = "ljohnson" Or LCase(Environ("username")) = "leojo" Then _
                                                                                   .Controls.Add(msoControlButton, 2952, , , True).BeginGroup = True '   Design View

        .Controls.Add msoControlButton, 21, , , True ' Cut
        .Controls.Add msoControlButton, 19, , , True ' Copy
        .Controls.Add msoControlButton, 22, , , True ' Paste
        '  .Controls.Add msoControlButton, 4, , , True  ' Print
        .Controls.Add(msoControlButton, 210, , , True).BeginGroup = True ' Sort ASC
        .Controls.Add msoControlButton, 211, , , True ' Sort Desc
        .Controls.Add(msoControlButton, 605, , , True).BeginGroup = True ' add/Remove Filer/Sort
        .Controls.Add msoControlButton, 11761, , , True ' Auto Filter menu

        ' .Controls.Add msoControlButton, 640, , , True 'Add the Filter by Selection
        ' .Controls.Add msoControlButton, 3017, , , True ' Excluding
        .Controls.Add msoControlButton, 10062, , , True ' Between
        .Controls.Add msoControlButton, 10068, , , True '     equals selection
        .Controls.Add msoControlButton, 10071, , , True ' not equals selection
        .Controls.Add msoControlButton, 10076, , , True '     Contains
        .Controls.Add msoControlButton, 10089, , , True ' not Contains

    End With

    Set cmb = CommandBars.Add("PrintMe", msoBarPopup, False, False)

    With cmb
        If FAUserAccess = 99 Then .Controls.Add(msoControlButton, 2952, , , True).BeginGroup = True '   Design View
        If FAUserAccess = 99 Then .Controls.Add msoControlButton, 13157, , , True '  Layout View
        .Controls.Add(msoControlButton, 923, , , True).BeginGroup = True ' Close
        .Controls.Add msoControlButton, 247, , , True ' Page Setup
        .Controls.Add msoControlButton, 15948, , , True ' QuickPrint
        .Controls.Add msoControlButton, 11253, , , True ' Report View
        .Controls.Add msoControlButton, 109, , , True ' Print Preview
        .Controls.Add msoControlButton, 12499, , , True ' PDF/XPs
        .Controls.Add(msoControlButton, 2188, , , True).BeginGroup = True ' Email attachment

    End With

    Set cmb = CommandBars.Add("FormsCmd", msoBarPopup, False, False)
    With cmb
        .Controls.Add msoControlButton, 21, , , True ' Cut
        .Controls.Add msoControlButton, 19, , , True ' Copy
        .Controls.Add msoControlButton, 22, , , True ' Paste
        .Controls.Add(msoControlButton, 2952, , , True).BeginGroup = True '   Design View
        .Controls.Add msoControlButton, 502, , , True '   Form View
        .Controls.Add msoControlButton, 123292, , , True ' Datasheet View
        .Controls.Add msoControlButton, 13157, , , True '  Layout View
        .Controls.Add msoControlButton, 1576, , , True ' Insert Hyperlink
    End With
    Set cmb = Nothing
End Function

Public Function SetAllReports_ShortCutMenu(ShortCutMenuX As String)
    Dim objAcc As AccessObject
    Dim objAll As Object
    Dim TargetRpt As String

    Set objAll = Application.CurrentProject
    For Each objAcc In objAll.AllReports
        Call SetShortCutReport(objAcc.Name, ShortCutMenuX)
        '      Debug.Print objAcc.Name
    Next objAcc
    MsgBox "Done!"
End Function

Public Function SetAllForms_ShortCutMenu(ShortCutMenuX As String)
    Dim objAcc As AccessObject
    Dim objAll As Object
    Dim TargetRpt As String

    Set objAll = Application.CurrentProject
    For Each objAcc In objAll.AllForms
        Call SetShortCutForm(objAcc.Name, ShortCutMenuX)
        '      Debug.Print objAcc.Name
    Next objAcc
    MsgBox "Done!"
End Function

Public Function SetShortCutReport(ReportX As String, ShortCutX As String)
    DoCmd.OpenReport ReportX, acViewDesign, , , acHidden
    Reports(ReportX).ShortcutMenuBar = ShortCutX
    Reports(ReportX).PopUp = True
    Reports(ReportX).AutoResize = False
    DoCmd.close acReport, ReportX, acSaveYes
End Function

Public Function SetShortCutForm(FormX As String, ShortCutX As String)
    DoCmd.OpenForm FormX, acViewDesign, , , acHidden
    Forms(FormX).ShortcutMenuBar = ShortCutX
    Forms(FormX).PopUp = True
    'Forms(FormX).AutoResize = False
    DoCmd.close acForm, FormX, acSaveYes
End Function

Public Function fGetBaseFileName(strFilePath As String) As String
    'This Function accepts the Absolute Path to a File and returns the Base File
    'Name (File Name without the Extension)

    'Make absolutely sure that it is a valid Path/Filename
    If Dir$(strFilePath) = "" Then Exit Function

    Dim strFileName As String
    Dim strBaseFileName As String

    strFileName = Right$(strFilePath, Len(strFilePath) - InStrRev(strFilePath, "\"))

    strBaseFileName = Left$(strFileName, InStr(strFileName, ".") - 1)
    fGetBaseFileName = strBaseFileName
End Function

Public Function GetLFR(Rentx As Double, OECx As Double)
    If Nz(Rentx, 0) = 0 Or Nz(OECx, 0) = 0 Then
        GetLFR = Null
    Else
        GetLFR = (Round(Rentx / OECx, 6))
    End If
End Function

Public Function OpenBuyersFrom(BuyerIDx As Long, BOrderIDx As Long)
    DoCmd.close acForm, "BuyersForm"
    DoCmd.OpenForm "BuyersForm"
    If BuyerIDx & "test" <> "test" Then
        Forms!BuyersForm.BuyerList = BuyerIDx
        Form_BuyersForm.BuyerList_Click
    End If
    Forms!BuyersForm.Form.TabControls.Pages(2).SetFocus
    Pause (0.5)
    If BOrderIDx > 0 Then
        Forms!BuyersForm.RemarketingInvoiceRequest_Sub.Form.InvoiceRequestID = BOrderIDx
        Form_RemarketingInvoiceRequest_Sub.InvoiceRequestID_Click
    End If
    DoCmd.SelectObject acForm, "BuyersForm"
End Function

Public Function OpenUnitsForm(AssetIDx As Long, Optional GroupIDx As Long)
    Dim SQLz, WhereX As String
    WhereX = " AssetID = " & AssetIDx
    DoCmd.close acForm, "UnitsForm"
    If GroupIDx > 0 Then WhereX = " SGrpID = " & GroupIDx
    SQLz = "Insert into UnitsFormFilter ("
    SQLz = SQLz & "  AssetIDx"
    SQLz = SQLz & ", ClientShNm "
    SQLz = SQLz & ", MLOrig"
    SQLz = SQLz & ", MLNo"
    SQLz = SQLz & ", UnitOEC"
    SQLz = SQLz & ", SchNo"
    SQLz = SQLz & ", sGrpNo"
    SQLz = SQLz & ", ProjOffLease"
    SQLz = SQLz & ", CurrRent"
    SQLz = SQLz & ", CurrRentType"
    SQLz = SQLz & ", LeaseExp"
    SQLz = SQLz & ", BLCD"
    SQLz = SQLz & ", FundedDate"
    SQLz = SQLz & ", SchID"
    SQLz = SQLz & ") Select"
    SQLz = SQLz & "  AssetID"
    SQLz = SQLz & ", ClientShortName As ClientShNm"
    SQLz = SQLz & ", MLOrig"
    SQLz = SQLz & ", MLNo"
    SQLz = SQLz & ", OECUnit"
    SQLz = SQLz & ", Schedule As SchNo"
    SQLz = SQLz & ", UnitGroup as sGrpNo"
    SQLz = SQLz & ", OffLeaseProj"
    SQLz = SQLz & ", CurrRent"
    SQLz = SQLz & ", CurrRentType"
    SQLz = SQLz & ", DateAdd('m',LeaseTermPrimary,BLCD)-1"
    SQLz = SQLz & ", BLCD"
    SQLz = SQLz & ", FundedDate"
    SQLz = SQLz & ", SchID"
    SQLz = SQLz & " From vw_SixKeys Where" & WhereX
    '
    DoCmd.SetWarnings False
    DoCmd.RunSQL ("Delete * From UnitsFormFilter")
    DoCmd.RunSQL (SQLz)
    DoCmd.SetWarnings True

    On Error Resume Next
    Forms!Unitsform.Filter = ""
    DoCmd.OpenForm "UnitsForm", , , "UnitID=" & AssetIDx
  '  Forms!Unitsform.RecordSource = "Select * From UnitsForm"
    On Error GoTo 0
End Function
