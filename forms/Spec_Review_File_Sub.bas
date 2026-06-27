Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    AllowEdits = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    ScrollBars =0
    TabularCharSet =204
    BorderStyle =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =9540
    DatasheetFontHeight =11
    ItemSuffix =11
    Left =2955
    Top =2295
    Right =23805
    Bottom =12285
    ShortcutMenuBar ="SortClip"
    RecSrcDt = Begin
        0x0c38c1bbae93e340
    End
    RecordSource ="SELECT AttachedFiles.AttachedFilePath, AttachedFiles.AttachedFileTypeRef, Attach"
        "edFiles.GroupRef FROM AttachedFiles WHERE (((AttachedFiles.AttachedFileTypeRef)="
        "2)); "
    Caption ="AttachedFiles subform"
    DatasheetFontName ="Calibri"
    OnLoad ="[Event Procedure]"
    FilterOnLoad =0
    ShowPageMargins =0
    DisplayOnSharePointSite =1
    DatasheetAlternateBackColor =16777215
    DatasheetGridlinesColor12 =15062992
    FitToScreen =1
    Begin
        Begin Label
            BackStyle =0
            TextFontCharSet =204
            FontSize =11
            BorderColor =-2147483609
            ForeColor =11830108
            FontName ="Calibri"
        End
        Begin Rectangle
            BackStyle =0
            BorderLineStyle =0
            BorderColor =-2147483609
        End
        Begin Line
            BorderLineStyle =0
            BorderColor =-2147483609
        End
        Begin Image
            BackStyle =0
            BorderLineStyle =0
            SizeMode =3
            PictureAlignment =2
            BorderColor =-2147483609
        End
        Begin CommandButton
            TextFontCharSet =204
            FontSize =11
            FontWeight =400
            ForeColor =10658208
            FontName ="Calibri"
            BorderLineStyle =0
        End
        Begin OptionButton
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
            BorderColor =-2147483609
        End
        Begin CheckBox
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
        End
        Begin OptionGroup
            BackStyle =1
            BorderLineStyle =0
            BorderColor =-2147483609
        End
        Begin BoundObjectFrame
            SizeMode =3
            BorderLineStyle =0
            BackStyle =0
            LabelX =-1800
            BorderColor =-2147483609
        End
        Begin TextBox
            FELineBreak = NotDefault
            TextFontCharSet =204
            BorderLineStyle =0
            LabelX =-1800
            FontSize =11
            BorderColor =-2147483609
            FontName ="Calibri"
            AsianLineBreak =1
        End
        Begin ListBox
            TextFontCharSet =204
            BorderLineStyle =0
            LabelX =-1800
            FontSize =11
            BorderColor =-2147483609
            FontName ="Calibri"
            AllowValueListEdits =1
            InheritValueList =1
        End
        Begin ComboBox
            TextFontCharSet =204
            BorderLineStyle =0
            LabelX =-1800
            FontSize =11
            BorderColor =-2147483609
            FontName ="Calibri"
            AllowValueListEdits =1
            InheritValueList =1
        End
        Begin Subform
            BorderLineStyle =0
            BorderColor =-2147483609
        End
        Begin UnboundObjectFrame
            BackStyle =0
            OldBorderStyle =1
            BorderColor =-2147483609
        End
        Begin CustomControl
            OldBorderStyle =1
            BorderColor =-2147483609
        End
        Begin ToggleButton
            TextFontCharSet =204
            FontSize =11
            FontWeight =400
            ForeColor =10658208
            FontName ="Calibri"
            BorderLineStyle =0
        End
        Begin Tab
            TextFontCharSet =204
            FontSize =11
            FontName ="Calibri"
            BorderLineStyle =0
        End
        Begin FormHeader
            Height =375
            Name ="FormHeader"
            Begin
                Begin CommandButton
                    OverlapFlags =85
                    TextFontCharSet =0
                    Left =60
                    Top =60
                    Width =1680
                    Height =315
                    FontSize =8
                    FontWeight =700
                    ForeColor =10040879
                    Name ="AddDocBtn"
                    Caption ="Add Spec Review"
                    OnClick ="[Event Procedure]"
                    FontName ="Tahoma"
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =60
                    LayoutCachedTop =60
                    LayoutCachedWidth =1740
                    LayoutCachedHeight =375
                End
            End
        End
        Begin Section
            Height =315
            Name ="Detail"
            AlternateBackColor =16777215
            Begin
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =0
                    IMESentenceMode =3
                    Left =5820
                    Width =1380
                    Height =285
                    FontSize =10
                    Name ="AttachedFilePath"
                    ControlSource ="AttachedFilePath"
                    FontName ="Tahoma"

                    LayoutCachedLeft =5820
                    LayoutCachedWidth =7200
                    LayoutCachedHeight =285
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontCharSet =0
                    Left =1740
                    Width =1740
                    Height =315
                    FontSize =8
                    FontWeight =700
                    TabIndex =1
                    ForeColor =10040879
                    Name ="OpenFileBtn"
                    Caption ="View Spec Review"
                    OnClick ="[Event Procedure]"
                    FontName ="Tahoma"
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =1740
                    LayoutCachedWidth =3480
                    LayoutCachedHeight =315
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontCharSet =0
                    Left =3540
                    Width =1920
                    Height =315
                    FontSize =8
                    FontWeight =700
                    TabIndex =2
                    ForeColor =2366701
                    Name ="RemoveFileBtn"
                    Caption ="Delete Spec Review"
                    OnClick ="[Event Procedure]"
                    FontName ="Tahoma"
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =3540
                    LayoutCachedWidth =5460
                    LayoutCachedHeight =315
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =0
                    IMESentenceMode =3
                    Left =9120
                    Width =420
                    Height =315
                    FontSize =8
                    TabIndex =3
                    ForeColor =2366701
                    Name ="GroupRef"
                    ControlSource ="GroupRef"
                    StatusBarText ="Link to Groups Table"
                    FontName ="Tahoma"

                    LayoutCachedLeft =9120
                    LayoutCachedWidth =9540
                    LayoutCachedHeight =315
                End
            End
        End
        Begin FormFooter
            Height =0
            Name ="FormFooter"
        End
    End
End
CodeBehindForm
' See "Spec_Review_File_Sub.cls"
