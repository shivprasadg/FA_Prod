Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    AutoResize = NotDefault
    PopUp = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    TabularCharSet =204
    DateGrouping =1
    GrpKeepTogether =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =14460
    DatasheetFontHeight =11
    ItemSuffix =264
    Left =5445
    Top =2235
    ShortcutMenuBar ="PrintMe"
    RecSrcDt = Begin
        0x92fc2588cdcee440
    End
    RecordSource ="FSLBankBillq_Toyota"
    DatasheetFontName ="Calibri"
    OnLoad ="[Event Procedure]"
    FilterOnLoad =0
    FitToPage =1
    DisplayOnSharePointSite =1
    DatasheetAlternateBackColor =16053492
    DatasheetGridlinesColor12 =15062992
    FitToScreen =1
    Begin
        Begin Label
            BackStyle =0
            TextFontFamily =0
            FontSize =11
            FontName ="Calibri"
        End
        Begin Rectangle
            BorderLineStyle =0
        End
        Begin Line
            BorderLineStyle =0
        End
        Begin Image
            OldBorderStyle =0
            BorderLineStyle =0
            SizeMode =3
            PictureAlignment =2
        End
        Begin CheckBox
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
        End
        Begin TextBox
            FELineBreak = NotDefault
            OldBorderStyle =0
            TextFontFamily =0
            BorderLineStyle =0
            LabelX =-1800
            FontSize =11
            FontName ="Calibri"
            AsianLineBreak =1
            ShowDatePicker =0
        End
        Begin ListBox
            TextFontFamily =0
            BorderLineStyle =0
            LabelX =-1800
            FontSize =11
            FontName ="Calibri"
            AllowValueListEdits =1
            InheritValueList =1
        End
        Begin ComboBox
            OldBorderStyle =0
            TextFontFamily =0
            BorderLineStyle =0
            LabelX =-1800
            FontSize =11
            FontName ="Calibri"
            AllowValueListEdits =1
            InheritValueList =1
        End
        Begin Subform
            BorderLineStyle =0
        End
        Begin BreakLevel
            GroupHeader = NotDefault
            GroupFooter = NotDefault
            ControlSource ="SGrpNO"
        End
        Begin BreakLevel
            GroupHeader = NotDefault
            GroupFooter = NotDefault
            KeepTogether =2
            ControlSource ="LocationName"
        End
        Begin BreakLevel
            ControlSource ="unitunitnum"
        End
        Begin BreakLevel
            ControlSource ="vin4"
        End
        Begin FormHeader
            KeepTogether = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =2103
            Name ="ReportHeader"
            Begin
                Begin TextBox
                    OldBorderStyle =1
                    TextFontCharSet =186
                    TextAlign =2
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =11850
                    Top =1680
                    Width =2430
                    Height =315
                    Name ="SchBnkClNO2"
                    ControlSource ="=[SchBnkClNO] & \"-\" & [SgrpBnkNo]"
                    StatusBarText ="Assignee Client No."

                    LayoutCachedLeft =11850
                    LayoutCachedTop =1680
                    LayoutCachedWidth =14280
                    LayoutCachedHeight =1995
                    Begin
                        Begin Label
                            TextFontCharSet =238
                            TextAlign =1
                            TextFontFamily =34
                            Left =11850
                            Top =1395
                            Width =2430
                            Height =315
                            FontWeight =700
                            Name ="Label218"
                            Caption ="Bank Contract Number:"
                            LayoutCachedLeft =11850
                            LayoutCachedTop =1395
                            LayoutCachedWidth =14280
                            LayoutCachedHeight =1710
                        End
                    End
                End
                Begin TextBox
                    TextFontCharSet =186
                    TextAlign =1
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =3645
                    Top =1665
                    Width =6525
                    Height =300
                    FontSize =12
                    FontWeight =700
                    TabIndex =1
                    ForeColor =-2147483615
                    Name ="Text1"
                    ControlSource ="=\"Rent Bill\" & \" for Period from \" & [BillFrom] & \" through \" & [BillTo]"

                    LayoutCachedLeft =3645
                    LayoutCachedTop =1665
                    LayoutCachedWidth =10170
                    LayoutCachedHeight =1965
                End
                Begin TextBox
                    Visible = NotDefault
                    TextFontCharSet =186
                    TextAlign =2
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =10245
                    Top =1770
                    Width =1065
                    FontSize =8
                    TabIndex =2
                    ForeColor =10855845
                    Name ="BillID"
                    ControlSource ="BillID"

                    LayoutCachedLeft =10245
                    LayoutCachedTop =1770
                    LayoutCachedWidth =11310
                    LayoutCachedHeight =2010
                End
                Begin TextBox
                    TextFontCharSet =186
                    TextAlign =1
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =3645
                    Top =1260
                    Width =6525
                    Height =360
                    FontSize =14
                    FontWeight =700
                    TabIndex =3
                    Name ="Text228"
                    ControlSource ="=UCase([MLNO]) & \" | Schedule \" & [schno]"

                    LayoutCachedLeft =3645
                    LayoutCachedTop =1260
                    LayoutCachedWidth =10170
                    LayoutCachedHeight =1620
                End
                Begin Rectangle
                    BackStyle =0
                    OldBorderStyle =0
                    Left =3495
                    Top =90
                    Width =6720
                    Height =1155
                    Name ="Box229"
                    LayoutCachedLeft =3495
                    LayoutCachedTop =90
                    LayoutCachedWidth =10215
                    LayoutCachedHeight =1245
                End
                Begin Label
                    SpecialEffect =5
                    OldBorderStyle =1
                    TextFontCharSet =186
                    TextFontFamily =34
                    Left =2700
                    Top =45
                    Width =690
                    Height =255
                    FontSize =10
                    FontWeight =700
                    Name ="Label5"
                    Caption ="Bill To:"
                    LayoutCachedLeft =2700
                    LayoutCachedTop =45
                    LayoutCachedWidth =3390
                    LayoutCachedHeight =300
                End
                Begin Label
                    TextFontCharSet =186
                    TextAlign =1
                    TextFontFamily =34
                    Left =150
                    Width =2310
                    Height =675
                    FontSize =32
                    FontWeight =700
                    Name ="Label231"
                    Caption ="INVOICE"
                    LayoutCachedLeft =150
                    LayoutCachedWidth =2460
                    LayoutCachedHeight =675
                End
                Begin Image
                    Left =120
                    Top =720
                    Width =2334
                    Height =1269
                    Name ="Auto_Logo0"
                    Picture ="SmallLogo.gif"
                    ImageData = Begin
                        0x47494638396119018800f70000ffffff7bc542008cc56b6b63635a63cecece6b ,
                        0x6b73f7efef42a5d6cee6efe6e6ef7b6b739ca5ef9c6b42736b1073a5106b5a9c ,
                        0x42e6529ca51010e65210cee69ca5c5a5d6ef949494a5105aa510194210de4210 ,
                        0x9c42cea5219ca519a52919a5087b7b84429ca510735abdc5bd0894ce3a739c10 ,
                        0x315a6ba5c542525a9c63ef3a73de42315a9c31de42ef194231199c319c9cefa5 ,
                        0x42a5736b63ef6b31de1031196b319c10a5731073196befa510ef1973ef631031 ,
                        0xde10319c10efa53a529c3a52de42ce191052196bcea510ce199c7394a5adad6b ,
                        0x6b42e6e6de109cef4a9c3108529ca56b6be6f7ff42efe6cec510ce5210ce8c10 ,
                        0xce19104aa5109c9c73848c8c639c9c739c7342e6739ca53110e67310efe69c52 ,
                        0x9c9ca552ded6d60873dea5ef634231de42319c42efa563b5e6ce6ba5ce29a510 ,
                        0x52dece4aa5ce08a5219cce427329739c52009ca508738c9c5a1010526b42c5de ,
                        0xa5efe6ef9cb5cec531ce5231ce8c31ce1931cec573ce6be6ce52739cef10ce29 ,
                        0xe6ce8c73ce1973ceade6efc510ef5210ef8c10ef1910ef9c94cec552ce4ae6ce ,
                        0x52529cce10ce08e6ce8c52ce1952ce8ce663efe69ca59c0094c57bc5deceefc5 ,
                        0x6bce196b9c3173b53aa5ce94425229ef6ba5ef29a56bef19ef4aa5ef08a59cce ,
                        0x5a429cefe6ceb5733a6b733a2973106b73102910106b9c73ce42106b9c10ef42 ,
                        0x10299c10ad42b5526b73ce6b10ef1010296b10ad10b5521010ef1010ad10ceb5 ,
                        0xbdce9cceef6bceef10e6ce946b94ef733a4a733a0873104a73100810104a9c52 ,
                        0xce42104a9c10ce4210089c108c4294526b52ce6b10ce1010086b108c10945210 ,
                        0x10ce10108c10ce9473ce52c59cb50873ad42734a9cced610524aefc573ef6be6 ,
                        0xef52739cef31efc531ef5231ef29e6ef8c31ef1931ef8c73ef1973efade6c59c ,
                        0x94efc552ef4ae6ef52529cce31ef08e6ef8c52ef1952ef8ce663c5de42736b9c ,
                        0x7b10a5ce737ba59c84efe6a53a6befef9ca53a29ceef42a53a4aa53a089c94a5 ,
                        0x9cceb5425208efef6befef1073ce73427308c5cee66b7ba5e6ceefefefc5c5ef ,
                        0x9cefef426be642ffefff6bc5422c00000000190188000008ff0001081c48b0a0 ,
                        0xc18308132a5cc8b0a1c38710234a9c48b1a2c58b18336adcc8b1a3c78f20438a ,
                        0x1c49b2a4c9932853aa5cc9b2a5cb973063ca9c49b3a6cd9b3873eadcc9b3a7cf ,
                        0x9f40830a1d4ab4a8d1a348375a7884ee44d3a74ea3c61b71e040d2ab5853364b ,
                        0x23a0cdb2365ec17e6da3a4cd9b12e80a58cdcab66d472608c8911040b7ae5d36 ,
                        0x5fdfbc29d265addbbf8025361300c7115dc3020c23a69bb7ca080581234b5e88 ,
                        0x204d1ac48bed26eeea0c45b323fe268b1ead804461cc9b15db5df606c5882ea3 ,
                        0x634b26970681eac37533b32e479589ecdf6d0fa4a1c03535eecd9bcd82500bbc ,
                        0x39d6c17034a3464e572f5fe7d89122a87c7b7aeadd238e64ff1f3f3481a3c2c7 ,
                        0x914f6fe30cde3df2f081d2b68d7b7a66bdafe3ebdf79e03c89eec6a9c61e0854 ,
                        0xed67e04dcd989699660c2ab15717be1d28214c5d405242086065a8e1866089e0 ,
                        0x4c3917c452c008234042a289259ea8628a54f935e18b129d234912334a62e38d ,
                        0x38e698c41a49d0d8638d34aa01a42442fa48642556dce3228c4c3214cb269b4c ,
                        0x2265009350f98f955656a9a5945c48e9a597544eb2c6945faea10633caf4d5e4 ,
                        0x9a0b41b2493ce70420e79c74d669a79d5cdc190017c9c839891a955c00c90111 ,
                        0xb2692841076c124b9c7e52e9a8968e465a65a393243369a473fea325336a50f1 ,
                        0xd8a1a012440925b1d4c9a79e92ce79e9a49652aa2a9103bcffb664a86b3eb9cf ,
                        0x26573a7aa596c970b16baf745609acaa72e6196c0066a2a04c01b4860ac939b7 ,
                        0x426ae7aa74b6aa6ba39862fba71a8d14d8aca1898e5a27a4d4ba6aeca3739e2b ,
                        0xed9f95785ad5b7868e5a2ab1d9627a6900bb8e8b6fbd93701a2b84f0b219cb17 ,
                        0xf3eabba7bdf4127baebe6b98590981a005bce6b34f127b6fae7bf679b0abd3d2 ,
                        0xb9c68c6aa87141016a4acc64a2b15072a7b1174fa969a52c93bb46983c866c04 ,
                        0x817d156af28b949c03c3b155feeaa82446f458741247abb1a3d244fa1872c856 ,
                        0x5c20ebac3b1fb80fb451a2ba3130674e6185d754783dc514207c6d45d8674f41 ,
                        0x4523cabca600a155377984a25953da6ad0011069c4b20a1cff71c01180f7fdf7 ,
                        0xdb472820b8df07f8f36edc27bfc968a3c35eea70112330e1cf52e854e0540599 ,
                        0x67be79531508e1d4635433ae5fcf946c82edea7e26d1aeb75b7515d6ec1c8285 ,
                        0xdf113a9bae9fad75b3bee51a41c4fab6400ac4b520837629c79ceefbcd3d02ae ,
                        0xc7a22be71af60cc0978b8f38421ff2c89b753df3f1250ac3e3900eebe79f2834 ,
                        0xb23c41955da6dafb9a99c51bdce08f0703a9c1f679e9ddaebb9b3b007150907d ,
                        0xea5317b38cac74f51bcdadf6113d96398a0bae8b957810321f00a9a72eac1141 ,
                        0x7e12d89cb9552c5b185355120850840208ee006f7b5b55eed10104cce5821734 ,
                        0x8cf210c841c078506540e39898d4d0b01efab0874628410bff0708430158e73d ,
                        0x358c8dbcb0b5b07b41aa4643f2913c80518410588688a8710478fc96c4c92cb0 ,
                        0x77fba296b0eea5a73335a20bf7388f0507c81e7814e07f5d64cb11cef1a41096 ,
                        0xeb540723639d00b51cf1541086f661cd1b361847b7840b87da9a93b5f4782760 ,
                        0x20293ca1e90f1c5eb898013a4239bd29645b9698303e55696139a493dee05132 ,
                        0x0040a792e9c90c7baea3c9ac7c715295d263bec815805ee56b72a51448fb1c61 ,
                        0xc9e3c8af45ad4c8a3eeaf1243d8a914eff38151927d1bf4c12e4086a3c0e2a93 ,
                        0xf7062aac2f984439001d5596ab4f26cc77166317c44263900a621139827c0d1c ,
                        0xb1e9939e150c997eba5baa4c65294059415608118e65a8e3ff9dc45cf20dcba1 ,
                        0x213b73f2ca843112617e9a512598e129d828642be8e12701e9b24a240ef42747 ,
                        0x20e63930064b8b152b6192a8991a06704605ac932070b8e244a7b945815e5426 ,
                        0x2853d93257d7251f7d8c4744b21908aee75281dc230de8992680d843855cbe14 ,
                        0x273d63604793a9b13c26410d4660c61a806104aa32e36b9453d2e21cf2470baa ,
                        0x479045e0e2517172b562766c52575a4325ac47b22eb4b50b5d50a1404e9a4fcb ,
                        0x5c26372b4d0c6b9643d7b1b68462d08b5e9d24b1d02228090056e9a9432c10cd ,
                        0x225287332554ac5f4da24d71458f96db22a56425420eedf5923abfdcea645f92 ,
                        0x548315aba9cc789d583f02cdd348f3b55d79435ae837da96ff9435b0e5da9598 ,
                        0xe4618522406624d98b0e5e012980743ab4b62cf1e0392ec5d4724da2b0460549 ,
                        0x6d8a73cec498c500059820725372c850626a726ae9ab46181b51e2226795d1dd ,
                        0xee4846a5543d7514500d45c9091c71d7ea16f70dbc199e7a4bb240d33ad19e9f ,
                        0x4209133ac04b5e0e17866d30c16cf74b92392eeab2d832931a70b6598d1cc17d ,
                        0x00eaa77556cb608f74975e4e546ba74c484e95dc23046cf0a7448d8349f17698 ,
                        0x22ecd59a941cc68c12ea97250a38c13252ac18f809e841157e314320f1e0194b ,
                        0x424c620ae953a13ab2f4a6c437cd180b5e0430e52a2f23b6f90db290f3498964 ,
                        0x1089534f0db3d2a01ab5021c962607680686c05296379885ff2c6f7e83829dbc ,
                        0xe5881c0045241a41dbf44ca2b6dd23ae5a2e89e2bad08c22a0a3089ce39ca113 ,
                        0x7de81bd7f922fe505c6201104972067a2591f40bd55cfce84e7bfad3a00eb5a8 ,
                        0x474dea529bfad4a80e665556cdea56bb5ab4466135a719c26abf8e0004b8ceb5 ,
                        0xae779d6b058c60018d400a130e4005102c800a17b030ae2fadbb220ce0d90320 ,
                        0x40b4a72ded6a479b0047288201928d94460cc000d06616460a308005ccba8623 ,
                        0x30c0022897e776bb9b444c18c100b89d9202904c235d5880010804827d331b00 ,
                        0x4c3882bab57b93aad8fbdf3f51c6b7cfad0c02045b25548895462ef06d66c9db ,
                        0x0023c04817fc7d6e951480002040b84fb4bd80e336241e04ffa0374aa840808c ,
                        0x8fdb00db16883fa8a06e934fa40b030001c16f728f9c8bbc27e936c06f1d226f ,
                        0x781ca42a85439c4091fe37142ec92f111f41a6113b698710bbdcc725b7018ad0 ,
                        0x10a6176e713d0fb964578da8a6038eec05112dd2fdc661c40a041203a002fdde ,
                        0x056baaff8ded68cfa7d9952e90ba539ded2aecb8412ebef385385be55d6844b1 ,
                        0xfbbd007d1f5b6a4f1fc10570ad6f981f7b6d0e3d02158edd781050e1f3a007bd ,
                        0xcd137271ae13e402354f48178a307963ab5bddc8ceaeba15a078951ffd0017a0 ,
                        0x02d70f20f96257def1c8f616008a10d0112cfef5976765118ecdf8d03bdff400 ,
                        0x2840117cfffa7d83e002614d7bef716d797563ff008a17b7ff40b6eff87d2fe0 ,
                        0x02eaab08c987de1065c47cfcd02ef7aeaf4d05ed6a7eda0bc0f5e29f4d008c03 ,
                        0x20dfd5a76fe5b600e5066e0bc07e0971008eb7241b3700467710054080d2967f ,
                        0xbaf66cdb967fbcf76da357108d400007787513c86bd0f68057377d28f06c1488 ,
                        0x6bf8071b8dd078af576e280883f3067f162882cf0602c7d50511276deaa66b56 ,
                        0x206d933700e27675355881d1b67514717123700f24c38468d484506815872710 ,
                        0x3867009e323cc38642ab076e17b01614b775b8d36a6e157103f01e29447145e0 ,
                        0x7486a3863da56d486810da268469a78039171e5691852854005f680556417172 ,
                        0x87100147805ca77545900064376c47700ff0006e2e177182ffa8167758150ab0 ,
                        0x7c71376c86436e3aa7426c9858201077900870ab7604cd40807f080014b76e26 ,
                        0x14895ae88672387ccf6658ae968814d77213e17ee0167fb8888b5c577402010f ,
                        0x39475705006eb0b171fea71011a77211277e73051100480508a100e0e68c05e1 ,
                        0x6c38b88c06c10411070290d1802e77107e2810454000d2981011278d9b67001b ,
                        0x3810e9b600da158ce308471b4700160581df261e3d578c09e186cca28006f080 ,
                        0x09e18b5450620fe1869dd78306997fafa70cc3977202010204a0900b4173bba8 ,
                        0x6e02158c0b007512a77ee0a68c05916e193910a8677b1008736b418d54138cc5 ,
                        0xe856e93810ee281034279205c1049d287e3817720b616f0bff2193650800f206 ,
                        0x02b44680cc720f24b9101b577212216f0ba02429b49428c4943ce98000a0009d ,
                        0xc8177f666f66e6567fa683d6e38a2050955e7995244372f44373dd2811e43683 ,
                        0x09a86f36d977fd56960851936bc18fd037105727925d90882a79975c84899416 ,
                        0x750ce19702e18e3d75007055954770976b1171b0417130199311f71e4809965e ,
                        0x9997bf868e470973e25128b9a333e1986c0a5079b9187fd2166dc9c68aa3399a ,
                        0xd0c683f50892b4481134277437096ed0a7936e7910d07880ea987a047171c775 ,
                        0x00c4778b476881ebf66b3e09002cc791063173ad0800615791afc97fd0997f3a ,
                        0x4800cce28bfe688c04001b41878ba8796de0469d12e16cb8fff910f2966c02b7 ,
                        0x6dac070fe9d908d8877df0a09e197701f9d708ac579fed5904eae99e50875d07 ,
                        0xe156b43578f237029433a002aa0c4530a00b1821fa569b06916f2117213a3997 ,
                        0x47a06f734971e5866c1790a1b9b77905688ecb498ee04985956810076066d157 ,
                        0x8318aaa1c58682fc1992b4968c4fb96e17000fec699f07da08f989800d117485 ,
                        0xa710da966c07d089f3c81053281111079105019847d76fdf49004efaa44e0a9d ,
                        0xdff67ec6f99109018d06e022f7287e624997fd76013a3a57b716777dd89ac6b8 ,
                        0x9c42799106416ece48721c192128847a72e86ce3989672d89322a170b0499e0c ,
                        0x898d569a4f26c49320d7107fb69f0c0aa308219e9ee77ca0ff3779a1d7894669 ,
                        0x9c04709d0881894bc29854c88804418c2b098efbe6921f7a7488fa7f64baa664 ,
                        0x4a71946a109fc99f1fb79609819241596e615a106814646ed8a3f9f87edf1891 ,
                        0x50d980c8591067397417404207319307316c0bf8100dc86d266978b24aab1b09 ,
                        0x0000a9339fb9930a81a7fe30ac0c9aa4667a04df362bf2e68c14f7700918ad74 ,
                        0xd885574a737228970b018da1baa3abf9100ac76d67494af94489e2968dd7b4a9 ,
                        0xe548102407307d8794b3ea6df8e81071c82ccb0ab0883289dfa68d70288e4167 ,
                        0x7358da08a5c304cd08aa7ef097fcd97730978604a1002c6774d438a403316c8b ,
                        0xc89ff1e6a43bc56e2322a0f2598048748ac247aba7287811ff0b11450a001560 ,
                        0x819f97a1ac977b9d987211d2050ee981180a0f1ada6f1587280b60b42afa6d8d ,
                        0xd98c1db7aae6367cb7d8b32aaaae34c78e24da895c7810ce76a11a8ab419da89 ,
                        0x9d2877cab9ad033173210a8ed1767e1a4a8046d9807187a1389aa1af0994ea18 ,
                        0xb450faa47588a8416a004e9ba1644b73fda7b63e3a7b10916e2a27795e0b9d06 ,
                        0xe829d44a7c04187fe0b66f17307af7f0b2b7b86f45209074f9859d9aab0be072 ,
                        0x05b0a2df066de08633187810ebe8b0d7f89b16389a06c81756e8921b9b80a30a ,
                        0x7095199cdb6651a75bb9df869adbd605a8a78c07700f2472a0cc3b02e1d589bd ,
                        0x397d02089cdf066cbfda10fef07510e174244a988589468eff169328f4674c98 ,
                        0x950b4198e44b6700a0bd1441982ea285e4fb67111370fff9b1eacb040a80464f ,
                        0x9895be91bdbfe5beb416be0097bfe41ba68653984da826c30630c93bba88a56f ,
                        0xa377c06844beb8c31382e711175cb287426e6a7aad3e976aa271005c985ebcd7 ,
                        0x891790c1202c14dad67f76aba2fa06723f97c23d51002d788bfdd77f0d0b0fb3 ,
                        0x2ac380316c59999539c3c3423cc4445cc4b6490edbf108169010cd308ff70007 ,
                        0x7191c405f108e26501db410e8f6052049100513c8f4cf00848b41453dc0cc483 ,
                        0xc451fcbf541c9360ac105cbc1d096010f7301fe430a40a400eec57c7437700e4 ,
                        0xf0c600f00857bc1d4bcc9c694c10cdb01d69f0087c3c1016ff40c605710f830c ,
                        0x708f501b81ec160a2000586c1b08403540b5c58ff00884d1c9255bc98cfc508e ,
                        0xf008b4e108a30c001660c99df508049106e9e0530290c8004002aecc9c96dcc9 ,
                        0x763c108e400eb9430293acaab5d10c29c57e9e8cc585dccb5bec08b4cc048e60 ,
                        0x514c00cc02b114e4c0ca614c0204c10472410ecde0c78e10cc8c15ccaa8ccd03 ,
                        0x411b9dfccd7f11cdbfa5009bdcc88e40023a9a0624ebc908c0107e3c10d96351 ,
                        0x1690060291206b9b06810c17e43010fe2c100940ce31c915f54c97009d108c35 ,
                        0x10f75028cd80ce03611ea37cd0d49a06892c1cf3980002e022fbcc3e6960cc24 ,
                        0x90c8abfccc84cccfc4c3cc06bdc34711cd89ecd1431ac503ff5d100d9dcd872c ,
                        0xcd0a71cfe59c0611f253060dcf039106a31cc9031117154d029a1cc9352d1044 ,
                        0x0d886944cb7375c80661012400a1255db24ca0d1259b0671f0b18ec07e719006 ,
                        0x6ba1002c5d1048acc8877cd6aaacd20047d59161d6ec8700b71cd40ab0cf3a23 ,
                        0x1ce27ccfe4d0d408c1d32bfd5bfb4c28080007af3cc9cefc1ece1cd3bddcd78f ,
                        0x0cd01e7dcb5bfdd509d1ca3af353ff93d8063dd25bacd1102acfcbac33212d10 ,
                        0x690dc78eb016560c009dc5c7a3ddcf248000529d154cb0d5b4dcd704410e860d ,
                        0x004f7dd8364dc61e1ddb0401d80097d500601e69f0dab47cd3ba34d0915c2807 ,
                        0xbd1d7131744f1dd94e2dce557d1975ddda361dd0c43dd5c9ff1dcc074dcb834d ,
                        0xda7e5db2c46dd5a49dd5da5ddc9d850028bc13669ddc757d000280008b5c1b38 ,
                        0x0dcea5bc14b6bcd30b1dd4b3ddd6956110cacd9c4a4dd79d7d74379d46ae6cdf ,
                        0x0d31d1813cde091dc818ddd91bedd915ed08a2edd6466d10078dda6e2d1707b0 ,
                        0xde9b5dde5841db196e51d903ddb6f1dd28355d868c40c21dc96b31d65b5dd74e ,
                        0x9dca5b2d173a73d0610acb106dcbbb7d748532064d4d02b15c108cf5d9cdccd6 ,
                        0x5b0ddedd0d40641dd4b18dc411a2dd95910e6e1de0d37cda6e91003eed1bec9c ,
                        0xdb39ce3e64aedb93ecd143e7ccd68dcf640ee110cdcf8ce5e2c15ddf0501e699 ,
                        0x3de4aacc4b243b5708eed475edc91d5dca19febe705ddc61f1bdc5574d10583e ,
                        0xe5f83ccb6a5d100820006e5dc8f8fcdf6ce1cc58dcde7eb1e404710f902e1014 ,
                        0x0d00508cd68e6e108361cab611cc5b61159d25dda91c95a36ed0a5dcc9593c10 ,
                        0x3aadc8b35e1013adc42e843dbdbcc8d5ecd7668dcc571eecd953dee621de3ecd ,
                        0x3e24400e1660c5a8cce88e601003bed2f69d3db18e15757cc57b8de30007ede5 ,
                        0xccc70700db05a1007030ab561c17bbbcc5beac4b933cc7e5e4e150bc1d7090c8 ,
                        0x739c3b887c52569c06edfee995b11d43ca04f5feee3e955208b0ed09f0ef0004 ,
                        0xeea6d43e0cdff024ead70980f06d3e5922f7de46bcf11cdff11efff1201ff222 ,
                        0x3ff2245ff226bf650101003b
                    End

                    LayoutCachedLeft =120
                    LayoutCachedTop =720
                    LayoutCachedWidth =2454
                    LayoutCachedHeight =1989
                    TabIndex =9
                End
                Begin Rectangle
                    BackStyle =0
                    OldBorderStyle =0
                    Left =3495
                    Top =1275
                    Width =6720
                    Height =735
                    Name ="Box233"
                    LayoutCachedLeft =3495
                    LayoutCachedTop =1275
                    LayoutCachedWidth =10215
                    LayoutCachedHeight =2010
                End
                Begin Label
                    SpecialEffect =5
                    OldBorderStyle =1
                    TextFontCharSet =186
                    TextFontFamily =34
                    Left =2760
                    Top =1155
                    Width =675
                    Height =240
                    FontSize =10
                    FontWeight =700
                    Name ="Label232"
                    Caption ="LEASE:"
                    LayoutCachedLeft =2760
                    LayoutCachedTop =1155
                    LayoutCachedWidth =3435
                    LayoutCachedHeight =1395
                End
                Begin TextBox
                    Visible = NotDefault
                    TextFontCharSet =186
                    TextAlign =2
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =10260
                    Top =1500
                    Width =720
                    FontSize =8
                    TabIndex =4
                    ForeColor =10855845
                    Name ="clientID"
                    ControlSource ="clientID"

                    LayoutCachedLeft =10260
                    LayoutCachedTop =1500
                    LayoutCachedWidth =10980
                    LayoutCachedHeight =1740
                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =3660
                    Top =165
                    Width =6405
                    Height =1035
                    TabIndex =5
                    Name ="Text264"
                    ControlSource ="=[BLocStreet] & Chr(13) & Chr(10) & [BLocSt2] & Chr(13) & Chr(10) & IIf(IsNull(["
                        "BLoST3]),Null,[BLoST3] & Chr(13) & Chr(10)) & [blOCcITY] & \", \" & [bLOCST] & \""
                        " \" & Left([BlOCzIP],5) & \"-\" & Right([BlOCzIP],4)"
                    StatusBarText ="default company name"

                    LayoutCachedLeft =3660
                    LayoutCachedTop =165
                    LayoutCachedWidth =10065
                    LayoutCachedHeight =1200
                End
                Begin TextBox
                    OldBorderStyle =1
                    TextFontCharSet =186
                    TextAlign =3
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =11850
                    Top =315
                    Width =2430
                    Height =330
                    FontSize =14
                    FontWeight =700
                    TabIndex =6
                    Name ="bank inv"
                    ControlSource ="=[BillNo]"
                    EventProcPrefix ="bank_inv"

                    LayoutCachedLeft =11850
                    LayoutCachedTop =315
                    LayoutCachedWidth =14280
                    LayoutCachedHeight =645
                    Begin
                        Begin Label
                            BackStyle =1
                            OldBorderStyle =1
                            TextFontCharSet =161
                            TextAlign =1
                            TextFontFamily =34
                            Left =11850
                            Top =45
                            Width =2430
                            Height =345
                            FontWeight =700
                            BackColor =8210719
                            BorderColor =-2147483615
                            ForeColor =16777215
                            Name ="Label170"
                            Caption ="Invoice Number:"
                            LayoutCachedLeft =11850
                            LayoutCachedTop =45
                            LayoutCachedWidth =14280
                            LayoutCachedHeight =390
                        End
                    End
                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Left =11850
                    Top =720
                    Width =2520
                    Height =315
                    TabIndex =7
                    Name ="Text315"
                    ControlSource ="=\"Invoice Date: <strong>\" & [Billdt] & \"</strong>\""
                    TextFormat =1

                    LayoutCachedLeft =11850
                    LayoutCachedTop =720
                    LayoutCachedWidth =14370
                    LayoutCachedHeight =1035
                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Left =11850
                    Top =1050
                    Width =2520
                    Height =300
                    TabIndex =8
                    Name ="Text316"
                    ControlSource ="=IIf([DueonRec]=-1,\"Due Date: <strong>On Receipt</strong>\",\"Due Date: <strong"
                        ">\" & [BillDueDt] & \"</strong>\")"
                    TextFormat =1

                    LayoutCachedLeft =11850
                    LayoutCachedTop =1050
                    LayoutCachedWidth =14370
                    LayoutCachedHeight =1350
                End
            End
        End
        Begin PageHeader
            Height =0
            Name ="PageHeaderSection"
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =285
            Name ="GroupHeader1"
            Begin
                Begin TextBox
                    OldBorderStyle =1
                    TextFontCharSet =238
                    TextFontFamily =34
                    IMESentenceMode =3
                    Top =30
                    Width =2865
                    Height =255
                    FontSize =10
                    FontWeight =700
                    BackColor =8210719
                    BorderColor =-2147483615
                    ForeColor =16777215
                    Name ="group"
                    ControlSource ="=\" ▼ Group: \" & [SGrpNO]"

                    LayoutCachedTop =30
                    LayoutCachedWidth =2865
                    LayoutCachedHeight =285
                End
                Begin Line
                    BorderWidth =2
                    Width =14460
                    Name ="Line221"
                    LayoutCachedWidth =14460
                    BorderThemeColorIndex =2
                End
            End
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            RepeatSection = NotDefault
            Height =675
            BreakLevel =1
            Name ="GroupHeader0"
            Begin
                Begin Label
                    SpecialEffect =5
                    TextFontCharSet =186
                    TextAlign =3
                    TextFontFamily =34
                    Left =12900
                    Top =360
                    Width =1335
                    Height =300
                    FontSize =10
                    FontWeight =700
                    Name ="Label96"
                    Caption ="Total Payment"
                    GroupTable =1
                    LayoutCachedLeft =12900
                    LayoutCachedTop =360
                    LayoutCachedWidth =14235
                    LayoutCachedHeight =660
                    ColumnStart =9
                    ColumnEnd =9
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    SpecialEffect =5
                    TextFontCharSet =186
                    TextAlign =2
                    TextFontFamily =34
                    Left =10680
                    Top =360
                    Width =1170
                    Height =300
                    FontSize =10
                    FontWeight =700
                    Name ="Label91"
                    Caption ="Tax Amt"
                    GroupTable =1
                    LayoutCachedLeft =10680
                    LayoutCachedTop =360
                    LayoutCachedWidth =11850
                    LayoutCachedHeight =660
                    ColumnStart =7
                    ColumnEnd =7
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    SpecialEffect =5
                    TextFontCharSet =186
                    TextAlign =3
                    TextFontFamily =34
                    Left =9750
                    Top =360
                    Width =870
                    Height =300
                    FontSize =10
                    FontWeight =700
                    Name ="Label146"
                    Caption ="Tax Rate"
                    GroupTable =1
                    LayoutCachedLeft =9750
                    LayoutCachedTop =360
                    LayoutCachedWidth =10620
                    LayoutCachedHeight =660
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    SpecialEffect =5
                    TextFontCharSet =186
                    TextAlign =2
                    TextFontFamily =34
                    Left =7860
                    Top =360
                    Width =1140
                    Height =300
                    FontSize =10
                    FontWeight =700
                    Name ="Label69"
                    Caption ="Rent"
                    GroupTable =1
                    LayoutCachedLeft =7860
                    LayoutCachedTop =360
                    LayoutCachedWidth =9000
                    LayoutCachedHeight =660
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    SpecialEffect =5
                    TextFontCharSet =186
                    TextFontFamily =34
                    Left =7155
                    Top =360
                    Width =645
                    Height =300
                    FontSize =10
                    FontWeight =700
                    Name ="Label206"
                    Caption ="Status"
                    GroupTable =1
                    LayoutCachedLeft =7155
                    LayoutCachedTop =360
                    LayoutCachedWidth =7800
                    LayoutCachedHeight =660
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    SpecialEffect =5
                    TextFontCharSet =186
                    TextFontFamily =34
                    Left =3720
                    Top =360
                    Width =3375
                    Height =300
                    FontSize =10
                    FontWeight =700
                    Name ="Label13"
                    Caption ="Description"
                    GroupTable =1
                    LayoutCachedLeft =3720
                    LayoutCachedTop =360
                    LayoutCachedWidth =7095
                    LayoutCachedHeight =660
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    SpecialEffect =5
                    TextFontCharSet =186
                    TextFontFamily =34
                    Left =1710
                    Top =360
                    Width =1950
                    Height =300
                    FontSize =10
                    FontWeight =700
                    Name ="Label12"
                    Caption ="VIN"
                    GroupTable =1
                    LayoutCachedLeft =1710
                    LayoutCachedTop =360
                    LayoutCachedWidth =3660
                    LayoutCachedHeight =660
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    SpecialEffect =5
                    TextFontCharSet =186
                    TextAlign =2
                    TextFontFamily =34
                    Left =885
                    Top =360
                    Width =765
                    Height =300
                    FontSize =10
                    FontWeight =700
                    Name ="Label11"
                    Caption ="Unit"
                    GroupTable =1
                    LayoutCachedLeft =885
                    LayoutCachedTop =360
                    LayoutCachedWidth =1650
                    LayoutCachedHeight =660
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    SpecialEffect =5
                    TextFontCharSet =186
                    TextAlign =2
                    TextFontFamily =34
                    Left =9060
                    Top =360
                    Width =630
                    Height =300
                    FontSize =10
                    FontWeight =700
                    Name ="Label238"
                    Caption ="Taxed?"
                    GroupTable =1
                    LayoutCachedLeft =9060
                    LayoutCachedTop =360
                    LayoutCachedWidth =9690
                    LayoutCachedHeight =660
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    SpecialEffect =5
                    TextFontCharSet =186
                    TextAlign =2
                    TextFontFamily =34
                    Left =11910
                    Top =360
                    Width =930
                    Height =300
                    FontSize =10
                    FontWeight =700
                    Name ="FeeLabel"
                    Caption =" "
                    GroupTable =1
                    LayoutCachedLeft =11910
                    LayoutCachedTop =360
                    LayoutCachedWidth =12840
                    LayoutCachedHeight =660
                    ColumnStart =8
                    ColumnEnd =8
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Line
                    Left =504
                    Top =315
                    Width =0
                    Height =317
                    BorderColor =-2147483615
                    Name ="Line245"
                    LayoutCachedLeft =504
                    LayoutCachedTop =315
                    LayoutCachedWidth =504
                    LayoutCachedHeight =632
                End
                Begin TextBox
                    FontUnderline = NotDefault
                    TextFontCharSet =186
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =285
                    Top =180
                    Width =3375
                    Height =225
                    ColumnWidth =2745
                    FontSize =10
                    FontWeight =700
                    ForeColor =-2147483615
                    Name ="LocationName2"
                    ControlSource ="=[LocationName]"

                    LayoutCachedLeft =285
                    LayoutCachedTop =180
                    LayoutCachedWidth =3660
                    LayoutCachedHeight =405
                    Begin
                        Begin Label
                            TextFontCharSet =186
                            TextFontFamily =34
                            Left =285
                            Width =660
                            Height =195
                            FontSize =8
                            FontWeight =700
                            ForeColor =4144959
                            Name ="Label234"
                            Caption ="Location"
                            LayoutCachedLeft =285
                            LayoutCachedWidth =945
                            LayoutCachedHeight =195
                        End
                    End
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            Height =353
            Name ="Detail"
            Begin
                Begin TextBox
                    TextFontCharSet =163
                    TextAlign =2
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =885
                    Top =30
                    Width =765
                    Height =270
                    FontSize =9
                    Name ="UnitUnitNum"
                    ControlSource ="UnitUnitNum"
                    GroupTable =1

                    LayoutCachedLeft =885
                    LayoutCachedTop =30
                    LayoutCachedWidth =1650
                    LayoutCachedHeight =300
                    RowStart =1
                    RowEnd =1
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    TextFontCharSet =163
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =1710
                    Top =30
                    Width =1950
                    Height =270
                    ColumnWidth =2130
                    FontSize =9
                    TabIndex =1
                    Name ="UnitVIN"
                    ControlSource ="UnitVIN"
                    GroupTable =1

                    LayoutCachedLeft =1710
                    LayoutCachedTop =30
                    LayoutCachedWidth =3660
                    LayoutCachedHeight =300
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    TextFontCharSet =163
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =3720
                    Top =30
                    Width =3375
                    Height =270
                    FontSize =9
                    TabIndex =2
                    Name ="Desc"
                    ControlSource ="Desc"
                    GroupTable =1

                    LayoutCachedLeft =3720
                    LayoutCachedTop =30
                    LayoutCachedWidth =7095
                    LayoutCachedHeight =300
                    RowStart =1
                    RowEnd =1
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextFontCharSet =163
                    TextAlign =3
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =7860
                    Top =30
                    Width =1140
                    Height =270
                    FontSize =9
                    TabIndex =3
                    Name ="Text107"
                    ControlSource ="=IIf([unitsbillrent]>0,[unitsbillrent],\"--  \")"
                    Format ="$#,##0.00;($#,##0.00)"
                    GroupTable =1

                    LayoutCachedLeft =7860
                    LayoutCachedTop =30
                    LayoutCachedWidth =9000
                    LayoutCachedHeight =300
                    RowStart =1
                    RowEnd =1
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextFontCharSet =163
                    TextAlign =3
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =10680
                    Top =30
                    Width =1170
                    Height =270
                    FontSize =9
                    TabIndex =4
                    Name ="Text139"
                    ControlSource ="=IIf([TaxAmt]>0,[TaxAmt],\"--  \")"
                    Format ="$#,##0.00;($#,##0.00)"
                    GroupTable =1

                    LayoutCachedLeft =10680
                    LayoutCachedTop =30
                    LayoutCachedWidth =11850
                    LayoutCachedHeight =300
                    RowStart =1
                    RowEnd =1
                    ColumnStart =7
                    ColumnEnd =7
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    DecimalPlaces =3
                    TextFontCharSet =163
                    TextAlign =3
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =9750
                    Top =30
                    Width =870
                    Height =270
                    ColumnWidth =2160
                    FontSize =9
                    TabIndex =7
                    Name ="LocTaxRt2"
                    ControlSource ="=IIf([TaxAmt]>0,[LTaxRate],\" --   \")"
                    Format ="Percent"
                    GroupTable =1

                    LayoutCachedLeft =9750
                    LayoutCachedTop =30
                    LayoutCachedWidth =10620
                    LayoutCachedHeight =300
                    RowStart =1
                    RowEnd =1
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    TextFontCharSet =163
                    TextAlign =3
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =12900
                    Top =30
                    Width =1335
                    Height =270
                    FontSize =9
                    TabIndex =8
                    Name ="LineTot2"
                    ControlSource ="=IIf([unitsbillrent]>0,[LineTot],\"--  \")"
                    Format ="$#,##0.00;($#,##0.00)"
                    GroupTable =1

                    LayoutCachedLeft =12900
                    LayoutCachedTop =30
                    LayoutCachedWidth =14235
                    LayoutCachedHeight =300
                    RowStart =1
                    RowEnd =1
                    ColumnStart =9
                    ColumnEnd =9
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    TextFontCharSet =163
                    TextAlign =2
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =7155
                    Top =30
                    Width =645
                    Height =270
                    ColumnWidth =1965
                    FontSize =9
                    TabIndex =9
                    Name ="UnitStatus"
                    ControlSource ="UnitStatus"
                    GroupTable =1

                    LayoutCachedLeft =7155
                    LayoutCachedTop =30
                    LayoutCachedWidth =7800
                    LayoutCachedHeight =300
                    RowStart =1
                    RowEnd =1
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    TextFontCharSet =163
                    TextAlign =2
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =9060
                    Top =30
                    Width =630
                    Height =270
                    FontSize =9
                    TabIndex =6
                    Name ="Text236"
                    ControlSource ="=IIf([TaxAmt]>0,IIf([UnitTaxExempt]=-1,\"No\",\"Yes\"),\"--\")"
                    GroupTable =1

                    LayoutCachedLeft =9060
                    LayoutCachedTop =30
                    LayoutCachedWidth =9690
                    LayoutCachedHeight =300
                    RowStart =1
                    RowEnd =1
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin CheckBox
                    Visible = NotDefault
                    Left =495
                    Top =60
                    Width =240
                    Height =165
                    TabIndex =10
                    BorderColor =-2147483615
                    Name ="UnitTaxExempt"
                    ControlSource ="UnitTaxExempt"

                    LayoutCachedLeft =495
                    LayoutCachedTop =60
                    LayoutCachedWidth =735
                    LayoutCachedHeight =225
                End
                Begin TextBox
                    TextFontCharSet =163
                    TextAlign =3
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =11910
                    Top =30
                    Width =930
                    Height =270
                    FontSize =9
                    TabIndex =5
                    Name ="SGrpFSLPmt2"
                    Format ="$#,##0.00;($#,##0.00)"
                    GroupTable =1

                    LayoutCachedLeft =11910
                    LayoutCachedTop =30
                    LayoutCachedWidth =12840
                    LayoutCachedHeight =300
                    RowStart =1
                    RowEnd =1
                    ColumnStart =8
                    ColumnEnd =8
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Line
                    Left =504
                    Top =15
                    Width =0
                    Height =338
                    BorderColor =-2147483615
                    Name ="Line244"
                    LayoutCachedLeft =504
                    LayoutCachedTop =15
                    LayoutCachedWidth =504
                    LayoutCachedHeight =353
                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =255
            BreakLevel =1
            Name ="GroupFooter1"
            Begin
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    TextFontCharSet =238
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Top =225
                    Width =11970
                    Height =30
                    FontSize =20
                    FontWeight =700
                    TabIndex =1
                    ForeColor =255
                    Name ="Text208"
                    ControlSource ="=IIf((IsNull([LTaxRate]) Or [LTaxRate]=0) And [UnitTaxExempt]=No And [sttaxtype]"
                        "=\"rr\",\"TAX RATE MISSING CONTACT OPERATIONS\",Null)"

                    LayoutCachedTop =225
                    LayoutCachedWidth =11970
                    LayoutCachedHeight =255
                End
                Begin Line
                    Left =504
                    Top =135
                    Width =12960
                    BorderColor =-2147483615
                    Name ="Line243"
                    LayoutCachedLeft =504
                    LayoutCachedTop =135
                    LayoutCachedWidth =13464
                    LayoutCachedHeight =135
                End
                Begin Line
                    Left =504
                    Width =0
                    Height =144
                    BorderColor =-2147483615
                    Name ="Line246"
                    LayoutCachedLeft =504
                    LayoutCachedWidth =504
                    LayoutCachedHeight =144
                End
                Begin TextBox
                    OldBorderStyle =1
                    TextFontCharSet =163
                    TextAlign =3
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =7845
                    Width =1110
                    FontSize =9
                    TabIndex =3
                    BorderColor =-2147483615
                    ForeColor =-2147483615
                    Name ="Text247"
                    ControlSource ="=Sum([unitsbillrent])"
                    Format ="$#,##0.00;($#,##0.00)"
                    ControlTipText ="LineTot Total"

                    LayoutCachedLeft =7845
                    LayoutCachedWidth =8955
                    LayoutCachedHeight =240
                End
                Begin TextBox
                    OldBorderStyle =1
                    TextFontCharSet =163
                    TextAlign =3
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =10649
                    Width =1152
                    FontSize =9
                    TabIndex =4
                    BorderColor =-2147483615
                    ForeColor =-2147483615
                    Name ="Text248"
                    ControlSource ="=Sum([TaxAmt])"
                    Format ="$#,##0.00;($#,##0.00)"
                    ControlTipText ="TaxAmt Total"

                    LayoutCachedLeft =10649
                    LayoutCachedWidth =11801
                    LayoutCachedHeight =240
                End
                Begin TextBox
                    OldBorderStyle =1
                    TextFontCharSet =186
                    TextAlign =3
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =3705
                    Width =4050
                    FontSize =9
                    FontWeight =700
                    BorderColor =-2147483615
                    ForeColor =-2147483615
                    Name ="Text136"
                    ControlSource ="=Sum([UnitCount]) & IIf(Sum([UnitCount])>1,\" units at \",\" unit at \") & [Loca"
                        "tionName] & \" - Sub Total ►\""

                    LayoutCachedLeft =3705
                    LayoutCachedWidth =7755
                    LayoutCachedHeight =240
                End
                Begin TextBox
                    OldBorderStyle =1
                    TextFontCharSet =163
                    TextAlign =3
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =12853
                    Width =1335
                    FontSize =9
                    TabIndex =2
                    BorderColor =-2147483615
                    ForeColor =-2147483615
                    Name ="Text235"
                    ControlSource ="=Sum([LineTot])"
                    Format ="$#,##0.00;($#,##0.00)"
                    ControlTipText ="LineTot Total"

                    LayoutCachedLeft =12853
                    LayoutCachedWidth =14188
                    LayoutCachedHeight =240
                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextFontCharSet =186
                    TextAlign =2
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =9150
                    Width =1155
                    FontSize =7
                    FontWeight =700
                    TabIndex =5
                    BorderColor =-2147483615
                    ForeColor =8355711
                    Name ="Text215"
                    ControlSource ="=\" Tax Type: \" & [sttaxtype] & \"\""
                    TextFormat =1

                    LayoutCachedLeft =9150
                    LayoutCachedWidth =10305
                    LayoutCachedHeight =240
                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =285
            Name ="GroupFooter2"
            Begin
                Begin TextBox
                    TextFontCharSet =163
                    TextAlign =3
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Left =3705
                    Top =45
                    Width =4050
                    FontSize =9
                    FontWeight =700
                    Name ="Text138"
                    ControlSource ="=Sum([UnitCount]) & IIf(Sum([UnitCount])>1,\" units\",\" unit\") & \" - Group \""
                        " & [sgrpno] & \" - Sub Total ►\""

                    LayoutCachedLeft =3705
                    LayoutCachedTop =45
                    LayoutCachedWidth =7755
                    LayoutCachedHeight =285
                End
                Begin TextBox
                    OldBorderStyle =1
                    TextFontCharSet =163
                    TextAlign =3
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =12868
                    Top =45
                    Width =1320
                    FontSize =9
                    TabIndex =1
                    BackColor =8210719
                    BorderColor =-2147483615
                    ForeColor =16777215
                    Name ="AccessTotalsLineTot1"
                    ControlSource ="=Sum([LineTot])"
                    Format ="$#,##0.00;($#,##0.00)"
                    ControlTipText ="LineTot Total"

                    LayoutCachedLeft =12868
                    LayoutCachedTop =45
                    LayoutCachedWidth =14188
                    LayoutCachedHeight =285
                End
                Begin TextBox
                    OldBorderStyle =1
                    TextFontCharSet =163
                    TextAlign =3
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =10648
                    Top =45
                    Width =1152
                    FontSize =9
                    TabIndex =2
                    BackColor =8210719
                    BorderColor =-2147483615
                    ForeColor =16777215
                    Name ="Text187"
                    ControlSource ="=Sum([taxamt])"
                    Format ="$#,##0.00;($#,##0.00)"
                    ControlTipText ="LineTot Total"

                    LayoutCachedLeft =10648
                    LayoutCachedTop =45
                    LayoutCachedWidth =11800
                    LayoutCachedHeight =285
                End
                Begin TextBox
                    OldBorderStyle =1
                    TextFontCharSet =163
                    TextAlign =3
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =7845
                    Top =45
                    Width =1110
                    FontSize =9
                    TabIndex =3
                    BackColor =8210719
                    BorderColor =-2147483615
                    ForeColor =16777215
                    Name ="Text188"
                    ControlSource ="=Sum([unitsbillrent])"
                    Format ="$#,##0.00;($#,##0.00)"
                    ControlTipText ="LineTot Total"

                    LayoutCachedLeft =7845
                    LayoutCachedTop =45
                    LayoutCachedWidth =8955
                    LayoutCachedHeight =285
                End
            End
        End
        Begin PageFooter
            Height =240
            Name ="PageFooterSection"
            Begin
                Begin TextBox
                    TextFontCharSet =161
                    TextAlign =1
                    TextFontFamily =34
                    IMESentenceMode =3
                    Width =4620
                    FontSize =9
                    Name ="Text128"
                    ControlSource ="=\"printed on \" & Date() & \"@ \" & Time()"

                    LayoutCachedWidth =4620
                    LayoutCachedHeight =240
                End
                Begin TextBox
                    TextFontCharSet =161
                    TextAlign =3
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Left =10740
                    Width =3600
                    FontSize =9
                    TabIndex =1
                    Name ="Text130"
                    ControlSource ="=\"Page \" & [Page] & \" of \" & [Pages]"

                    LayoutCachedLeft =10740
                    LayoutCachedWidth =14340
                    LayoutCachedHeight =240
                End
            End
        End
        Begin FormFooter
            KeepTogether = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =4380
            Name ="ReportFooter"
            Begin
                Begin TextBox
                    TextFontCharSet =186
                    TextAlign =3
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =3705
                    Top =45
                    Width =4050
                    FontSize =9
                    FontWeight =700
                    ForeColor =13107
                    Name ="Text133"
                    ControlSource ="=Sum([UnitCount]) & IIf(Sum([UnitCount])>1,\" units\",\" unit\") & \" - Schedule"
                        " \" & [schno] & \" - Invoice Total ►\""

                    LayoutCachedLeft =3705
                    LayoutCachedTop =45
                    LayoutCachedWidth =7755
                    LayoutCachedHeight =285
                End
                Begin Line
                    OldBorderStyle =2
                    BorderWidth =3
                    BorderLineStyle =1
                    Top =1290
                    Width =14400
                    Name ="Line147"
                    LayoutCachedTop =1290
                    LayoutCachedWidth =14400
                    LayoutCachedHeight =1290
                End
                Begin TextBox
                    OldBorderStyle =1
                    TextFontCharSet =163
                    TextAlign =3
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =10648
                    Top =45
                    Width =1152
                    FontSize =9
                    TabIndex =1
                    BackColor =13107
                    BorderColor =13107
                    ForeColor =16777215
                    Name ="AccessTotalsTaxAmt"
                    ControlSource ="=Sum([TaxAmt])"
                    Format ="$#,##0.00;($#,##0.00)"
                    ControlTipText ="TaxAmt Total"

                    LayoutCachedLeft =10648
                    LayoutCachedTop =45
                    LayoutCachedWidth =11800
                    LayoutCachedHeight =285
                End
                Begin TextBox
                    OldBorderStyle =1
                    TextFontCharSet =163
                    TextAlign =3
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =12868
                    Top =50
                    Width =1320
                    FontSize =9
                    TabIndex =2
                    BackColor =13107
                    BorderColor =13107
                    ForeColor =16777215
                    Name ="AccessTotalsLineTot"
                    ControlSource ="=Sum([LineTot])"
                    Format ="$#,##0.00;($#,##0.00)"
                    ControlTipText ="LineTot Total"

                    LayoutCachedLeft =12868
                    LayoutCachedTop =50
                    LayoutCachedWidth =14188
                    LayoutCachedHeight =290
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    TextFontCharSet =163
                    TextAlign =3
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =12375
                    Top =1140
                    Height =60
                    FontSize =10
                    FontWeight =700
                    TabIndex =3
                    Name ="Text196"
                    ControlSource ="=IIf(IsNull([sumofbilladj]),Null,[sumofbilladj]+Sum([LineTot]))"
                    Format ="$#,##0.00;($#,##0.00)"

                    LayoutCachedLeft =12375
                    LayoutCachedTop =1140
                    LayoutCachedWidth =13815
                    LayoutCachedHeight =1200
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    TextFontCharSet =161
                    TextAlign =3
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =8460
                    Top =1125
                    Width =3225
                    Height =60
                    FontSize =10
                    FontWeight =700
                    TabIndex =4
                    Name ="Text197"
                    ControlSource ="=IIf(IsNull([sumofbilladj]),Null,\"Total Due including Adjustments\")"

                    LayoutCachedLeft =8460
                    LayoutCachedTop =1125
                    LayoutCachedWidth =11685
                    LayoutCachedHeight =1185
                End
                Begin Subform
                    Locked = NotDefault
                    CanShrink = NotDefault
                    Left =2085
                    Top =465
                    Width =10080
                    Height =659
                    TabIndex =5
                    Name ="AdjSubrpt"
                    SourceObject ="Report.AdjSubrpt"
                    LinkChildFields ="BillAdjBillref"
                    LinkMasterFields ="BillID"

                    LayoutCachedLeft =2085
                    LayoutCachedTop =465
                    LayoutCachedWidth =12165
                    LayoutCachedHeight =1124
                    Begin
                        Begin Label
                            Visible = NotDefault
                            Left =2085
                            Top =225
                            Width =1005
                            Height =315
                            Name ="AdjSubrpt Label"
                            Caption ="AdjSubrpt"
                            EventProcPrefix ="AdjSubrpt_Label"
                            LayoutCachedLeft =2085
                            LayoutCachedTop =225
                            LayoutCachedWidth =3090
                            LayoutCachedHeight =540
                        End
                    End
                End
                Begin TextBox
                    OldBorderStyle =1
                    TextFontCharSet =163
                    TextAlign =3
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =7845
                    Top =45
                    Width =1110
                    FontSize =9
                    TabIndex =6
                    BackColor =13107
                    BorderColor =13107
                    ForeColor =16777215
                    Name ="Text213"
                    ControlSource ="=Sum([unitsbillrent])"
                    Format ="$#,##0.00;($#,##0.00)"
                    ControlTipText ="LineTot Total"

                    LayoutCachedLeft =7845
                    LayoutCachedTop =45
                    LayoutCachedWidth =8955
                    LayoutCachedHeight =285
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    TextFontCharSet =163
                    TextAlign =3
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =12375
                    Top =1320
                    Height =60
                    FontSize =10
                    FontWeight =700
                    TabIndex =7
                    Name ="Text214"
                    ControlSource ="=IIf(IsNull([sumofbilladj]),Null,[sumofbilladj]+Sum([LineTot]))"
                    Format ="$#,##0.00;($#,##0.00)"

                    LayoutCachedLeft =12375
                    LayoutCachedTop =1320
                    LayoutCachedWidth =13815
                    LayoutCachedHeight =1380
                End
                Begin Rectangle
                    BackStyle =0
                    BorderWidth =1
                    Left =2565
                    Top =1605
                    Width =8415
                    Height =2700
                    Name ="Box52"
                    LayoutCachedLeft =2565
                    LayoutCachedTop =1605
                    LayoutCachedWidth =10980
                    LayoutCachedHeight =4305
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    TextFontCharSet =186
                    TextAlign =3
                    TextFontFamily =34
                    Left =900
                    Top =1515
                    Width =2355
                    Height =315
                    FontSize =12
                    FontWeight =700
                    BorderColor =-2147483617
                    ForeColor =-2147483615
                    Name ="Label239"
                    Caption ="Payment Instructions:"
                    LayoutCachedLeft =900
                    LayoutCachedTop =1515
                    LayoutCachedWidth =3255
                    LayoutCachedHeight =1830
                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =5130
                    Top =1680
                    Width =5550
                    Height =315
                    TabIndex =8
                    Name ="Text240"
                    ControlSource ="=IIf([BILLTO]<=[AssignmentDate],\"Wells Fargo Bank, N.A\",\"Fifth Third Bank, Ci"
                        "ncinnati, OH\")"

                    LayoutCachedLeft =5130
                    LayoutCachedTop =1680
                    LayoutCachedWidth =10680
                    LayoutCachedHeight =1995
                End
                Begin Label
                    TextFontCharSet =163
                    TextAlign =3
                    TextFontFamily =34
                    Left =3675
                    Top =1680
                    Width =1320
                    Height =315
                    FontSize =12
                    FontWeight =700
                    Name ="Label241"
                    Caption ="Bank Name:"
                    LayoutCachedLeft =3675
                    LayoutCachedTop =1680
                    LayoutCachedWidth =4995
                    LayoutCachedHeight =1995
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    IMESentenceMode =3
                    Left =5130
                    Top =2070
                    Width =5550
                    Height =345
                    TabIndex =9
                    Name ="Text242"
                    ControlSource ="=IIf([BILLTO]<=[AssignmentDate],\"Fleet Advantage, LLC\",\"Fifth Third Equipment"
                        " Finance Company\")"

                    LayoutCachedLeft =5130
                    LayoutCachedTop =2070
                    LayoutCachedWidth =10680
                    LayoutCachedHeight =2415
                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =5130
                    Top =3015
                    Width =5550
                    Height =315
                    TabIndex =10
                    Name ="Text244"
                    ControlSource ="=IIf([BILLTO]<=[AssignmentDate],\"2000037221232\",\"790-56487\")"

                    LayoutCachedLeft =5130
                    LayoutCachedTop =3015
                    LayoutCachedWidth =10680
                    LayoutCachedHeight =3330
                End
                Begin Label
                    TextFontCharSet =163
                    TextAlign =3
                    TextFontFamily =34
                    Left =2595
                    Top =3015
                    Width =2400
                    Height =315
                    FontSize =12
                    FontWeight =700
                    Name ="Label245"
                    Caption ="Account Number:"
                    LayoutCachedLeft =2595
                    LayoutCachedTop =3015
                    LayoutCachedWidth =4995
                    LayoutCachedHeight =3330
                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =5130
                    Top =2520
                    Width =5550
                    Height =315
                    TabIndex =11
                    Name ="Text246"
                    ControlSource ="=IIf([BILLTO]<=[AssignmentDate],\"063107513\",\"042000314\")"

                    LayoutCachedLeft =5130
                    LayoutCachedTop =2520
                    LayoutCachedWidth =10680
                    LayoutCachedHeight =2835
                End
                Begin Label
                    TextFontCharSet =163
                    TextAlign =3
                    TextFontFamily =34
                    Left =3135
                    Top =2535
                    Width =1860
                    Height =315
                    FontSize =12
                    FontWeight =700
                    Name ="Label247"
                    Caption ="ABA Routing No.:"
                    LayoutCachedLeft =3135
                    LayoutCachedTop =2535
                    LayoutCachedWidth =4995
                    LayoutCachedHeight =2850
                End
                Begin TextBox
                    TextFontCharSet =186
                    TextAlign =3
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =2655
                    Top =3495
                    Width =7995
                    Height =315
                    TabIndex =12
                    Name ="Text220"
                    ControlSource ="=\"<i>Note: Please reference \" & IIf([billto]<[AssignmentDate],\" invoice numbe"
                        "r: \",\" Bank Contract number: \") & \"</i><Strong>\" & IIf([billto]<[Assignment"
                        "Date],[Billto],[schbnkclno] & \"-\" & \"109\") & \"</Strong>\""
                    TextFormat =1

                    LayoutCachedLeft =2655
                    LayoutCachedTop =3495
                    LayoutCachedWidth =10650
                    LayoutCachedHeight =3810
                End
                Begin Label
                    TextFontCharSet =163
                    TextAlign =3
                    TextFontFamily =34
                    Left =3345
                    Top =2115
                    Width =1650
                    Height =315
                    FontSize =12
                    FontWeight =700
                    Name ="Label228"
                    Caption ="Account Name:"
                    LayoutCachedLeft =3345
                    LayoutCachedTop =2115
                    LayoutCachedWidth =4995
                    LayoutCachedHeight =2430
                End
                Begin TextBox
                    TextFontCharSet =186
                    TextAlign =3
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =2640
                    Top =3885
                    Width =8025
                    Height =315
                    TabIndex =13
                    Name ="Text263"
                    ControlSource ="=\"<i>and Bank Contract number: </i><Strong>\" & [schbnkclno] & \"-117</Strong>\""
                    TextFormat =1

                    LayoutCachedLeft =2640
                    LayoutCachedTop =3885
                    LayoutCachedWidth =10665
                    LayoutCachedHeight =4200
                End
                Begin TextBox
                    OldBorderStyle =1
                    TextFontCharSet =186
                    TextAlign =3
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =11820
                    Top =1890
                    Width =2385
                    Height =330
                    FontSize =14
                    FontWeight =700
                    TabIndex =14
                    Name ="Text253"
                    ControlSource ="=[BillNo]"

                    LayoutCachedLeft =11820
                    LayoutCachedTop =1890
                    LayoutCachedWidth =14205
                    LayoutCachedHeight =2220
                    Begin
                        Begin Label
                            BackStyle =1
                            OldBorderStyle =1
                            TextFontCharSet =161
                            TextAlign =1
                            TextFontFamily =34
                            Left =11820
                            Top =1620
                            Width =2385
                            Height =345
                            FontWeight =700
                            BackColor =8210719
                            BorderColor =-2147483615
                            ForeColor =16777215
                            Name ="Label254"
                            Caption ="Invoice Number:"
                            LayoutCachedLeft =11820
                            LayoutCachedTop =1620
                            LayoutCachedWidth =14205
                            LayoutCachedHeight =1965
                        End
                    End
                End
                Begin Label
                    OldBorderStyle =1
                    TextFontCharSet =161
                    TextAlign =1
                    TextFontFamily =34
                    Left =11820
                    Top =3015
                    Width =2385
                    Height =315
                    FontWeight =700
                    Name ="Label259"
                    Caption ="Total Rent:"
                    LayoutCachedLeft =11820
                    LayoutCachedTop =3015
                    LayoutCachedWidth =14205
                    LayoutCachedHeight =3330
                End
                Begin TextBox
                    OldBorderStyle =1
                    TextFontCharSet =186
                    TextAlign =3
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =11820
                    Top =3270
                    Width =2385
                    Height =435
                    FontSize =16
                    FontWeight =700
                    TabIndex =15
                    Name ="Text221"
                    ControlSource ="=Sum([LineTot])"
                    Format ="$#,##0.00;($#,##0.00)"
                    ControlTipText ="LineTot Total"

                    LayoutCachedLeft =11820
                    LayoutCachedTop =3270
                    LayoutCachedWidth =14205
                    LayoutCachedHeight =3705
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    TextFontCharSet =161
                    TextAlign =1
                    TextFontFamily =34
                    Left =11820
                    Top =2280
                    Width =2385
                    Height =315
                    FontWeight =700
                    BackColor =15527148
                    Name ="Label273"
                    Caption ="Total Taxes:"
                    LayoutCachedLeft =11820
                    LayoutCachedTop =2280
                    LayoutCachedWidth =14205
                    LayoutCachedHeight =2595
                End
                Begin TextBox
                    OldBorderStyle =1
                    TextFontCharSet =186
                    TextAlign =3
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =11820
                    Top =2535
                    Width =2385
                    Height =360
                    FontSize =14
                    FontWeight =700
                    TabIndex =16
                    Name ="Text274"
                    ControlSource ="=Sum([TaxAmt])"
                    Format ="$#,##0.00;($#,##0.00)"
                    ControlTipText ="LineTot Total"

                    LayoutCachedLeft =11820
                    LayoutCachedTop =2535
                    LayoutCachedWidth =14205
                    LayoutCachedHeight =2895
                End
            End
        End
    End
End
CodeBehindForm
' See "FABankBill_Toyota10.cls"
