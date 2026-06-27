Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    TabularFamily =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =14430
    DatasheetFontHeight =10
    ItemSuffix =117
    Left =2955
    Top =2295
    Right =17640
    Bottom =13035
    Filter ="[groupID]=1153"
    ShortcutMenuBar ="SortClip"
    RecSrcDt = Begin
        0x9d8a2d7a701ee440
    End
    RecordSource ="groups_viewQuery"
    Caption ="View/Edit Groups"
    OnCurrent ="[Event Procedure]"
    DatasheetFontName ="Arial"
    OnLoad ="[Event Procedure]"
    FilterOnLoad =0
    NavigationCaption ="Groups"
    ShowPageMargins =0
    DatasheetGridlinesColor12 =12632256
    Begin
        Begin Label
            BackStyle =0
            FontName ="Tahoma"
        End
        Begin Rectangle
            SpecialEffect =3
            BackStyle =0
            BorderLineStyle =0
        End
        Begin Image
            BackStyle =0
            OldBorderStyle =0
            BorderLineStyle =0
            SizeMode =3
            PictureAlignment =2
        End
        Begin CommandButton
            FontSize =8
            FontWeight =400
            ForeColor =-2147483630
            FontName ="Tahoma"
            BorderLineStyle =0
        End
        Begin CheckBox
            SpecialEffect =2
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
        End
        Begin TextBox
            FELineBreak = NotDefault
            SpecialEffect =2
            OldBorderStyle =0
            BorderLineStyle =0
            FontName ="Tahoma"
            AsianLineBreak =255
        End
        Begin ComboBox
            BorderLineStyle =0
            LabelX =-1800
            FontSize =11
            BorderColor =12632256
            FontName ="Calibri"
            AllowValueListEdits =1
            InheritValueList =1
        End
        Begin Subform
            SpecialEffect =2
            BorderLineStyle =0
        End
        Begin Tab
            FontName ="Tahoma"
            BorderLineStyle =0
        End
        Begin FormHeader
            Height =1380
            BackColor =15707392
            Name ="FormHeader"
            Begin
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    SpecialEffect =0
                    OverlapFlags =85
                    BackStyle =0
                    IMESentenceMode =3
                    Left =2640
                    Top =255
                    Width =9300
                    Height =480
                    FontSize =18
                    ForeColor =16777215
                    Name ="CompanyAndDBA"
                    ControlSource ="CompanyAndDBA"

                    LayoutCachedLeft =2640
                    LayoutCachedTop =255
                    LayoutCachedWidth =11940
                    LayoutCachedHeight =735
                End
                Begin TextBox
                    Visible = NotDefault
                    SpecialEffect =0
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =13380
                    Top =540
                    Width =660
                    Height =255
                    TabIndex =1
                    ForeColor =255
                    Name ="deal Ref"
                    ControlSource ="dealRef"
                    StatusBarText ="link to Deals table"
                    EventProcPrefix ="deal_Ref"

                    LayoutCachedLeft =13380
                    LayoutCachedTop =540
                    LayoutCachedWidth =14040
                    LayoutCachedHeight =795
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =12525
                    Width =1320
                    Height =390
                    FontSize =9
                    FontWeight =700
                    TabIndex =2
                    ForeColor =-2147483615
                    Name ="CloseFormBtn"
                    Caption ="Save && Close"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =12525
                    LayoutCachedWidth =13845
                    LayoutCachedHeight =390
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =0
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =10034
                    Top =795
                    Width =720
                    Height =285
                    FontSize =10
                    TabIndex =3
                    BackColor =15707392
                    ForeColor =16777215
                    Name ="dealSchedNum"
                    ControlSource ="dealSchedNum"

                    LayoutCachedLeft =10034
                    LayoutCachedTop =795
                    LayoutCachedWidth =10754
                    LayoutCachedHeight =1080
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =85
                            TextAlign =3
                            Left =8700
                            Top =795
                            Width =1275
                            Height =285
                            FontSize =10
                            FontWeight =700
                            BackColor =15707392
                            ForeColor =16777215
                            Name ="Label90"
                            Caption ="Schedule #:"
                            LayoutCachedLeft =8700
                            LayoutCachedTop =795
                            LayoutCachedWidth =9975
                            LayoutCachedHeight =1080
                        End
                    End
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    SpecialEffect =0
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =2640
                    Top =795
                    Width =5970
                    Height =285
                    FontSize =10
                    TabIndex =4
                    BackColor =15707392
                    ForeColor =16777215
                    Name ="dealShortDesc"
                    ControlSource ="dealShortDesc"

                    LayoutCachedLeft =2640
                    LayoutCachedTop =795
                    LayoutCachedWidth =8610
                    LayoutCachedHeight =1080
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =0
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =11850
                    Top =795
                    Width =945
                    Height =285
                    FontSize =10
                    TabIndex =5
                    BackColor =15707392
                    ForeColor =16777215
                    Name ="Text94"
                    ControlSource ="groupNo"

                    LayoutCachedLeft =11850
                    LayoutCachedTop =795
                    LayoutCachedWidth =12795
                    LayoutCachedHeight =1080
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =85
                            TextAlign =3
                            Left =10815
                            Top =795
                            Width =975
                            Height =285
                            FontSize =10
                            FontWeight =700
                            BackColor =15707392
                            ForeColor =16777215
                            Name ="Label95"
                            Caption ="Group #:"
                            LayoutCachedLeft =10815
                            LayoutCachedTop =795
                            LayoutCachedWidth =11790
                            LayoutCachedHeight =1080
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    SpecialEffect =0
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =13335
                    Top =930
                    Width =750
                    Height =255
                    TabIndex =6
                    ForeColor =255
                    Name ="clientRef"
                    ControlSource ="clientRef"
                    StatusBarText ="link to clients table"

                    LayoutCachedLeft =13335
                    LayoutCachedTop =930
                    LayoutCachedWidth =14085
                    LayoutCachedHeight =1185
                End
                Begin Image
                    BackStyle =1
                    Left =15
                    Top =75
                    Width =2289
                    Height =1194
                    Name ="Auto_Logo0"
                    Picture ="FA_Logo Trans Small.png"
                    ImageData = Begin
                        0x89504e470d0a1a0a0000000d49484452000000a00000005108060000005501ad ,
                        0x91000000206348524d00007a25000080830000f9ff000080e9000075300000ea ,
                        0x6000003a980000176f925fc546000000097048597300002e2200002e2201aae2 ,
                        0xdd9200001a5249444154785eed5d09781365fa1f2aa0a0ab0bea2ac87a728b78 ,
                        0x71ca62299682a8c861b9a502b5d096243399496f08b45ca54dd20bb080dc0854 ,
                        0x6eca5d68699216a4522e5910565474d575571f9fdd6557dd3fdffff77e9369d3 ,
                        0x3651cb8a66e9ccf37c4f26c977bcdfeffd7deff1cd642208faa123a023a023a0 ,
                        0x23a023a023a023a023a023a023a023a0231058080c5d7ffa919b676d0b11a6e6 ,
                        0xf4bd455ce45596f46d36c5d6a3cd2bd6968125b12ecd0d85c0e48d6706bdb0a6 ,
                        0x92b5b21d60c2ac9dac515a6155094a2e6041e2b20b4d221764fe76b8fcc00d35 ,
                        0x717d32018340a36ecbde59155df8116b9177940959479890ed551c654c48dbc5 ,
                        0x82a26c979b8e945f0a18a975416e2004e277dc1fbaf2ddcfc2379c6582ddcd04 ,
                        0x221d952c94ec729590e945ac5164c6959b872ba137d0ccf5a9040a02b765144f ,
                        0x9dbce57dd629ff981709898c1e421211e7ec6541afa65e6a1eaedc1b2872eb72 ,
                        0xdc28088487dff4e8e2f29229db2eb05bc8f26956d0fb350b244c5ccf9abc92b0 ,
                        0xe84699b63e8f404260f6fe6e2faf3d7965c0ea531e2b08c2d52660e661264466 ,
                        0x7cd76c585caf40125d97e50641e0f68ce33322377dc05ae75532c15601227a0a ,
                        0x9d5371bccb04eb1e268c4c2bba41a6ac4f23501030ecea1e39715d882b62d5c4 ,
                        0xabe3564e6483964e60cf579508cf79041bb464021b681fc2262cfb43e9ebab7a ,
                        0x1546adeebdcb6f59d37b57e4aa5ebb5f9edd65df4bd627bb06ca5c7539020c81 ,
                        0xc4fdfd82579d8cffcef16e388bdbdf8d251ee8c1a617f56429073c85ceb5f745 ,
                        0xbd584a711f96e2eccfa6bb9f63292edf45fb2ef2ad1e6c407cbb0dbda436cd02 ,
                        0x6cdaba3881804044e1ddf7ae3d9578bef0031bb33ac3584a69ff9fa13cc7a683 ,
                        0x98b19b9f610312da9f088d7afa8e4098ab2e43e021d0687145f4a68397df6069 ,
                        0xae975862493f967cb8ff7f574060229f79575f367856c7bf87183b3e1d78d3d6 ,
                        0x250a0804d2dde1f1073e5ac4b2de99c0128a9ffdaf88479673ba13ee184502f9 ,
                        0x5e9ef7e8f72162fb570362a2ba10818740f2c190fe5bcf2df876d569852594f4 ,
                        0xad43be9452d58d6aa42262fd50490601e30e04b3c8b53dd80bd68e179f35b61b ,
                        0x1678b3d6250a0804a6ec6c73dfda534917b75f5c00e285b024146fd74bd62cfe ,
                        0x403f16bba9378b2ee8cd5ffd95187c37657d4f366ee193df0e9ddba53234ae43 ,
                        0x72cfc90fdd131013d585084004ac4210e2beed459717b354d70b88fb82eb902f ,
                        0xe150081b93f7040b4d68bf0071dc0b03ad2d4347ceba3564c29c3b7c9671d36f ,
                        0xeb3b3e41681b80b3d5450a3404d2dda3538a3e5ecc1cef8cf719f791eb2537da ,
                        0x5f6eb78964678c35095b7e3eafb175ff4a41deb242b06cf355560a5397ae1246 ,
                        0xa43e1a68f3d5e5092004120e868421eefb7ec529d967dc371d7b7bc66d7dd8a0 ,
                        0x940e1f84c676684da21714b09b06afae5c397ed765d662e13beadd3139b8752b ,
                        0x477bc5792e3e4fdbcd1a8f49392458ad410134655d944041207aef23bf5f773a ,
                        0xe5d2b68be92ca9a46edc474984bcf759366476e7effa1ada86d590dbb8fd9e01 ,
                        0x2b2b3f1c5d70ce73ab16ae13d3cd095aa1bb65ec2e2618f3d9cd2394d70365ce ,
                        0xba1c818240b870d31b15b1bb28ee9be51cec77bf2f6259b7fff48d7d44f12576 ,
                        0x13fbd17193b79c638fe6d3f561affb06b51b168884f3f6b3a088b44f9a0db170 ,
                        0xeba91f3a021c81ccb2b1568afb6c47c7c2f5fadeef4b2ae9cfa6acebf9fdab6f ,
                        0x3c553c7945cf7d912b7b1651797d75cfa289cbba1d0c4b6eb7fdc9bc1da7a66c ,
                        0xb9f823b76c6d608d47262cd1a1d711e008a494840ede7e3ee33fcb4f89483aea ,
                        0xeef7796fbf1009933dd778e95aeef4b2e758125cf3f8254fb181721b8b907ae4 ,
                        0xf1a16b4ffe6bc0ea93b83bc6871524978c5bb61a45d9fe73eb70294457410347 ,
                        0x20baf0b187d79d99f1f1960b7311f7d165b69afb7d752ebbf1ebc0d870a657da ,
                        0x78c6f9a4d5dd597fa5fd5a0dca7bb39de9d13b2eb2d60b9184f872c544c2d442 ,
                        0x163476fab10782adb734701534dce9b7cd166e5e5a612cd9ff711e9b593aa85e ,
                        0xd779390171158436990724b43bd94bea5cfdf34cc3eedb7b2f3b7af6b54de759 ,
                        0x9076cb7e8d1b57e98e6a584779056b1a6e896fb81a68e033b7978dcf3cf8493e ,
                        0xcb281fe937eeab61013dd68f2ebf51463c6deb33ecf9191dbf4052f2541d28d3 ,
                        0x0f0e19bff10cebfee671ffae38fd206b3471fe574d871bdb35705534bce9cf2a ,
                        0x79217ce7fbf6ab4b2a6359bc8fb88fdc2b11adf6b55dda9e9176f765e3f39fa2 ,
                        0xbdc0c3c1b11dbaf943af6d5e5941ecf63fb1df50f6cb2da0f623264f6c48ae38 ,
                        0x65336b3c2a996f68eb470341c0b4e7b10e1bde9bf5c5a6f3a9752eb391c523f7 ,
                        0x6ad917ccaff14edd40d77a9f61d11b7bb349cbbbb3d1d98f7ff6626ae7c21053 ,
                        0xbb319dc33b37fd41c852f7b71bbce6c4df86ac3b53d70a6a3f6cb29532212697 ,
                        0x351b1137bc81c0dfb0a7d9c626345b7a5c74eefb2897594b07d689fbf84d0645 ,
                        0xfd58b8a32b0b4be8943c24ed89d0e1b3bb0e1c96d675d08bc95dff106aec7c7f ,
                        0x7d106c613f9c3075db45f6e0625c0da9fd9b62fedb6258c1d97b58d0f899e7ef ,
                        0x181bdda23e7deb75ff0711c82a9fe0a0b82fbd3c1ce4f3bddff7da8a6e2cc4dc ,
                        0x6ed9cf32bda81dcd7b2c3d5249bf296eec2b21d19293b8b5acc9c8b8f93fcb98 ,
                        0x7a278189c0cce2c1630a2f38587e658ccffd3efa0d47f4db3cab3dfeb3de22bf ,
                        0xe060d8a8f5a7aef659815fd2d5d996413c48567041316b3439fd4af36162dd84 ,
                        0x2630e1d4a5aa0f0286dd4f74def85eea976f9f9be933eea38483ee521e3cb3e3 ,
                        0xd7c131ed9ea84fdf3fa5ee3d8e9295a3375d606d16d672c5de97e9666e634163 ,
                        0x928bf49b157e0aa2ff43759ece179abf795c2a53e3be01b8b9b4fa771dda5dcd ,
                        0xca9e67d9d0f95d5888b1edd8eb32b5c48d77ffce5e5211bcfa2c6b450f39aaf1 ,
                        0x6405b28288073d372b3479459e7c5d64d03bfd751078e378746ef1674b58c6d1 ,
                        0x7096e40c566fa3f7dc4a9f886d15dacf1b32e7d16ffa496dafafe2e3d73f786b ,
                        0xfa81f287f32bd99d0b2b5850d553b640487e0b176ed9c2de60d06b732fdf395a ,
                        0xbf59e1d761cbcf3ceaf2e3e6c8b56713d8ec9291cc54f80c77b354c4c2be7c6b ,
                        0x656456d72b83523a16f8dc4cfe9965e1dd452f6a21ccd89add346de73741209b ,
                        0x301fcf1d9ce755e87dfc3aec0d26bd2544e537b91e22e87dfe4208e457443549 ,
                        0x778e8f9ab4bce7ac97e6b69d3e6241175e865199df65fa8ba98f45f637b6ff55 ,
                        0x9e46d0d4b4b4536353fe143c79354598e248ae2ad1d9c9c254474ad0a894f83b ,
                        0x22acbffd85a0d287d111d011d011d011d011d011d011d011d011d011d011d011 ,
                        0xd011d011d011d01168a00888a21c65962cc59228978926a55c34c99e527d6e96 ,
                        0xe28e49263935ce10d74632291b45517c38d0e1124de64cb3d97214f2be6d3299 ,
                        0xae699fcd6834b6134565bdd1185faf5bb3021d9b80924f9294c520e05510f10d ,
                        0x49b2e400f0dcdac56c8e5b281ae508498a7b0c04fd97187b7defe49065f921c8 ,
                        0x127cad40499234283e2e89615e2e458e63a2d13ceb5afa9a364dea81f9fe5314 ,
                        0xe302e2311bb1b1b1adc5694aa8d56a6d7c2df309c836b06c209d7ceea7082749 ,
                        0xf18fc34a7e214d931eff29f5afb58e24c9d330ce1fafa53d94d314d6fc24cad9 ,
                        0x70fce50216532688f82f799afc507dfb9362255a709f994c719deadbf67ad437 ,
                        0x99cc4361d93fb30a37d0233f340292b27e0cb4da04c4c37a1a4946cb60106626 ,
                        0xfac90238b926101ad62bc56834ffc1bb3f83c170bb6454c6a3ee3cd1a8649bb9 ,
                        0xa5951db21c2f4d339b3b525d58be07cc66254132990b49f1b05c49e837918a28 ,
                        0xe25c92658321f1ee1f92134a8a26eb271ad4bfd38a8951ee358b96bf994df21a ,
                        0x5fed602d5bc24d4f40df73309e0384cd26f70dcff0ba6810c3b0103e1645cb53 ,
                        0x66936592c9a43ce16f6cb339a13de434c6c4586f8b8f8abf4332caa365c93203 ,
                        0x72cfa78205315b162d530d06e511ad0f8cd547342861d41678c4418e2cbc2ec4 ,
                        0x980ba9ad6c927b52dd4463e23df85cc1e7ebe189be467fb3d01638a9d8f02249 ,
                        0xf751dd7853fc838a14f73abe4fd3c6465f33713e8ee6ea2d3f42a94715b36226 ,
                        0x3d48264b1ee90fe3cc87fea053a917e63c44ab1f1717d786e447bf0b50271b63 ,
                        0xe6a1cf4c60134518ff1877fc7e4f9d41804b66b3b9bd01311ec579f4ea7d8eef ,
                        0x7e0f02ddac28895dc9022a46a52b115612953765b3e5ef50de3e08bf1c4acb07 ,
                        0x0156e3b5189f7f6316cd561a186def06b1ca51ff4b08be0d65a959949710a092 ,
                        0x68a9304bcad76837c41c6b6e8fb61b41bee3a80ba0e50d5aa1cfa1c4150484bf ,
                        0xc9a8e3583e479b3dde75401cd1a224308412cf787f8e18ef7eccff14fafd1c65 ,
                        0x2be6b00c0a5e82f62bf0f921c8f7017dc7634193b29c3c05e1507b7c5a889ed8 ,
                        0xf9901164419b4acce973941d344793515e4bf3469f17308faf0da2c89f37438a ,
                        0x8775fe026d4f41e9c520e04a947c94d5f8fc1d60f82d7019433a41dd35c0c48d ,
                        0xfefe01c56fa6d894b0c158bc505c8ec5df13e7c05d392ba10ecedf425f6bb100 ,
                        0xf710f6e8f3242d721adb6854c602fb7f402fa7284ec6f90a2cfc95a8b30d754f ,
                        0xa3fc15a598ea42ff4f42fe3fa3ce45ae1fe88e748873b4932fa25c226f714d24 ,
                        0x44631b267a15e07f818e3f23c0f1ca8b768e15f4a56c8a7f16827424029aa799 ,
                        0x3b1a8df2d300ee2a26ddcfd7c0b218f70adaff9b260c82c6d084a2cce6bb7cd5 ,
                        0xe52b5094dfd7948bf7538984f59d10149d219be39852cb52613537437fef4369 ,
                        0x4eb8e8aa2756c1e2cd07e07f86f25af9964b99cf635e28579e2aff8e16051695 ,
                        0xa9765d28781c14f2bdc5626987f976c6f9bb9aa26bd7451f7bf1fd0195800a8d ,
                        0xff4fc9e0fbc9097c81a2af2a8b09ab0e792e168417f8f456269334056d76f99a ,
                        0x0b7477175f1426c50c0c6e01e93e24038287ce35aa5d3f3a3aa1052d221092cb ,
                        0x09635180718ffad30764741189ebab2f5e9f5c2706fb1402be084042fd943004 ,
                        0xc077aa4988f285c964e940ee14c05f417d2b1459d31d98e0574df23c728570c5 ,
                        0x0328c101d01f4248b186db50dd8705135c0530ae4a52027fb02356b809e39ca8 ,
                        0xcf8490e976a2580f0b629e4f05182dcfc952dcff9145a952a8493e4456d52fb0 ,
                        0x887531bfbf50df2a568a8885fa95b7cbb11aacb7630e9729d6d4fac1e2ec8245 ,
                        0x371a96371a6da6c01d4f0606af9285423d05f5f9e2a236b054dccaf83a80f158 ,
                        0xd4396fb546f0a72ac8a2fc0a59651ad3577d4a4ec8adcbb2e555102d9a763830 ,
                        0xdf4970a563b008db424fbb317606c5c3c01bc995d2c7efd8929c0ad2f13fe141 ,
                        0xbd13c0f63091b78efe8cb24cc91ee678a9a0e0c7c3b83ae35d4b124204044922 ,
                        0xb1a2fe81896c26538cc96eaa5536c0bc2f84f23aa82e45f914df17f8a8b709ed ,
                        0xc945cd83abe38fb3c54a1661ee2bbd85cdcecebe190a7c9edcac4f65617b2821 ,
                        0x218541f1592816f4478ba0aaf0984954fe8df12fc252fd46055676a3cf3c7f4a ,
                        0x900d7267b4fbd08057aa43161a7d9c0509176a6d30c64c5a94b1090977aab2c3 ,
                        0x0d9a14b82bb86f935c8af3835024886e29a3e408fd7d445b5e550444bceb6f7c ,
                        0x5a2ca87b8e2c38af6f9047a2cf3f7913b0a0a08012ad3058dcdf49b08014fa60 ,
                        0xfe273166098d4d0407f6c73c61cd5720611a27235c39c9ea6f6cf439abca5223 ,
                        0x4480b13909797ce919fab32064312bd7949d7b08f83e2670ab3f61aac0f664c1 ,
                        0x444008138e497d690df6bf254096c080d5cab77844f3e11fea9f027c720dd504 ,
                        0x944f7ad797246b4b8a6fc8a2d6ee07fdf707f0209f7286bb0e64c0758a09718d ,
                        0x4939a628f18c5631f5813ae40eb7f925a051ee0e45fd8d420fad8ed1280dc6bc ,
                        0x119b993ac5c65a5aa38f7f5282528d91622777ef2ba98b8a8a6a8e18733b5cdb ,
                        0x318d8090dda7cbe4f24914a7551350b580662c88ea38341cee983c18595c7293 ,
                        0x88ff7cfeca0f19dc7d58141f92db8f8f479224ca5f5168e46feea8bb13891b77 ,
                        0xc18429b5f3573722c27a8b0c5de3fb3aeefcc73825d06a46e71ffc1402124948 ,
                        0x21b42f468909ceaf900ba77dbba8286b7300dc840acf2c2579b282780cdf3d0d ,
                        0x2007c10230b4477626b622a2d16aa12d130b94c8b32ac457b48a496004c82300 ,
                        0xe6f766a3f9b9989898db28b3849247a98aaf9989927b82124f60bc7304ec0f4d ,
                        0x98012058eee58813af98a3213f624d90f22a5cdd308a7b48b154c842f2584e24 ,
                        0x97a57c3b0d31af77bf90632794b29fe21e840f476bc5950b48d1da5cbcdb1913 ,
                        0x13ef41dbfd68eb71c1887d2565af5f1248d835a0205fb380a2a52ffaa68536c2 ,
                        0x23e3ad64fdf0d977c846f19d0c2bab6ca8ab4bd688b692d0ee53c86ce7e447a6 ,
                        0x8df97f2b4bf20c64fb2f937741d6fd3c3e1b2949e67cc208b8ece60434c973d0 ,
                        0xf61b189d8109b0f48411e98f88a720ab87fe4a30ee114ac67e9470b52b501282 ,
                        0x499cfe29040431ba62a08f0008cf7828ae81701f61527f423983499da642ab16 ,
                        0x66ff32b65be61321494150fc5c72c3946ca00dd545f649f5e50ba43010eb354d ,
                        0x36be65232a5b2943c4eb7ba8fb1eaec6900b77507f35c8208a118a1cff6fef2d ,
                        0x831f02213ede783face5e7e833d76a0d6f4adb1e2824d739924b954dfe23de9f ,
                        0x27eb84f23e597cef3ec9faa1ce5f69016221d6c8ac294ef664911fa2ce09d522 ,
                        0x2bc7d10f65f64eca64b5ec12c9dd3c7cbfd59fbcc8e84791dbd6086885c2a1ab ,
                        0xe5c0f12f5c46c88ab97c02c2accccf8f6a02828c82abfc8464f678824ac8781c ,
                        0xba791775f7613c6a33471b8fc7740843601c68df542be44116a1ff75e887ef26 ,
                        0xd0b612305aaf86510a1f97eb8ff082f142bd53302c7deb4d3e6a8095d49a00f6 ,
                        0x5ec5fe57a4d48ce222cd55523d220bcc6f776c710ce07b82289419fbda2e21eb ,
                        0x47812fad244f3dac38a50f591f5f6352fa4f2b1c967080bfcd6045511e89abe7 ,
                        0x950a315a7c981693362605e5646dc9bd5251c73375a2554eeed760c8aeb3f542 ,
                        0x0913ed06f8929b5b276040965536c50d0791c2a1a49768ab8bacb4d168e10f29 ,
                        0x12c5a456b46fe70f6fba84e84b3764004846c2c67b1eaa3e70b9d4008f2329c3 ,
                        0x24a334020b339c3c105964e8aa0d850d3c9645f247d8f91b1bc6035b3df206ef ,
                        0xef690f13197b30f768a46b83340896b79bafada96b22a3dea8612040db38b070 ,
                        0x648de179cc63c88a83d43d68f399626c2c18f256df6311e9cfb1691894f8e567 ,
                        0x294949f78168480e293993cfc06bc1ad9adfa31009ee966ee2a007acd73fa6fb ,
                        0xe5a7a28ff8bf8c00bf9c8aa4912e71d266b56592ba45a51f3a023a023a023a02 ,
                        0x3a023a023a02ff1502f32bee106cce5e82bdccc7f3955dbf1132f6a95760f28a ,
                        0x6f131caede78704f375ef22bd4fdbebc33b70919277c5fa5c93b7aaf602fed23 ,
                        0x6494f24b65fcb0163716b2dc5d846c67f59f43a7631caa6b2b6b2638ca1facaa ,
                        0x6b3fdc0a9fb514e87b47794f3c5af759fcb5c2ef791fb6227e7b133ff22b9af3 ,
                        0xcf6b1fb80c26d85c2182dd59f38655fbe1a70487f37981fad78edca3770a54e8 ,
                        0xa07659cefbf92bdd5840e7561604591e83ecdd856cf7334266f15d5caeacf2ea ,
                        0x7fdb9c7be04ec1563a88e3e92d5b869b6fe0ab7801475b59b5ecd457564977fc ,
                        0xd15d034d281cae51506c39c8b51acfcc5b8167a4545fad70b8ed785ad4540e1c ,
                        0x29c1e65e89bae751379503c93f77cfc67b731de5d3070ed73c287f17c85d80d7 ,
                        0xd99c38e9aed620b013636de0fd10f0f6b2507c365fc8c7d80ef736415398ddb5 ,
                        0x9e93d5ee7a196397e249561968d75f28463f0ef72e4e2e3a72ca7ae0fdd23a32 ,
                        0xd85d6918770dc6c913f241182291dd9584766fe3b38538df0902a9f745dadc31 ,
                        0x20ce147ebec8d902fdad136c7b5be26145b7f33ef2b150ed6e13da9dc6773b50 ,
                        0xf77121b3f4459ccf54dbe3bdddb91def1703b3b7f87c48cecc8abbd06e37c8ae ,
                        0x2e7092dfe1cee1e7d9ee2821cb5580392e811c0df401e936f74428d6e201311d ,
                        0x96815f875589e22c03609b85ecddea462fad5c228e766497de0d859470926996 ,
                        0xd29b050e670e27c98ae25b506f8b603f12a6f68b3ef2cadaa2dd215899a642b6 ,
                        0x2b0c0acbe54db39c0b50f7359578185b25e838be386c65c1bc3d11d9e176a29c ,
                        0x00419ee6751d65756f6875b8d7825829c2a213aa05220b9ae5de8ef6ea7f8290 ,
                        0x7526c2933577b88d20d134cf3c5be2fd166e7d89800e10962c1eefc309e2a29d ,
                        0x4aba70c83557951173b297f19b6c39d11dee35fc3d7907bbbb12a54c4807c9b8 ,
                        0xc503e15479e60083857c7137d8c3ee8c0010ea7f64d80ef70050aa221dce9938 ,
                        0x2732c05a940ee59f2dae7c9003adb90b22aedd958d026be1acfb3c3fbbcb8195 ,
                        0xadde9c40ca250b426e8a149273e45dbcbeea51c4401035cba3fcc720cf46ae68 ,
                        0xb2ce7464b947638c23789f056bd78713c6ee5e25e4b887c282148210069f16d0 ,
                        0x0a2b4d96c80e32d98fb512b24a27c23a25d5d0b51d84a485441630d315cdbf23 ,
                        0x574cf326f2917bb73b37e05525a0ddb908f5d5dfc3d85c2340fc54c18a456483 ,
                        0x1c144268072d22b2b6e4511ce4399cc3b9c5b53b63ab161b0f114ae3f11e1612 ,
                        0x614d833c1cee49000856e214dc8e6b39caeb2a495cc7397064314849a4f49c23 ,
                        0x0f7165d24d926a9d773cae7b3300de576559aa945096c7632dd57def131c2588 ,
                        0x214104bb6b8f4a6c0fc172ca118fc125562b0fdf39ff586555b9458432b5c30a ,
                        0x529085a2f8956243bbeb12cae61afa2332d85d912a695c6fa1fe501e0bda5d45 ,
                        0x553158b6732477816451ed2e895b413a54d2ed1516b81e11720e3f0439f75485 ,
                        0x1c0ee752ee465532527bf52e149b0b58b9d47045c56627e6d08fe3ca31c39179 ,
                        0x78203efb140b3a9fbf27ec692c477972952b6f702474b886e03f750f815cebb8 ,
                        0x4be0013e2c8ecda5ba650eb4fb0deee6165052e1719536e7304e5ced70b87301 ,
                        0x6ccdfbd7c802e41cdd817a9b6029d4f82abbb80dfa70f064230b96858278724b ,
                        0x36d78ceabe602d1cb00cda41b156f69162f4b51c4a7f49c0ff0243a65c2e0f57 ,
                        0xacfb39bc77d4d01dd52172908b25abae252d9920a5bd6c2b6458c52d6d8e5bbd ,
                        0x5e9c553656c8ad28e271294f384ac7a30ed5db0ad9275463519a569564d8ca06 ,
                        0x411ef51f3b699111d1b3d02f2d06b2f65c36b86e92554bd4b2c86a620c15578a ,
                        0x2941542cc6acf25fe5f1743530fb55de507ce7287e9067a1da41b10f29413b08 ,
                        0x3c8af1e87679724b74cc43a04e64ad6ed38c6785de077781ce87795ded20b7a3 ,
                        0x253a14ec930520eb4ad974559d334d796ca81d246386fb01b8c6f63c2ea383e4 ,
                        0xf0965193ab368894b9d6962b97e25024055a264f6d565c429c8a4541e3682106 ,
                        0x59c06c14ef83fad2da71b9bc7600a89dc3f5448d988e6424d9bcb35c6f5933b1 ,
                        0x00bcb3e45f8504faa03a023a023a023a023a023a023a023a023a023a023a023a ,
                        0x023a023a023a023a023a023a023a023a023a02d71181ff070c60ff4e76ca6c5b ,
                        0x0000000049454e44ae426082
                    End

                    LayoutCachedLeft =15
                    LayoutCachedTop =75
                    LayoutCachedWidth =2304
                    LayoutCachedHeight =1269
                    TabIndex =7
                End
            End
        End
        Begin Section
            CanGrow = NotDefault
            Height =9375
            Name ="Detail"
            Begin
                Begin TextBox
                    SpecialEffect =0
                    OldBorderStyle =1
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =1110
                    Width =420
                    TabIndex =1
                    BorderColor =13017476
                    Name ="groupNo"
                    ControlSource ="groupNo"

                    LayoutCachedLeft =1110
                    LayoutCachedWidth =1530
                    LayoutCachedHeight =240
                End
                Begin TextBox
                    DecimalPlaces =0
                    SpecialEffect =0
                    OldBorderStyle =1
                    OverlapFlags =87
                    TextAlign =3
                    IMESentenceMode =3
                    Left =1530
                    Width =900
                    TabIndex =2
                    BorderColor =13017476
                    Name ="groupUnits"
                    ControlSource ="groupUnits"

                    LayoutCachedLeft =1530
                    LayoutCachedWidth =2430
                    LayoutCachedHeight =240
                End
                Begin TextBox
                    DecimalPlaces =0
                    SpecialEffect =0
                    OldBorderStyle =1
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =2460
                    Width =600
                    TabIndex =3
                    BorderColor =13017476
                    Name ="groupYear"
                    ControlSource ="groupYear"

                    LayoutCachedLeft =2460
                    LayoutCachedWidth =3060
                    LayoutCachedHeight =240
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =0
                    OverlapFlags =93
                    TextAlign =3
                    IMESentenceMode =3
                    Left =13650
                    Top =480
                    Width =480
                    TabIndex =10
                    ForeColor =255
                    Name ="groupID"
                    ControlSource ="groupID"

                    LayoutCachedLeft =13650
                    LayoutCachedTop =480
                    LayoutCachedWidth =14130
                    LayoutCachedHeight =720
                End
                Begin Tab
                    OverlapFlags =247
                    BackStyle =0
                    Top =690
                    Width =14415
                    Height =8325
                    TabIndex =11
                    Name ="GroupTabCtl"

                    LayoutCachedTop =690
                    LayoutCachedWidth =14415
                    LayoutCachedHeight =9015
                    Begin
                        Begin Page
                            OverlapFlags =119
                            Left =15
                            Top =1095
                            Width =14265
                            Height =7785
                            Name ="Spec Review"
                            EventProcPrefix ="Spec_Review"
                            LayoutCachedLeft =15
                            LayoutCachedTop =1095
                            LayoutCachedWidth =14280
                            LayoutCachedHeight =8880
                            Begin
                                Begin Subform
                                    OverlapFlags =247
                                    SpecialEffect =0
                                    Left =15
                                    Top =1215
                                    Width =14265
                                    Height =7664
                                    BorderColor =16777215
                                    Name ="Spec_Review_sub"
                                    SourceObject ="Form.Spec_Review_sub"
                                    LinkChildFields ="groupRef"
                                    LinkMasterFields ="groupID"

                                    LayoutCachedLeft =15
                                    LayoutCachedTop =1215
                                    LayoutCachedWidth =14280
                                    LayoutCachedHeight =8879
                                End
                            End
                        End
                        Begin Page
                            OverlapFlags =247
                            Left =105
                            Top =1095
                            Width =14175
                            Height =7785
                            Name ="Credit Request"
                            EventProcPrefix ="Credit_Request"
                            LayoutCachedLeft =105
                            LayoutCachedTop =1095
                            LayoutCachedWidth =14280
                            LayoutCachedHeight =8880
                            Begin
                                Begin Subform
                                    OverlapFlags =247
                                    OldBorderStyle =0
                                    SpecialEffect =0
                                    Left =105
                                    Top =1650
                                    Width =14040
                                    Height =4919
                                    Name ="Credit_Request_sub"
                                    SourceObject ="Form.Credit_Request_sub"
                                    LinkChildFields ="groupRef"
                                    LinkMasterFields ="groupID"

                                    LayoutCachedLeft =105
                                    LayoutCachedTop =1650
                                    LayoutCachedWidth =14145
                                    LayoutCachedHeight =6569
                                End
                            End
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    Visible = NotDefault
                    OverlapFlags =93
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =14100
                    Top =90
                    Width =15
                    FontSize =8
                    TabIndex =14
                    BorderColor =13017476
                    Name ="groupTypeRef"
                    ControlSource ="groupTypeRef"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT lookup_groupTypes.groupTypeID, lookup_groupTypes.groupType FROM lookup_gr"
                        "oupTypes ORDER BY lookup_groupTypes.groupType;"
                    ColumnWidths ="0;1440"
                    OnDblClick ="[Event Procedure]"
                    FontName ="Tahoma"
                    ListItemsEditForm ="GroupType_Add"

                    LayoutCachedLeft =14100
                    LayoutCachedTop =90
                    LayoutCachedWidth =14115
                    LayoutCachedHeight =330
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    Visible = NotDefault
                    OverlapFlags =223
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =1440
                    Left =13920
                    Top =225
                    Width =255
                    FontSize =8
                    TabIndex =12
                    BorderColor =13017476
                    Name ="groupMakeRef"
                    ControlSource ="groupMakeRef"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT lookup_GroupMakes.groupMakeID, lookup_GroupMakes.groupMake FROM lookup_Gr"
                        "oupMakes ORDER BY lookup_GroupMakes.groupMake;"
                    ColumnWidths ="0;1440"
                    OnExit ="Requery.ReqGrpmodel"
                    OnDblClick ="[Event Procedure]"
                    FontName ="Tahoma"
                    OnChange ="Requery.ReqGrpmodel"
                    ListItemsEditForm ="Make_add"

                    LayoutCachedLeft =13920
                    LayoutCachedTop =225
                    LayoutCachedWidth =14175
                    LayoutCachedHeight =465
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    Visible = NotDefault
                    OverlapFlags =87
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =13845
                    Top =120
                    Width =75
                    FontSize =8
                    TabIndex =13
                    BorderColor =13017476
                    Name ="groupModelRef"
                    ControlSource ="groupModelRef"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT lookup_groupModels.groupModelID, lookup_groupModels.groupModel FROM looku"
                        "p_groupModels ORDER BY lookup_groupModels.groupModel;"
                    ColumnWidths ="0;1440"
                    OnDblClick ="[Event Procedure]"
                    FontName ="Tahoma"
                    ListItemsEditForm ="Model_Add"

                    LayoutCachedLeft =13845
                    LayoutCachedTop =120
                    LayoutCachedWidth =13920
                    LayoutCachedHeight =360
                End
                Begin Label
                    FontItalic = NotDefault
                    OverlapFlags =85
                    Left =1110
                    Top =300
                    Width =420
                    Height =240
                    ForeColor =9211020
                    Name ="Label23"
                    Caption ="Grp"
                    LayoutCachedLeft =1110
                    LayoutCachedTop =300
                    LayoutCachedWidth =1530
                    LayoutCachedHeight =540
                End
                Begin Label
                    FontItalic = NotDefault
                    OverlapFlags =93
                    Left =1575
                    Top =300
                    Width =900
                    Height =240
                    ForeColor =9211020
                    Name ="Label24"
                    Caption ="# of Units"
                    LayoutCachedLeft =1575
                    LayoutCachedTop =300
                    LayoutCachedWidth =2475
                    LayoutCachedHeight =540
                End
                Begin Label
                    FontItalic = NotDefault
                    OverlapFlags =215
                    TextAlign =3
                    Left =2460
                    Top =300
                    Width =600
                    Height =240
                    ForeColor =9211020
                    Name ="Label25"
                    Caption ="Year"
                    LayoutCachedLeft =2460
                    LayoutCachedTop =300
                    LayoutCachedWidth =3060
                    LayoutCachedHeight =540
                End
                Begin Label
                    FontItalic = NotDefault
                    OverlapFlags =85
                    Left =3090
                    Top =300
                    Width =1815
                    Height =240
                    ForeColor =9211020
                    Name ="Label26"
                    Caption ="Make"
                    LayoutCachedLeft =3090
                    LayoutCachedTop =300
                    LayoutCachedWidth =4905
                    LayoutCachedHeight =540
                End
                Begin Label
                    FontItalic = NotDefault
                    OverlapFlags =85
                    Left =4920
                    Top =300
                    Width =2355
                    Height =240
                    ForeColor =9211020
                    Name ="Label29"
                    Caption ="Model"
                    LayoutCachedLeft =4920
                    LayoutCachedTop =300
                    LayoutCachedWidth =7275
                    LayoutCachedHeight =540
                End
                Begin Label
                    FontItalic = NotDefault
                    OverlapFlags =85
                    Left =7305
                    Top =300
                    Width =1935
                    Height =240
                    ForeColor =9211020
                    Name ="Label27"
                    Caption ="Type"
                    LayoutCachedLeft =7305
                    LayoutCachedTop =300
                    LayoutCachedWidth =9240
                    LayoutCachedHeight =540
                End
                Begin Label
                    FontItalic = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    Left =11985
                    Top =300
                    Width =1500
                    Height =240
                    ForeColor =9211020
                    Name ="Label8"
                    Caption ="Preliminary Total "
                    LayoutCachedLeft =11985
                    LayoutCachedTop =300
                    LayoutCachedWidth =13485
                    LayoutCachedHeight =540
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    SpecialEffect =0
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =15
                    Width =555
                    Height =255
                    BackColor =9211020
                    BorderColor =9605778
                    ForeColor =16777215
                    Name ="FAID"
                    ControlSource ="FAID"

                    LayoutCachedLeft =15
                    LayoutCachedWidth =570
                    LayoutCachedHeight =255
                End
                Begin Label
                    FontItalic = NotDefault
                    OverlapFlags =85
                    Left =60
                    Top =300
                    Width =495
                    Height =240
                    ForeColor =9211020
                    Name ="Label77"
                    Caption ="FAID:"
                    LayoutCachedLeft =60
                    LayoutCachedTop =300
                    LayoutCachedWidth =555
                    LayoutCachedHeight =540
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =9300
                    Width =2295
                    FontSize =8
                    TabIndex =7
                    BorderColor =13017476
                    Name ="Combo105"
                    ControlSource ="GRPLOCSN"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Locations.Locationsn FROM Locations WHERE (((Locations.ClientREF)=[forms]"
                        "![GROUPS_VIEW]![CLIENTREF]) AND ((Locations.LocObsolete)=No));"
                    ColumnWidths ="2160"
                    StatusBarText ="lOCATION SHORTNM FOR REPOPRT (PIPELINE)"
                    FontName ="Tahoma"

                    LayoutCachedLeft =9300
                    LayoutCachedWidth =11595
                    LayoutCachedHeight =240
                    Begin
                        Begin Label
                            FontItalic = NotDefault
                            OverlapFlags =85
                            Left =9255
                            Top =300
                            Width =2310
                            Height =240
                            ForeColor =9211020
                            Name ="Label106"
                            Caption ="Location:"
                            LayoutCachedLeft =9255
                            LayoutCachedTop =300
                            LayoutCachedWidth =11565
                            LayoutCachedHeight =540
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =615
                    TabIndex =9
                    Name ="AltEquip"
                    ControlSource ="AltEquip"

                    LayoutCachedLeft =615
                    LayoutCachedWidth =875
                    LayoutCachedHeight =240
                    Begin
                        Begin Label
                            FontItalic = NotDefault
                            OverlapFlags =85
                            Left =570
                            Top =300
                            Width =510
                            Height =240
                            FontWeight =700
                            ForeColor =2366701
                            Name ="Label110"
                            Caption ="AltEq"
                            LayoutCachedLeft =570
                            LayoutCachedTop =300
                            LayoutCachedWidth =1080
                            LayoutCachedHeight =540
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    SpecialEffect =2
                    OldBorderStyle =0
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =3105
                    Width =1785
                    FontSize =8
                    TabIndex =4
                    BorderColor =0
                    Name ="groupMake"
                    ControlSource ="groupMake"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT DISTINCT lookup_GroupMakes.groupMake FROM lookup_GroupMakes ORDER BY look"
                        "up_GroupMakes.groupMake;"
                    OnExit ="Requery.ReqGrpmodel"
                    FontName ="Tahoma"
                    OnChange ="Requery.ReqGrpmodel"
                    AllowValueListEdits =0

                    LayoutCachedLeft =3105
                    LayoutCachedWidth =4890
                    LayoutCachedHeight =240
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    SpecialEffect =2
                    OldBorderStyle =0
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =4920
                    Width =2400
                    FontSize =8
                    TabIndex =5
                    BorderColor =0
                    Name ="groupModel"
                    ControlSource ="groupModel"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT DISTINCT lookup_groupModels.groupModel FROM lookup_groupModels WHERE (((l"
                        "ookup_groupModels.groupmake)=[FORMS]![GROUPS_vIEW]![GROUPMAKE])) ORDER BY lookup"
                        "_groupModels.groupModel;"
                    FontName ="Tahoma"
                    AllowValueListEdits =0

                    LayoutCachedLeft =4920
                    LayoutCachedWidth =7320
                    LayoutCachedHeight =240
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    SpecialEffect =2
                    OldBorderStyle =0
                    OverlapFlags =87
                    IMESentenceMode =3
                    Left =7335
                    Width =1935
                    FontSize =8
                    TabIndex =6
                    BorderColor =0
                    Name ="groupType"
                    ControlSource ="groupType"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT DISTINCT lookup_groupTypes.groupType FROM lookup_groupTypes ORDER BY look"
                        "up_groupTypes.groupType;"
                    FontName ="Tahoma"
                    AllowValueListEdits =0

                    LayoutCachedLeft =7335
                    LayoutCachedWidth =9270
                    LayoutCachedHeight =240
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =11655
                    Top =30
                    TabIndex =8
                    Name ="DealPLB"
                    ControlSource ="DealPLB"

                    LayoutCachedLeft =11655
                    LayoutCachedTop =30
                    LayoutCachedWidth =11915
                    LayoutCachedHeight =270
                    Begin
                        Begin Label
                            FontItalic = NotDefault
                            OverlapFlags =85
                            TextAlign =1
                            Left =11595
                            Top =300
                            Width =360
                            Height =240
                            ForeColor =9211020
                            Name ="Label114"
                            Caption ="PLB"
                            LayoutCachedLeft =11595
                            LayoutCachedTop =300
                            LayoutCachedWidth =11955
                            LayoutCachedHeight =540
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =255
                    Left =405
                    Top =1215
                    Width =2040
                    Height =330
                    FontSize =9
                    FontWeight =700
                    TabIndex =15
                    ForeColor =12349952
                    Name ="AddCreditRequestBtn"
                    Caption ="Add Credit Request"
                    OnClick ="[Event Procedure]"
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =405
                    LayoutCachedTop =1215
                    LayoutCachedWidth =2445
                    LayoutCachedHeight =1545
                    Overlaps =1
                End
                Begin TextBox
                    DecimalPlaces =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =11955
                    Height =255
                    TabIndex =16
                    Name ="groupTotal"
                    ControlSource ="groupTotal"
                    Format ="$#,##0.00;($#,##0.00)"
                    StatusBarText ="moved to quotes table"

                    LayoutCachedLeft =11955
                    LayoutCachedWidth =13395
                    LayoutCachedHeight =255
                End
            End
        End
        Begin FormFooter
            Height =0
            BackColor =-2147483633
            Name ="FormFooter"
        End
    End
End
CodeBehindForm
' See "Groups_view.cls"
