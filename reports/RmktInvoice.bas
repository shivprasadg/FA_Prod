Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    AutoResize = NotDefault
    PopUp = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    TabularCharSet =163
    DateGrouping =1
    GrpKeepTogether =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =14385
    DatasheetFontHeight =11
    ItemSuffix =177
    Left =5850
    Top =3165
    ShortcutMenuBar ="PrintMe"
    RecSrcDt = Begin
        0x652208691490e540
    End
    RecordSource ="SELECT * FROM vw_BuyersOrder WHERE (((vw_BuyersOrder.InvoiceRequestNum)=[Forms]!"
        "[ReMarketInvoiceManager]![InvoiceRequestID]));"
    DatasheetFontName ="Calibri"
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
        Begin EmptyCell
            Height =240
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin BreakLevel
            GroupHeader = NotDefault
            GroupFooter = NotDefault
            ControlSource ="VIN"
        End
        Begin FormHeader
            KeepTogether = NotDefault
            Height =2797
            Name ="ReportHeader"
            Begin
                Begin TextBox
                    OldBorderStyle =1
                    TextFontCharSet =186
                    TextAlign =2
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Left =12120
                    Top =345
                    Width =1590
                    Height =315
                    BorderColor =4210752
                    Name ="Text30"
                    ControlSource ="InvoiceNumber"
                    GroupTable =15
                    RightPadding =38
                    BottomPadding =38

                    LayoutCachedLeft =12120
                    LayoutCachedTop =345
                    LayoutCachedWidth =13710
                    LayoutCachedHeight =660
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =5
                    GroupTable =15
                    Begin
                        Begin Label
                            OldBorderStyle =1
                            TextFontCharSet =238
                            TextFontFamily =34
                            Left =10425
                            Top =345
                            Width =1620
                            Height =315
                            FontWeight =700
                            BorderColor =4210752
                            Name ="Label31"
                            Caption =" Invoice #:"
                            GroupTable =15
                            RightPadding =38
                            BottomPadding =38
                            LayoutCachedLeft =10425
                            LayoutCachedTop =345
                            LayoutCachedWidth =12045
                            LayoutCachedHeight =660
                            LayoutGroup =5
                            GroupTable =15
                        End
                    End
                End
                Begin TextBox
                    OldBorderStyle =1
                    TextFontCharSet =186
                    TextAlign =2
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Left =12120
                    Top =735
                    Width =1590
                    Height =315
                    TabIndex =1
                    BorderColor =4210752
                    Name ="Text32"
                    ControlSource ="InvoiceDate"
                    GroupTable =15
                    RightPadding =38
                    BottomPadding =38

                    LayoutCachedLeft =12120
                    LayoutCachedTop =735
                    LayoutCachedWidth =13710
                    LayoutCachedHeight =1050
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =5
                    GroupTable =15
                    Begin
                        Begin Label
                            OldBorderStyle =1
                            TextFontCharSet =238
                            TextFontFamily =34
                            Left =10425
                            Top =735
                            Width =1620
                            Height =315
                            FontWeight =700
                            BorderColor =4210752
                            Name ="Label33"
                            Caption =" Invoice Date:"
                            GroupTable =15
                            RightPadding =38
                            BottomPadding =38
                            LayoutCachedLeft =10425
                            LayoutCachedTop =735
                            LayoutCachedWidth =12045
                            LayoutCachedHeight =1050
                            RowStart =1
                            RowEnd =1
                            LayoutGroup =5
                            GroupTable =15
                        End
                    End
                End
                Begin Image
                    Left =90
                    Width =3528
                    Name ="Auto_Logo0"
                    Picture ="SmallLogo.gif"
                    ImageData = Begin
                        0x474946383961470136007000002c000000004701360087ffffff52525a7bc542 ,
                        0x525a5affffefefffe6adadb5eff7d652524abde6bd73c53aa5ad94b5ce9cbdc5 ,
                        0xc5d6ded68c63196363197b636310198c6310946bce946bdece6b9cce6bde528c ,
                        0xc58c846b7b73b542def7bdb5ef63e6ef63e6de3142ef1942ad19b5de3110ef19 ,
                        0x10ad1910dead109cad10efde10adde6be619e6ce6342ce19428c1910ce19108c ,
                        0x1910cede108cdeeff7f784b54aa5c58c84c54a8c3a5a8c63a5633a5a8c105a63 ,
                        0x105a8c945adef7ce8ce65a636384effff742ef5a84634242ad5a949c9ce619ef ,
                        0xe619ad10ef5a10ad5ab519efb519ad6b9ca5e67bad6b945a42ce5a428c5ae619 ,
                        0xcee6198c10ce5a108c5ab519ceb5198ce67b8c73b521635a638c8c8cded6b58c ,
                        0x5aef42efde3a6bde42dead8cef943a5aad429cad8c19ef3a19ad42adde8c3194 ,
                        0x3a29de3a6b193a2919635aef105aad6319ef1019ad633194106bde1029de106b ,
                        0x191029196bef946bdeef6b9cefb5deef8cdece8c9cce6b941942cede3a4ade42 ,
                        0x8cde3a08de3a4a193a0819104ade1008de104a19100819ade694525a633a4a52 ,
                        0x738484b5a56bdee6e6efe6efe6a56be64aefe6426be64aade6a529e64229b542 ,
                        0x6bb5a529b54229b5736be6736bb54aefe6106bb54aade67329e61029b5106bb5 ,
                        0x7329b51029b57befe67befb5a54ae6a54ae64acee6424ae64a8ce6a508e64208 ,
                        0xb5424ab5a508b54208b5734ae6734ab54acee6104ab54a8ce67308e61008b510 ,
                        0x4ab57308b51008b57bcee67bcead84adad848c3a6b5ae6adad3a295a6363a510 ,
                        0x6b5a10295ae6ad8c3a085a104a5a10085a8cce19b5adce8cad19c5e69c8cdeef ,
                        0x8cef198c9cef8c8c19636b6be6adefb5adefe6ef10b5ef1010ef8c10ad8ce6d6 ,
                        0x94e6adcee6ce10b5ce1010ce8c108c8cb5ce5aadce7b8cc5638c3a29633a298c ,
                        0x10296310298c6bce42ef8c3a6b8c42ad8c8c29ce3a298c8c10a5636bce106b8c ,
                        0x6329ce8c3a08633a088c10086310088c4ace42ce8c3a4a8c428c8c8c08ce3a08 ,
                        0x8c8c1084634ace104a8c6308ce6ba584e6f79c8cb56b94ad84efd6d6d6ceef94 ,
                        0x8c9c6bb5526bc519526b63ffdeef73c5526bc54252526b08ff0037509b41b0a0 ,
                        0xc18308132a5c3883da0602000810d8402f8680193b3800da50a000808f20438a ,
                        0x1c49f26383400102a04cc9b2654b63001ab80c00b3a4c8203373ba0c1404004e ,
                        0x9d400318009021a8519a36932a5dcab4a9d3a750436ee817c3a202053304f423 ,
                        0x7831ab80af60c38a1d4b166c8c6ad4200228108caa000d7035cca0174c87c788 ,
                        0x00f89154cb54e64c632d07b4ac62859f01c1825366e8c13443e29c55060788dc ,
                        0x80684ac19159267e0c3840e57d2b2707cd9cb20ae1a8a853ab5ead1a10bd191a ,
                        0xc0c68e2117ae00da6f61e78e5d7bb7efdab379c3a667776db08a17bbc6a0572d ,
                        0x860c8e7cf7365d507a26e2d203100c18046001e9c8dc9716ff008c39b169c92c ,
                        0xab3430646c00e9cd9331974650e59e21d3e725d3179d9ee502d6000628206bd4 ,
                        0xf4268002167da5c057b1ddb61b6c10d6162184b925f7558233c4a0430f3d1420 ,
                        0x836e0d8648103d1b1467537422f19541209955619e682e9656c57f835c16c000 ,
                        0x03fcb794212845861902faf15785210ed888a38f995d076400111470524a4046 ,
                        0xd6e28d5092c652650366a9e596235573db826f31982056619655d60c58851866 ,
                        0x835fcd90c05a0450534d410d2ad020411a24c02187292615dd20ed19e9a27bd9 ,
                        0xb95805903d65106366432d55a479dab9489fa18446660c9154c617007d95565a ,
                        0x85313019609a7b4802b9a4949e4e862597acb6ca9a576022ff28179807c680a0 ,
                        0x9db66a90ebae761ea82b56b726386b6cd51cc06101cdc9a6e0b21a544317634e ,
                        0xa9a5962145d9188015410461c502d96acb6d100e0010e8a9abe24552034926f6 ,
                        0x2db7d8b21b04b7001820daa9ec2eb0edb6da6e3b5403190cd2afbf5614756476 ,
                        0xd7f6dbaf12195851aeab0c37bcd40eba2a6891ad3100a32c82b51e9815c60b52 ,
                        0x7c6bb057a1696b861cdca5c33433586c91066052712056cb99d8547485b41723 ,
                        0x7d0bf7c91eaa2985bb5403ee5d3640068630256f7c385a51f4520418d2912104 ,
                        0x747458601914e2d1d24d3bacf5d625cd7955c661d27a6badb12e3bd6d80ba2fd ,
                        0x155a0430a64345221b58d585670182e2891fb5fd9103fb54ff89a3cf49a14b70 ,
                        0x7cf630250b8e546ec7140141f8589a15ab1940df7e03402e120c5c67aef90e73 ,
                        0xbf5c660c54581516ad669a497a550c78d4c3dbcecee97a35b0c73e03ecf41cb0 ,
                        0xf4477729c518baf37e0ab84d40274e532121ddfd911582ed8780e54af163c58d ,
                        0x9bf5d4675356249912f39a67af392031887cdbac60911e16e81e837e9bcb668d ,
                        0xedf16dd504c3580f8504233f20f303d20c20d4c82f0320f437a3c3ddb913c9ed ,
                        0xe245a5254560802591d7a2867617ccd9c431889b0cf68a97b7b574063b8d1a89 ,
                        0xf12ee798ea08e52908d49e0851338f8ae0e656b0b118c8287620b071ac7c69b3 ,
                        0x4a82de4297dc1da000516b1a016e18911c4aa42312218903ff9362807d200647 ,
                        0x83082100a0f511ef68874a5618a24d0a3088f200298397fb0863ee23a918350a ,
                        0x5a4c5c4a0f2280b8fd60512957830a010c60809c6d4d895a2b0035e8219797b1 ,
                        0x104c147319c8ca54365b7dc97c64bb4d0c3892971f4244226a799f164112c490 ,
                        0xe8052252bcdc4f0e9512e9299224cf9bcf00ce68ae90b047339371c021193992 ,
                        0x22b5443bbf0349188508804f3a2e00a9b4090cf0118141b8d1260e008c8e4618 ,
                        0xc0381e2759647a0b1e63e84219fef140a18bd51dbb472244c2c99938644c0e71 ,
                        0xa8c5a8e9c5813d885a5e6cd2037c382e329c84d6edaa779d0f3aca3451ea192b ,
                        0x4302344205601f19f8dd2a7b299242f8683f5500dc06f522ff13c19c118e26b9 ,
                        0x912c46d8000300d4551e4ad6cb6a93203f76eca15f0a931ffd18d188cea65888 ,
                        0xd44102ae90800d6c34011bd5413012100c8f7eb4a3202d51217c18462952eb46 ,
                        0x94dce43a3f52880e5e6761c6eba7f082608020c8825b3ff569ee78f7129e66ab ,
                        0xa73de569498628b8971cb49a1d24cc23fb421f0688d094b17458014e560d1536 ,
                        0x684c2a7be1d842d742b17a8c63b7a1060e3b528f8af4c384198ac15b975395e6 ,
                        0xb855ae7505845da8a9c49a0d4e30c6f017a8fa658c0c64004bb9bc113e6f2912 ,
                        0x51658a3f99124c208c41bc8f1c2d689b22cd4a5092812146d224858acf01fbb2 ,
                        0x292b68a7b2ab24c9d118dbb02219432fd9630b35aad23ddeff5ce86b7bdc23f9 ,
                        0x208a20f4453406a9c3e1db66e09636a5c92c61815056e24a0f1924409b8e8c88 ,
                        0x038084b8f7d0443303f059038c914ef7647524d411cde06204d3c9446088d5a3 ,
                        0xee636194924084272938890c750721ada40c823092cba753bc638cc269af7a13 ,
                        0xe49a447420a739d5d662e8b3d3d83460b1b2aaf0402a6c7083d0075c88f42018 ,
                        0xcd61d39dd2443add682c36718d01206e184688e8745054b2124df6b134a226c6 ,
                        0x18bf73200283009f9b41f646f81862543da8e2c8047824f8b05600a417dd8fb4 ,
                        0x38250d204076fe69136f1ae3b31ff9ec5395f259e4fd18ca23496d5410390f19 ,
                        0x98704e198a6b73421cd7d9752fcc631e336dc37c96604484ff2d5bb168f8d637 ,
                        0xbe5e8ded22d590c1ff28e897e4656a3fd5c9004864a2de2a0c02b64ac9c03e62 ,
                        0x8a1d4a196a973b2be7c02835194e8e2493529a51962f5d850c1400068eb1023d ,
                        0x5b6900c3dad210f75d4c03b0a5440358018b86587506022b8bcaf4c010aed66e ,
                        0x100cdba8a239005d846980b011d800590ca296f830e8124d828fc2da12003030 ,
                        0x04631cd8006d19436121ec48010e90001978d9cb7394133dc20dee398efbdce1 ,
                        0xa6c734d04d0d708f5b944b3c4e868829313c3a08641e9e8d739ec6c8a97d9354 ,
                        0x4792d27bbd539ecaa99224d2bcaf625d54d88647c0180fcf00163bc3a90134fc ,
                        0xe2b38e006b97a8a8c70e05cb0135277586046d763ecf8887ff3af6b508e89956 ,
                        0xb2133006b80bbf52d237ed5cea2334a94ce3a0b7f298406953a2513601ee8193 ,
                        0xf63c064b86789ed13d48d98fc4973f44c6dd9b7fa86dbbd8e50007d800d66fa8 ,
                        0x83ab6b7bdb60efc801bade756d1f20880500443fc097315b890d90b94dd3889e ,
                        0x7b13c5c2c80a831844c0f41e8157836492518afab2475253f5de28e64e834101 ,
                        0x3aa47890c02097e4c5d9e29d76ebc5573629ec790c65300f80207bdab2ef4c65 ,
                        0x62af5550d3de087245d2b4489e375a18fc64106d6c001923c0180750c69b11e8 ,
                        0x2977b31b9380a544ef56b08276b91bec551f0a4b5680f9af7f52d8f5bc740041 ,
                        0x1076e310704b6a2e31778844e4e221824368ae054e61cc66fff6a3c6bd31b1b0 ,
                        0x85cb1cdd87297a16e2740424a68d91608250007e34ad0050ebd03837759d9ea4 ,
                        0x968943d737e5054b4c710fe921182ca61ab9f422049814b6b70f44667b013050 ,
                        0x209174a6812510e158ff01031137444307734e2741b05500b3161eb637344356 ,
                        0x598e4564a2a25f2fd76980232a37476806c01711702d7741632df711ce130082 ,
                        0x568149516278d1361eb1784bc4181ea137aac327aa833b05400f542136b7010c ,
                        0xb1b12075525162b120c08020edc3173000184f7428962640e46123e5f259bbb3 ,
                        0x2454026346034a9f723b60446501f54d24e7788b743c34713b35731a25d70082 ,
                        0x800035c84e48868797022d30605ab4171391b14bad544926ff415d56202da9b7 ,
                        0x4bc8630c0710120e501489087f9f0211def4795af43c96d34f8d626145f11f53 ,
                        0x1565baf3665d377663b7555ce78a3ac075ad68753904003ad01c1393310fd550 ,
                        0x668531e1a38568b1785b640cf2852332357824710052921e0d007283a62993d1 ,
                        0x744b31727eb61825a16522a15388110197772e55f00f8c0800c977734b140186 ,
                        0x8668267128f7107218780f20482dcb334012d828f28200e6388744568e46067a ,
                        0xfbf03b04705f35e118f8b07af0c418f8b00fa3b537e1b571d3833bdd460d7232 ,
                        0x911389161649911729030cc07d3a805731f45031843eb3314c77e65bc5721785 ,
                        0xc022ca934f7c012ddc475382404986924a98b34af7c81f3fff96454106744963 ,
                        0x88aad8140a941e56a08d1f116475081287a31e21378edd3119970800f780337f ,
                        0x671a1e31350be317c8074edb3888ad54143fb83793217886404696e3183de140 ,
                        0x0ec022439158414024fc628c01c09478d348d9f43670b31554919767b6975b41 ,
                        0x1055110cd49400c8c131c0386768f531c62416d5f0268c61808e8323d2a6148f ,
                        0x92294f668721813998b300418338053912b0354edf316472689920115e942378 ,
                        0xfd08002778586d147bc8634e0a144bde94806b71833f984b80783cf0446c87e2 ,
                        0x33ac974a8e8574c6f85e8ab883205133a2498e16777945573432612934772ded ,
                        0xf814d0b54404a0030c4047047115720716b0b115c1321780ff1911c1d03d60d2 ,
                        0x2054f847caf45b56184878f41069582508d0241ab48d4b82195fe95997564029 ,
                        0xa1230ee440bd542d87a21d58a29f25b754219829aa577203944929c69387a223 ,
                        0x8d938f46961211101241b00f35d16c379758139474ea4414fba0348d3519fee5 ,
                        0x0082c188ac476c291188db457a33670c8d426810872d0e50341b34123549125b ,
                        0xf51aa2c320c2642759e15019d20f1b002709401be7875675222c64e38be52300 ,
                        0x8b494d41497335a18c23713499450b87256cc2760fe1e252d5221ac6800f0e00 ,
                        0xa6606a08b7b33435853889610067fa6bf61149eae849e484341448120b9830c1 ,
                        0xd7a7c1670cbb49144373996b2418433968915108f60082aaffa94b2271822447 ,
                        0x823d1712d5438df548a9a062084324a9c6f9a29a92016a212fc6b03021e4853f ,
                        0x798722c116f53043bd123ebb012c55412245c800ddc32ce9a39ee7135177f616 ,
                        0x4a7a1b478a875532a9f5f9113ba6862a41497d57a000e05737230834f1299f12 ,
                        0x586d84898b96599b628c9f1201f09401e072702ee7496508538c1564a35a1299 ,
                        0x142e805a35e01a1951c73b05c5244b13951318469944a235736d02c47ad1689c ,
                        0x202a6a82d89b98585005057d20e158dfe592a4a48e5a761704f04b2bb3320220 ,
                        0x610ea2316f45480420035e02360c967ebb8a3164b58b5f65228a821888c2140a ,
                        0xf7183a51058240a2cb6a77ec255f93612919f49c3c074a97ff011a32aa148520 ,
                        0x1f98714b0ff895593a88b3394038411f67848efe624eaa691ac4e917f3a89a61 ,
                        0x99aa8e818df1621ad3793cc678792f951270d400cb3368eec149674a417c91a3 ,
                        0xc5d311027116b5d5428619b1dd430d76914dc8c1201f295154a03eb9653e7544 ,
                        0x2c3a10658a263c434194adb474f2211904c387340b1f32b229e96400fcd41f42 ,
                        0x322f30c25a7d065396581246c958a9d713f1857438b1730b135f03409f46d61e ,
                        0xde681210e7183ae21767c43706d7449b122e44e2132dba37cf735f55f094ad74 ,
                        0xa6c2461d6c0800070018976b64db95bbd281a03c1a3f735410112b266c225770 ,
                        0x2b4dc1903233241689599867e22067710098698cd9913c0effb933cd1834e495 ,
                        0x3882112e40c32903e812c71800c413944754b39df162bf3340406328f441a1cb ,
                        0xb6a83e08729f24689eda53cf032e5c198d3812b8797334323aa1893514fce0b5 ,
                        0x4ac944124864db850083200bc107b53e566a82b06a4383390eb06b93d28c4904 ,
                        0x7a55c003d1e76abfe7273a9a8dd7a7031b803f50b866dd7316b4030845580082 ,
                        0x1962b4d51cba485b401cc440ac8bb3330d69f111f0881295b22ada88a24714bf ,
                        0x93112993a3695e5b408b8224f351694dd98ce38b23e9941939032dcf19863919 ,
                        0x0428c15a04d041456394a2f11fd3e582885a05a40b120e4046e771ba80616b23 ,
                        0x77955a0c96defbc1cec370c7478e86069570f96a5efa3c4affd94a0b904ea531 ,
                        0x86ac314db3e851f2130c0c70c9edc6000e404d12c16d0c60c957f0c9972cca95 ,
                        0x4ccaa66cc9a8cc000970c91c352d06800fdc12cb53660f3db500dc6200b6fc53 ,
                        0xb71ccbb61c7d31b1003f55cbef92cbb6cccbf6e200fc10c2ef32ccc6cccbccdc ,
                        0x007a81353ce800bdbc00aeb61ea49417aba66c0ed853b22b6b0b8025b4bc0085 ,
                        0x43008886cbcaaa1685b000c786cd6bd4adbdc7cd15c82de1a2160e907caf7679 ,
                        0x06e02ff04ccd0d5000e8f8bb6ac10f7d9c17210c3006a0bbac324a3f647fe317 ,
                        0x44dc8748d9d47ddaf77d14fdd0187d8b7a33421cad14043044ab64b62e5c64df ,
                        0xc734c34ab6245d41db1415a684533bab1ea3a4398a74176b6805d1dfc7444928 ,
                        0x754598cd528784d60927ddf7d3ac316a0c43d451018d4c81d414a435f5151201 ,
                        0x64d3275d9f8486747f4713dfa53d97b4d123bd69a81a8402b42307da3022bd1a ,
                        0x754a9a5b2d7525d1d465bd8a4fa16539aad42f793745c38c9fb21eb7760f1d04 ,
                        0xc9ff1dddc2dc1420826bd64db1d68b338401124989d44a638da559244bde7ad4 ,
                        0x619d145214493b6a877a038d775162e812082c6274611b1233ddd17f9d255806 ,
                        0x5b82ed1497841a227d3ba57d224228639ed4c292ed140e8468fce011b13d1208 ,
                        0xc418a968a0b2fd935086398774d09f92014a00a7221dda9ae3da2a4d12a40d12 ,
                        0xabad144bf3dcaa24dd067a3bd1c1d0cb6d61c6f35953c517d6e74852f6d5a801 ,
                        0xdc395aa71ea18efc0003deedd57a51a74a0dd801a4a637da4397d927463d20cf ,
                        0x5dd64ddd49dbd4d465bbd2aa21dd5c43db002e12a918112536d0fcfdadd8cd83 ,
                        0xce1d2d8d64ced271dd318de079b3d653469a52b4d61bb4df8c94d803924d0bbe ,
                        0x4dffd2e4e0e6021102bed725f1dc52a4e0debadf2d8ed6f549e178e3d92a0d11 ,
                        0x4543e028aed258362dcb9dd63c3ee3db84e1008ea3fc2d11fafde30fced4fccd ,
                        0xe1cb1646f7ade21e6d134a8499f53d12bde4818b9dd24b21d8664b4feb7daaf4 ,
                        0x94e325cddce251dfea58e1655e488754ddd09d394cb44aff67e32211e78d1d46 ,
                        0x5395e203a2e63cadd486bd6c743a53581a42484d940078aa4df1de6b1ee34bed ,
                        0xd53a7ad950d1520d83dc7748e933d3d177c37d75ba342097d38b84868a2d45f1 ,
                        0xddd7d9f6e488ee990d547298434fd9e4d34c24ea71d84ba6ead48d0de8c83b40 ,
                        0xd0a217b9c3278a2de50282b01ece3077032dd1a44d45b3d3aa5334ad4e4d39ac ,
                        0x37ff4a54340e6461d6a94d48084911617f15b4de4d78d3956ddbaa1444228eec ,
                        0xa20ed4df2a45e14e84d90435aa134d7668e7b8f3ea3da437c40eee83578412be ,
                        0x6cef9735ae22b8964eebf58ed620ee3028824856a34ddbe6003ab05205500809 ,
                        0x8075f8a7030ee0000d7f0011af51cf953b05b0c990aef1115108f0b916ff8392 ,
                        0xffc3181eaf84c05bf10a5f08081ff11d214a975d005780432bcff2db6664f0b3 ,
                        0xf0094076d8ac8a128ff06b55085dc7f2aa69792bf53eab738998630823ff3e5b ,
                        0x15f145aff420cfee8570431e0ff13ae0341ae5f38dc4e401cfe56dddd5dab343 ,
                        0xb6fcf30be06d32b0550c60cba9d36d65afc909b0006d25030bb001b753000b10 ,
                        0x5cff10e100f510f3d9440df5b02130a003ce4584570098126100f570a4d97400 ,
                        0x65ef6d735f000c20f71a9975f5f05c30904d57e07e8e7fc96cb1006717350c20 ,
                        0x5232d056779f00f5877f0790f6d470f71dc100f590f6a15f08d47043917f76f5 ,
                        0xa703777f7dc14008174f0085c00084e0f81b120c0b50001bb0001abf46cfe5f7 ,
                        0xb09ff08def6d74e730a5dde0d8ddf5b50ee45a53dbc65f0f9e8f43977c761c71 ,
                        0xc93b6417dffff938741c5ad77deade5637d4fa89bf78824f0f7caf0339400856 ,
                        0x03f91bb955324008c1b0fe0041ed008103c10e146010ac00c1828464102800a0 ,
                        0x10357a0474d49b5740638160f5182c2c50ef60810432741040490000c47a1b08 ,
                        0x6cffa097e0800c85042db62cb000a38e1e1ce951835840063d0641172cd051e0 ,
                        0xc0061dfc822d38308f5e3d8d04120a6530f225bd0303192cd80040ec58b265cd ,
                        0x9e456b56655ab66bd9be851b572e5943637b00989920a157191baa3260e057e3 ,
                        0x15400317a2d4c155e55aa55f8116d021434682850c687e7ce99141526a805096 ,
                        0x0c464de33c91108322bcd275e10619f50cfa9471909ee0cf473fe65d5d1a2580 ,
                        0xbb425fd324405335cb8c9665689c490d33689305589f44f9f761496ac7111a14 ,
                        0x1e74a65f888064cc051f5efc78f2e5d33a7fb841a4d205449316aa473369827a ,
                        0xca132c4e0035e2588b97751c702d2189a891a99efb362010b0020029cab70394 ,
                        0xe3ff6881b1921a6801ff08d02b98d8224b201807429201b042080826988b9812 ,
                        0x28a2920e58692c8d64002498af94922c210708c229a1ab940bed4583628c1022 ,
                        0x89367a8aa3d842f4ad007a4264e04307628b2e187acca3b24a2baf444ba5fdc4 ,
                        0x2260016a368049a11e5eaaae90854ae2ea291d0c0260217be8818ebfc80ab882 ,
                        0xa6e312c2909e79e6914120d60e72ad33029e0c660305350caa810514fa6a8304 ,
                        0x4eba22b084643c203002e801e4d1ca8aa2ef3995fe6cd1c5ea2cec6a26061278 ,
                        0x14a2d972d2213f0d3790b1277a0a6580aa92aa6ac032ab3efa738345850bf1d1 ,
                        0x93ecb1aea78eb044365965e7e287ac1e0a916c01cb1ad408a6a44613492fe4f0 ,
                        0x8bff6dcb361ffc8f1a0383637456c9b8b296a42519f528c47a74383128067db3 ,
                        0x89245eeb7977a6a456ad4a46946a6d0e8015b9e40d86903272adabd90c2368d5 ,
                        0xe66abdc248021a20eaabbe74008bc8a26644502307eac3ee358d2e3aa8a74597 ,
                        0x45396595c7e47245880e88c98142352ca4d035057a759e03461c144eb7c4bae8 ,
                        0xa40422b6ca4419cc84682faa3e9b9535977c4468367825bb0e4c97229befc3c8 ,
                        0x0a79d066bf340469281d56526f20b35cc3b02f025cb35aa31cefd3a1011d0c09 ,
                        0x6deda4ba7b88013d0fb8f9805be9fbab9a60d606a46d1daad9c06f40e8f15665 ,
                        0xc71f372f227e763d8da68e28fa9a9e25fd4a40f3a99202e08af5c8b268b3d3ae ,
                        0x9357b1801e8ced6b3d964a7c08242935ea68c9be2af5dcd605346ab3a7cd94f4 ,
                        0x3cb6188382a9aa81c5baab4de082eb6b28cdeb31b3251d4e0daa4d19c1dea8f3 ,
                        0x81a4ac675d024427514c00223388a2e7978aa19a9f623b2020003b
                    End

                    LayoutCachedLeft =90
                    LayoutCachedWidth =3618
                    LayoutCachedHeight =1440
                    TabIndex =5
                End
                Begin Label
                    TextFontCharSet =186
                    TextAlign =2
                    TextFontFamily =34
                    Left =90
                    Top =1095
                    Width =2520
                    Height =780
                    FontSize =34
                    FontWeight =700
                    Name ="Label231"
                    Caption ="INVOICE"
                    LayoutCachedLeft =90
                    LayoutCachedTop =1095
                    LayoutCachedWidth =2610
                    LayoutCachedHeight =1875
                End
                Begin Rectangle
                    Top =2505
                    Width =14370
                    Height =285
                    BorderColor =-2147483615
                    Name ="Box146"
                    LayoutCachedTop =2505
                    LayoutCachedWidth =14370
                    LayoutCachedHeight =2790
                    BackThemeColorIndex =2
                End
                Begin Label
                    TextFontCharSet =186
                    TextAlign =2
                    TextFontFamily =34
                    Left =12405
                    Top =2310
                    Width =1485
                    Height =225
                    FontSize =8
                    FontWeight =700
                    Name ="Label107"
                    Caption ="┌─── Deposit ────┐"
                    LayoutCachedLeft =12405
                    LayoutCachedTop =2310
                    LayoutCachedWidth =13890
                    LayoutCachedHeight =2535
                    ForeThemeColorIndex =2
                End
                Begin Label
                    TextFontCharSet =186
                    TextAlign =2
                    TextFontFamily =34
                    Left =30
                    Top =2520
                    Width =2203
                    Height =240
                    FontSize =10
                    FontWeight =700
                    ForeColor =16777215
                    Name ="Label135"
                    Caption ="VIN"
                    GroupTable =1
                    LayoutCachedLeft =30
                    LayoutCachedTop =2520
                    LayoutCachedWidth =2233
                    LayoutCachedHeight =2760
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    TextFontCharSet =186
                    TextAlign =2
                    TextFontFamily =34
                    Left =2295
                    Top =2520
                    Width =5256
                    Height =240
                    FontSize =10
                    FontWeight =700
                    ForeColor =16777215
                    Name ="Label136"
                    Caption ="Unit Description"
                    GroupTable =1
                    LayoutCachedLeft =2295
                    LayoutCachedTop =2520
                    LayoutCachedWidth =7551
                    LayoutCachedHeight =2760
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    TextFontCharSet =186
                    TextFontFamily =34
                    Left =7605
                    Top =2520
                    Width =1905
                    Height =240
                    FontSize =10
                    FontWeight =700
                    ForeColor =16777215
                    Name ="Label140"
                    Caption ="Location"
                    GroupTable =1
                    LayoutCachedLeft =7605
                    LayoutCachedTop =2520
                    LayoutCachedWidth =9510
                    LayoutCachedHeight =2760
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    TextFontCharSet =186
                    TextAlign =2
                    TextFontFamily =34
                    Left =10455
                    Top =2520
                    Width =1320
                    Height =240
                    FontSize =10
                    FontWeight =700
                    ForeColor =16777215
                    Name ="Label141"
                    Caption ="Sale Price"
                    GroupTable =1
                    RightPadding =38
                    LayoutCachedLeft =10455
                    LayoutCachedTop =2520
                    LayoutCachedWidth =11775
                    LayoutCachedHeight =2760
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    TextFontCharSet =186
                    TextAlign =2
                    TextFontFamily =34
                    Left =9570
                    Top =2520
                    Width =825
                    Height =240
                    FontSize =10
                    FontWeight =700
                    ForeColor =16777215
                    Name ="Label142"
                    Caption ="Taxrate"
                    GroupTable =1
                    LayoutCachedLeft =9570
                    LayoutCachedTop =2520
                    LayoutCachedWidth =10395
                    LayoutCachedHeight =2760
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    TextFontCharSet =186
                    TextAlign =2
                    TextFontFamily =34
                    Left =11850
                    Top =2520
                    Width =1305
                    Height =240
                    FontSize =10
                    FontWeight =700
                    ForeColor =16777215
                    Name ="Label143"
                    Caption ="Amount"
                    GroupTable =1
                    LayoutCachedLeft =11850
                    LayoutCachedTop =2520
                    LayoutCachedWidth =13155
                    LayoutCachedHeight =2760
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    TextFontCharSet =186
                    TextAlign =2
                    TextFontFamily =34
                    Left =13215
                    Top =2520
                    Width =1065
                    Height =240
                    FontSize =10
                    FontWeight =700
                    ForeColor =16777215
                    Name ="Label144"
                    Caption =" Date"
                    GroupTable =1
                    LayoutCachedLeft =13215
                    LayoutCachedTop =2520
                    LayoutCachedWidth =14280
                    LayoutCachedHeight =2760
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    Visible = NotDefault
                    TextFontCharSet =186
                    TextAlign =1
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Left =15
                    Top =2235
                    Width =2715
                    Height =270
                    FontSize =12
                    FontWeight =700
                    TabIndex =3
                    Name ="Text90"
                    ControlSource ="=IIf([Forms]![ReMarketInvoiceManager]![PrtInv]=True,\"Invoice for Purchase of Un"
                        "its Listed Below\",\"Sale Purchase Order for Units Listed Below\")"

                    LayoutCachedLeft =15
                    LayoutCachedTop =2235
                    LayoutCachedWidth =2730
                    LayoutCachedHeight =2505
                End
                Begin TextBox
                    OldBorderStyle =1
                    TextFontCharSet =186
                    TextAlign =2
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Left =12120
                    Top =1125
                    Width =1590
                    Height =315
                    TabIndex =2
                    BorderColor =4210752
                    Name ="Text169"
                    ControlSource ="InvoiceDueDate"
                    GroupTable =15
                    RightPadding =38
                    BottomPadding =38

                    LayoutCachedLeft =12120
                    LayoutCachedTop =1125
                    LayoutCachedWidth =13710
                    LayoutCachedHeight =1440
                    RowStart =2
                    RowEnd =2
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =5
                    GroupTable =15
                    Begin
                        Begin Label
                            OldBorderStyle =1
                            TextFontCharSet =186
                            TextFontFamily =34
                            Left =10425
                            Top =1125
                            Width =1620
                            Height =315
                            FontWeight =700
                            BorderColor =4210752
                            Name ="Label170"
                            Caption =" Payment Due:"
                            GroupTable =15
                            RightPadding =38
                            BottomPadding =38
                            LayoutCachedLeft =10425
                            LayoutCachedTop =1125
                            LayoutCachedWidth =12045
                            LayoutCachedHeight =1440
                            RowStart =2
                            RowEnd =2
                            LayoutGroup =5
                            GroupTable =15
                        End
                    End
                End
                Begin TextBox
                    BorderWidth =1
                    TextFontCharSet =238
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Left =4095
                    Top =705
                    Width =5775
                    Height =1500
                    FontSize =12
                    TabIndex =4
                    BorderColor =4210752
                    Name ="clientCompanyName"
                    ControlSource ="BuyerAddr"

                    LayoutCachedLeft =4095
                    LayoutCachedTop =705
                    LayoutCachedWidth =9870
                    LayoutCachedHeight =2205
                End
                Begin Rectangle
                    BackStyle =0
                    BorderWidth =1
                    Left =4020
                    Top =510
                    Width =5955
                    Height =1740
                    BorderColor =4210752
                    Name ="Box167"
                    LayoutCachedLeft =4020
                    LayoutCachedTop =510
                    LayoutCachedWidth =9975
                    LayoutCachedHeight =2250
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    TextFontCharSet =186
                    TextAlign =2
                    TextFontFamily =34
                    Left =4170
                    Top =360
                    Width =1635
                    Height =255
                    FontSize =10
                    FontWeight =700
                    BorderColor =4210752
                    Name ="Label1"
                    Caption ="Buyer Information:"
                    LayoutCachedLeft =4170
                    LayoutCachedTop =360
                    LayoutCachedWidth =5805
                    LayoutCachedHeight =615
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
            Height =0
            Name ="GroupHeader0"
        End
        Begin Section
            KeepTogether = NotDefault
            Height =330
            Name ="Detail"
            Begin
                Begin TextBox
                    DecimalPlaces =2
                    OldBorderStyle =1
                    TextFontCharSet =238
                    TextAlign =3
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Left =10455
                    Top =30
                    Width =1320
                    Height =255
                    FontSize =10
                    TabIndex =4
                    BorderColor =4210752
                    Name ="SoldPrice"
                    ControlSource ="SoldPrice"
                    Format ="$#,##0.00;($#,##0.00)"
                    GroupTable =1
                    RightPadding =38
                    BottomPadding =38

                    LayoutCachedLeft =10455
                    LayoutCachedTop =30
                    LayoutCachedWidth =11775
                    LayoutCachedHeight =285
                    RowStart =1
                    RowEnd =1
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    TextFontCharSet =238
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Left =30
                    Top =30
                    Width =2203
                    Height =255
                    ColumnWidth =2085
                    FontSize =10
                    Name ="UnitVIN"
                    ControlSource ="VIN"
                    StatusBarText ="VIN number must be unique in table"
                    GroupTable =1
                    BottomPadding =38

                    LayoutCachedLeft =30
                    LayoutCachedTop =30
                    LayoutCachedWidth =2233
                    LayoutCachedHeight =285
                    RowStart =1
                    RowEnd =1
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    DecimalPlaces =3
                    TextFontCharSet =238
                    TextAlign =2
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Left =9570
                    Top =30
                    Width =825
                    Height =255
                    FontSize =10
                    TabIndex =3
                    Name ="Text66"
                    ControlSource ="=[Forms]![ReMarketInvoiceManager]![UseTxRt]"
                    Format ="Percent"
                    GroupTable =1
                    BottomPadding =38

                    LayoutCachedLeft =9570
                    LayoutCachedTop =30
                    LayoutCachedWidth =10395
                    LayoutCachedHeight =285
                    RowStart =1
                    RowEnd =1
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    TextFontCharSet =238
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Left =7605
                    Top =30
                    Width =1905
                    Height =255
                    FontSize =10
                    TabIndex =2
                    Name ="Location"
                    ControlSource ="=[LCity] & \" \" & [LState]"
                    GroupTable =1
                    BottomPadding =38

                    LayoutCachedLeft =7605
                    LayoutCachedTop =30
                    LayoutCachedWidth =9510
                    LayoutCachedHeight =285
                    RowStart =1
                    RowEnd =1
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    TextFontCharSet =238
                    TextAlign =3
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Left =11850
                    Top =30
                    Width =1305
                    Height =255
                    FontSize =10
                    TabIndex =7
                    Name ="RmktDeposit"
                    ControlSource ="Deposit"
                    Format ="$#,##0.00;($#,##0.00)"
                    GroupTable =1
                    BottomPadding =38

                    LayoutCachedLeft =11850
                    LayoutCachedTop =30
                    LayoutCachedWidth =13155
                    LayoutCachedHeight =285
                    RowStart =1
                    RowEnd =1
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    TextFontCharSet =238
                    TextAlign =2
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Left =13215
                    Top =30
                    Width =1065
                    Height =255
                    FontSize =10
                    TabIndex =8
                    Name ="RmktDepdt"
                    ControlSource ="DepositDate"
                    GroupTable =1
                    BottomPadding =38
                    ShowDatePicker =1

                    LayoutCachedLeft =13215
                    LayoutCachedTop =30
                    LayoutCachedWidth =14280
                    LayoutCachedHeight =285
                    RowStart =1
                    RowEnd =1
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    TextFontCharSet =238
                    TextAlign =1
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Left =2295
                    Top =30
                    Width =5256
                    Height =255
                    FontSize =10
                    TabIndex =1
                    Name ="DescriptionLong"
                    ControlSource ="DescriptionLong"
                    GroupTable =1
                    BottomPadding =38

                    LayoutCachedLeft =2295
                    LayoutCachedTop =30
                    LayoutCachedWidth =7551
                    LayoutCachedHeight =285
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    Visible = NotDefault
                    IMESentenceMode =3
                    Left =9645
                    Top =300
                    Width =1320
                    Height =0
                    TabIndex =5
                    Name ="Shipping"
                    ControlSource ="Shipping"

                    LayoutCachedLeft =9645
                    LayoutCachedTop =300
                    LayoutCachedWidth =10965
                    LayoutCachedHeight =300
                    Begin
                        Begin Label
                            Visible = NotDefault
                            Left =7950
                            Top =300
                            Width =1635
                            Height =0
                            Name ="Label124"
                            Caption ="Shipping:"
                            LayoutCachedLeft =7950
                            LayoutCachedTop =300
                            LayoutCachedWidth =9585
                            LayoutCachedHeight =300
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    IMESentenceMode =3
                    Left =9645
                    Top =330
                    Width =1320
                    Height =0
                    TabIndex =6
                    Name ="Discount"
                    ControlSource ="Discount"

                    LayoutCachedLeft =9645
                    LayoutCachedTop =330
                    LayoutCachedWidth =10965
                    LayoutCachedHeight =330
                    Begin
                        Begin Label
                            Visible = NotDefault
                            Left =7935
                            Top =330
                            Width =1650
                            Height =0
                            Name ="Label131"
                            Caption ="Discount:"
                            LayoutCachedLeft =7935
                            LayoutCachedTop =330
                            LayoutCachedWidth =9585
                            LayoutCachedHeight =330
                        End
                    End
                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =0
            Name ="GroupFooter1"
        End
        Begin PageFooter
            Height =300
            Name ="PageFooterSection"
            Begin
                Begin TextBox
                    TextFontCharSet =238
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Left =30
                    Top =45
                    Width =8985
                    Height =210
                    FontSize =9
                    Name ="Text45"
                    ControlSource ="=\"Printed on \" & Date() & \" @ \" & Time()"

                    LayoutCachedLeft =30
                    LayoutCachedTop =45
                    LayoutCachedWidth =9015
                    LayoutCachedHeight =255
                End
                Begin TextBox
                    TextFontCharSet =238
                    TextAlign =3
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Left =11640
                    Top =45
                    Width =2610
                    Height =210
                    FontSize =9
                    TabIndex =1
                    Name ="Text47"
                    ControlSource ="=\"Page \" & [Page] & \" of \" & [Pages]"

                    LayoutCachedLeft =11640
                    LayoutCachedTop =45
                    LayoutCachedWidth =14250
                    LayoutCachedHeight =255
                End
            End
        End
        Begin FormFooter
            KeepTogether = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =3465
            Name ="ReportFooter"
            Begin
                Begin TextBox
                    DecimalPlaces =2
                    OldBorderStyle =1
                    TextFontCharSet =186
                    TextAlign =3
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =10470
                    Top =180
                    Width =1350
                    Height =274
                    FontSize =10
                    TabIndex =7
                    Name ="Text22"
                    ControlSource ="=Sum([SoldPrice])"
                    Format ="$#,##0.00;($#,##0.00)"
                    GroupTable =11
                    RightPadding =38
                    BottomPadding =38

                    LayoutCachedLeft =10470
                    LayoutCachedTop =180
                    LayoutCachedWidth =11820
                    LayoutCachedHeight =454
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =2
                    BorderThemeColorIndex =2
                    GroupTable =11
                    Begin
                        Begin Label
                            OldBorderStyle =1
                            TextFontCharSet =186
                            TextAlign =1
                            TextFontFamily =34
                            Left =8640
                            Top =180
                            Width =1755
                            Height =274
                            FontSize =10
                            BorderColor =4210752
                            Name ="Label147"
                            Caption =" Sub Total"
                            GroupTable =11
                            RightPadding =38
                            BottomPadding =38
                            LayoutCachedLeft =8640
                            LayoutCachedTop =180
                            LayoutCachedWidth =10395
                            LayoutCachedHeight =454
                            LayoutGroup =2
                            GroupTable =11
                        End
                    End
                End
                Begin TextBox
                    TextAlign =1
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Left =2220
                    Top =1935
                    Width =1860
                    Height =315
                    FontSize =10
                    FontWeight =700
                    TabIndex =9
                    Name ="Text82"
                    ControlSource ="InvoiceNumber"
                    GroupTable =13
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =2220
                    LayoutCachedTop =1935
                    LayoutCachedWidth =4080
                    LayoutCachedHeight =2250
                    RowStart =4
                    RowEnd =4
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =3
                    GroupTable =13
                    Begin
                        Begin Label
                            TextAlign =1
                            TextFontFamily =34
                            Left =960
                            Top =1935
                            Width =1260
                            Height =315
                            FontSize =10
                            Name ="Label83"
                            Caption ="Invoice #:"
                            GroupTable =13
                            LeftPadding =0
                            TopPadding =0
                            RightPadding =0
                            BottomPadding =0
                            LayoutCachedLeft =960
                            LayoutCachedTop =1935
                            LayoutCachedWidth =2220
                            LayoutCachedHeight =2250
                            RowStart =4
                            RowEnd =4
                            LayoutGroup =3
                            GroupTable =13
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =2
                    OldBorderStyle =1
                    TextFontCharSet =186
                    TextAlign =3
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =10470
                    Top =525
                    Width =1350
                    Height =270
                    FontSize =10
                    TabIndex =1
                    Name ="Text75"
                    ControlSource ="=(Sum([SoldPrice]))*Forms!ReMarketInvoiceManager!UseTxRt"
                    Format ="$#,##0.00;($#,##0.00)"
                    GroupTable =11
                    RightPadding =38
                    BottomPadding =38

                    LayoutCachedLeft =10470
                    LayoutCachedTop =525
                    LayoutCachedWidth =11820
                    LayoutCachedHeight =795
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =2
                    BorderThemeColorIndex =2
                    GroupTable =11
                    Begin
                        Begin Label
                            OldBorderStyle =1
                            TextFontCharSet =186
                            TextAlign =1
                            TextFontFamily =34
                            Left =8640
                            Top =525
                            Width =1755
                            Height =270
                            FontSize =10
                            BorderColor =4210752
                            Name ="Label148"
                            Caption =" Tax Amount:"
                            GroupTable =11
                            RightPadding =38
                            BottomPadding =38
                            LayoutCachedLeft =8640
                            LayoutCachedTop =525
                            LayoutCachedWidth =10395
                            LayoutCachedHeight =795
                            RowStart =1
                            RowEnd =1
                            LayoutGroup =2
                            GroupTable =11
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =2
                    OldBorderStyle =1
                    TextFontCharSet =186
                    TextAlign =3
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =10470
                    Top =960
                    Width =1350
                    Height =270
                    FontSize =10
                    TabIndex =6
                    Name ="Text76"
                    ControlSource ="=Sum([SoldPrice])+(Sum([SoldPrice]))*Forms!ReMarketInvoiceManager!UseTxRt"
                    Format ="$#,##0.00;($#,##0.00)"
                    GroupTable =11
                    RightPadding =38
                    BottomPadding =38

                    LayoutCachedLeft =10470
                    LayoutCachedTop =960
                    LayoutCachedWidth =11820
                    LayoutCachedHeight =1230
                    RowStart =3
                    RowEnd =3
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =2
                    BorderThemeColorIndex =2
                    GroupTable =11
                    Begin
                        Begin Label
                            OldBorderStyle =1
                            TextFontCharSet =186
                            TextAlign =1
                            TextFontFamily =34
                            Left =8640
                            Top =960
                            Width =1755
                            Height =270
                            FontSize =10
                            BorderColor =4210752
                            Name ="Label149"
                            Caption =" Total Sale Price:"
                            GroupTable =11
                            RightPadding =38
                            BottomPadding =38
                            LayoutCachedLeft =8640
                            LayoutCachedTop =960
                            LayoutCachedWidth =10395
                            LayoutCachedHeight =1230
                            RowStart =3
                            RowEnd =3
                            LayoutGroup =2
                            GroupTable =11
                        End
                    End
                End
                Begin Label
                    FontUnderline = NotDefault
                    TextFontCharSet =238
                    TextFontFamily =34
                    Left =1005
                    Top =465
                    Width =2925
                    Height =315
                    FontSize =12
                    FontWeight =700
                    ForeColor =-2147483615
                    Name ="Label37"
                    Caption ="Payment Via Wire Transfer:"
                    LayoutCachedLeft =1005
                    LayoutCachedTop =465
                    LayoutCachedWidth =3930
                    LayoutCachedHeight =780
                End
                Begin Label
                    TextAlign =1
                    TextFontFamily =34
                    Left =2220
                    Top =840
                    Width =1860
                    Height =240
                    FontSize =10
                    FontWeight =700
                    Name ="Label48"
                    Caption ="Fifth Third Bank"
                    GroupTable =13
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =2220
                    LayoutCachedTop =840
                    LayoutCachedWidth =4080
                    LayoutCachedHeight =1080
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =3
                    GroupTable =13
                End
                Begin Label
                    TextAlign =1
                    TextFontFamily =34
                    Left =2220
                    Top =1080
                    Width =1860
                    Height =285
                    FontSize =10
                    FontWeight =700
                    Name ="Label49"
                    Caption ="Fleet Advantage, LLC"
                    GroupTable =13
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =2220
                    LayoutCachedTop =1080
                    LayoutCachedWidth =4080
                    LayoutCachedHeight =1365
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =3
                    GroupTable =13
                End
                Begin Label
                    TextAlign =1
                    TextFontFamily =34
                    Left =2220
                    Top =1650
                    Width =1860
                    Height =285
                    FontSize =10
                    FontWeight =700
                    Name ="Label50"
                    Caption ="042000314"
                    GroupTable =13
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =2220
                    LayoutCachedTop =1650
                    LayoutCachedWidth =4080
                    LayoutCachedHeight =1935
                    RowStart =3
                    RowEnd =3
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =3
                    GroupTable =13
                End
                Begin Label
                    TextAlign =1
                    TextFontFamily =34
                    Left =2220
                    Top =1365
                    Width =1860
                    Height =285
                    FontSize =10
                    FontWeight =700
                    Name ="Label51"
                    Caption ="7432859127"
                    GroupTable =13
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =2220
                    LayoutCachedTop =1365
                    LayoutCachedWidth =4080
                    LayoutCachedHeight =1650
                    RowStart =2
                    RowEnd =2
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =3
                    GroupTable =13
                End
                Begin Rectangle
                    BackStyle =0
                    BorderWidth =1
                    Left =765
                    Top =285
                    Width =5955
                    Height =2085
                    BorderColor =4210752
                    Name ="Box104"
                    LayoutCachedLeft =765
                    LayoutCachedTop =285
                    LayoutCachedWidth =6720
                    LayoutCachedHeight =2370
                End
                Begin Label
                    Visible = NotDefault
                    FontUnderline = NotDefault
                    TextFontCharSet =238
                    TextFontFamily =34
                    Left =5940
                    Top =630
                    Width =2595
                    Height =315
                    FontSize =12
                    FontWeight =700
                    ForeColor =-2147483615
                    Name ="Label38"
                    Caption ="Payment Via Check Mail To:"
                    GroupTable =14
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =5940
                    LayoutCachedTop =630
                    LayoutCachedWidth =8535
                    LayoutCachedHeight =945
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =4
                    GroupTable =14
                End
                Begin Label
                    Visible = NotDefault
                    TextFontCharSet =186
                    TextAlign =1
                    TextFontFamily =34
                    Left =5940
                    Top =945
                    Width =2595
                    Height =285
                    FontWeight =700
                    Name ="Label41"
                    Caption ="Fleet Advantage, LLC"
                    GroupTable =14
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =5940
                    LayoutCachedTop =945
                    LayoutCachedWidth =8535
                    LayoutCachedHeight =1230
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =4
                    GroupTable =14
                End
                Begin Label
                    Visible = NotDefault
                    TextFontCharSet =186
                    TextAlign =1
                    TextFontFamily =34
                    Left =5940
                    Top =1230
                    Width =2595
                    Height =285
                    FontWeight =700
                    Name ="Label42"
                    Caption ="401 East Las Olas Blvd. Suite 1720"
                    GroupTable =14
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =5940
                    LayoutCachedTop =1230
                    LayoutCachedWidth =8535
                    LayoutCachedHeight =1515
                    RowStart =2
                    RowEnd =2
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =4
                    GroupTable =14
                End
                Begin Label
                    Visible = NotDefault
                    TextFontCharSet =186
                    TextAlign =1
                    TextFontFamily =34
                    Left =5940
                    Top =1515
                    Width =2595
                    Height =285
                    FontWeight =700
                    Name ="Label43"
                    Caption ="Attn: Accounts Receivable"
                    GroupTable =14
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =5940
                    LayoutCachedTop =1515
                    LayoutCachedWidth =8535
                    LayoutCachedHeight =1800
                    RowStart =3
                    RowEnd =3
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =4
                    GroupTable =14
                End
                Begin Label
                    Visible = NotDefault
                    TextFontCharSet =186
                    TextAlign =1
                    TextFontFamily =34
                    Left =5940
                    Top =1800
                    Width =2595
                    Height =285
                    FontWeight =700
                    Name ="Label44"
                    Caption ="Fort Lauderdale, FL 33301"
                    GroupTable =14
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =5940
                    LayoutCachedTop =1800
                    LayoutCachedWidth =8535
                    LayoutCachedHeight =2085
                    RowStart =4
                    RowEnd =4
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =4
                    GroupTable =14
                End
                Begin TextBox
                    DecimalPlaces =2
                    OldBorderStyle =1
                    TextFontCharSet =186
                    TextAlign =3
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =10470
                    Top =1725
                    Width =1350
                    Height =270
                    FontSize =10
                    TabIndex =3
                    Name ="Text110"
                    ControlSource ="=Sum([Deposit])"
                    Format ="$#,##0.00;($#,##0.00)"
                    GroupTable =11
                    RightPadding =38
                    BottomPadding =38

                    LayoutCachedLeft =10470
                    LayoutCachedTop =1725
                    LayoutCachedWidth =11820
                    LayoutCachedHeight =1995
                    RowStart =6
                    RowEnd =6
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =2
                    BorderThemeColorIndex =2
                    GroupTable =11
                End
                Begin TextBox
                    DecimalPlaces =2
                    OldBorderStyle =1
                    TextFontCharSet =186
                    TextAlign =3
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =10470
                    Top =2070
                    Width =1350
                    Height =270
                    FontSize =10
                    TabIndex =4
                    Name ="Text118"
                    ControlSource ="=(Sum([SoldPrice])+(Sum([SoldPrice]))*[Forms]![ReMarketInvoiceManager]![UseTxRt]"
                        ")-Sum([Deposit])+Nz(Sum([Shipping]),0)-Nz(Sum([Discount]),0)"
                    Format ="$#,##0.00;($#,##0.00)"
                    GroupTable =11
                    RightPadding =38
                    BottomPadding =38

                    LayoutCachedLeft =10470
                    LayoutCachedTop =2070
                    LayoutCachedWidth =11820
                    LayoutCachedHeight =2340
                    RowStart =7
                    RowEnd =7
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =2
                    BorderThemeColorIndex =2
                    GroupTable =11
                    Begin
                        Begin Label
                            OldBorderStyle =1
                            TextFontCharSet =186
                            TextAlign =1
                            TextFontFamily =34
                            Left =8640
                            Top =2070
                            Width =1755
                            Height =270
                            FontSize =10
                            BorderColor =4210752
                            Name ="Label153"
                            Caption =" Remaining Balance:"
                            GroupTable =11
                            RightPadding =38
                            BottomPadding =38
                            LayoutCachedLeft =8640
                            LayoutCachedTop =2070
                            LayoutCachedWidth =10395
                            LayoutCachedHeight =2340
                            RowStart =7
                            RowEnd =7
                            LayoutGroup =2
                            GroupTable =11
                        End
                    End
                End
                Begin Line
                    OldBorderStyle =6
                    BorderLineStyle =5
                    Left =45
                    Top =60
                    Width =14025
                    Name ="Line121"
                    LayoutCachedLeft =45
                    LayoutCachedTop =60
                    LayoutCachedWidth =14070
                    LayoutCachedHeight =60
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    DecimalPlaces =2
                    OldBorderStyle =1
                    TextFontCharSet =186
                    TextAlign =3
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =10470
                    Top =1650
                    Width =1350
                    Height =0
                    FontSize =10
                    TabIndex =2
                    Name ="discountamt"
                    ControlSource ="=(IIf(Sum([discount])=0,Null,Sum([discount])))*-1"
                    Format ="$#,##0.00;($#,##0.00)"
                    GroupTable =11
                    RightPadding =38
                    BottomPadding =38

                    LayoutCachedLeft =10470
                    LayoutCachedTop =1650
                    LayoutCachedWidth =11820
                    LayoutCachedHeight =1650
                    RowStart =5
                    RowEnd =5
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =2
                    BorderThemeColorIndex =2
                    GroupTable =11
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    DecimalPlaces =2
                    OldBorderStyle =1
                    TextFontCharSet =186
                    TextAlign =3
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =10470
                    Top =1305
                    Width =1350
                    Height =270
                    FontSize =10
                    TabIndex =5
                    Name ="shipingcost"
                    ControlSource ="=CCur(nz(Sum([shipping]),0))"
                    Format ="$#,##0.00;($#,##0.00)"
                    GroupTable =11
                    RightPadding =38
                    BottomPadding =38

                    LayoutCachedLeft =10470
                    LayoutCachedTop =1305
                    LayoutCachedWidth =11820
                    LayoutCachedHeight =1575
                    RowStart =4
                    RowEnd =4
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =2
                    BorderThemeColorIndex =2
                    GroupTable =11
                    Begin
                        Begin Label
                            OldBorderStyle =1
                            TextFontCharSet =186
                            TextAlign =1
                            TextFontFamily =34
                            Left =8640
                            Top =1305
                            Width =1755
                            Height =270
                            FontSize =10
                            BorderColor =4210752
                            Name ="Label150"
                            Caption =" Shipping"
                            GroupTable =11
                            RightPadding =38
                            BottomPadding =38
                            LayoutCachedLeft =8640
                            LayoutCachedTop =1305
                            LayoutCachedWidth =10395
                            LayoutCachedHeight =1575
                            RowStart =4
                            RowEnd =4
                            LayoutGroup =2
                            GroupTable =11
                        End
                    End
                End
                Begin TextBox
                    TextFontCharSet =186
                    TextAlign =1
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Left =8640
                    Top =1650
                    Width =1755
                    Height =0
                    FontSize =10
                    TabIndex =11
                    BorderColor =4210752
                    Name ="Label151"
                    ControlSource ="=IIf(Sum([discount])=0,Null,\" Discount:\")"
                    GroupTable =11
                    RightPadding =38
                    BottomPadding =38

                    LayoutCachedLeft =8640
                    LayoutCachedTop =1650
                    LayoutCachedWidth =10395
                    LayoutCachedHeight =1650
                    RowStart =5
                    RowEnd =5
                    LayoutGroup =2
                    GroupTable =11
                End
                Begin TextBox
                    OldBorderStyle =1
                    TextFontCharSet =186
                    TextAlign =1
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Left =8640
                    Top =1725
                    Width =1755
                    Height =270
                    FontSize =10
                    TabIndex =10
                    BorderColor =4210752
                    Name ="Label152"
                    ControlSource ="=\" Deposit \" & IIf(IsNull([rmktdepdt]),\"Due:\",\"Paid:\")"
                    GroupTable =11
                    RightPadding =38
                    BottomPadding =38

                    LayoutCachedLeft =8640
                    LayoutCachedTop =1725
                    LayoutCachedWidth =10395
                    LayoutCachedHeight =1995
                    RowStart =6
                    RowEnd =6
                    LayoutGroup =2
                    GroupTable =11
                End
                Begin TextBox
                    OldBorderStyle =1
                    TextFontCharSet =186
                    TextAlign =3
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =10470
                    Top =870
                    Width =1350
                    Height =30
                    FontSize =10
                    TabIndex =8
                    BackColor =-2147483615
                    ForeColor =-2147483615
                    Name ="Text155"
                    GroupTable =11
                    RightPadding =38

                    LayoutCachedLeft =10470
                    LayoutCachedTop =870
                    LayoutCachedWidth =11820
                    LayoutCachedHeight =900
                    RowStart =2
                    RowEnd =2
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =2
                    BorderThemeColorIndex =2
                    GroupTable =11
                    Begin
                        Begin Label
                            BackStyle =1
                            OldBorderStyle =1
                            TextFontCharSet =186
                            TextAlign =1
                            TextFontFamily =34
                            Left =8640
                            Top =870
                            Width =1755
                            Height =30
                            FontSize =10
                            BackColor =-2147483615
                            Name ="Label156"
                            Caption =":"
                            GroupTable =11
                            RightPadding =38
                            LayoutCachedLeft =8640
                            LayoutCachedTop =870
                            LayoutCachedWidth =10395
                            LayoutCachedHeight =900
                            RowStart =2
                            RowEnd =2
                            LayoutGroup =2
                            BorderThemeColorIndex =2
                            GroupTable =11
                        End
                    End
                End
                Begin Label
                    TextAlign =1
                    TextFontFamily =34
                    Left =960
                    Top =840
                    Width =1260
                    Height =240
                    FontSize =10
                    Name ="Label157"
                    Caption ="Bank Name:"
                    GroupTable =13
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =960
                    LayoutCachedTop =840
                    LayoutCachedWidth =2220
                    LayoutCachedHeight =1080
                    LayoutGroup =3
                    GroupTable =13
                End
                Begin Label
                    TextAlign =1
                    TextFontFamily =34
                    Left =960
                    Top =1080
                    Width =1260
                    Height =285
                    FontSize =10
                    Name ="Label158"
                    Caption ="Account Name:"
                    GroupTable =13
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =960
                    LayoutCachedTop =1080
                    LayoutCachedWidth =2220
                    LayoutCachedHeight =1365
                    RowStart =1
                    RowEnd =1
                    LayoutGroup =3
                    GroupTable =13
                End
                Begin Label
                    TextAlign =1
                    TextFontFamily =34
                    Left =960
                    Top =1365
                    Width =1260
                    Height =285
                    FontSize =10
                    Name ="Label159"
                    Caption ="Account #:"
                    GroupTable =13
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =960
                    LayoutCachedTop =1365
                    LayoutCachedWidth =2220
                    LayoutCachedHeight =1650
                    RowStart =2
                    RowEnd =2
                    LayoutGroup =3
                    GroupTable =13
                End
                Begin Label
                    TextAlign =1
                    TextFontFamily =34
                    Left =960
                    Top =1650
                    Width =1260
                    Height =285
                    FontSize =10
                    Name ="Label160"
                    Caption ="ABA Routing #:"
                    GroupTable =13
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =960
                    LayoutCachedTop =1650
                    LayoutCachedWidth =2220
                    LayoutCachedHeight =1935
                    RowStart =3
                    RowEnd =3
                    LayoutGroup =3
                    GroupTable =13
                End
                Begin Label
                    TextFontCharSet =186
                    TextAlign =1
                    TextFontFamily =34
                    Left =5880
                    Top =945
                    Width =60
                    Height =285
                    Name ="Label161"
                    Caption =" "
                    GroupTable =14
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =5880
                    LayoutCachedTop =945
                    LayoutCachedWidth =5940
                    LayoutCachedHeight =1230
                    RowStart =1
                    RowEnd =1
                    LayoutGroup =4
                    GroupTable =14
                End
                Begin Label
                    TextFontCharSet =186
                    TextAlign =1
                    TextFontFamily =34
                    Left =5880
                    Top =1230
                    Width =60
                    Height =285
                    Name ="Label162"
                    Caption =" "
                    GroupTable =14
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =5880
                    LayoutCachedTop =1230
                    LayoutCachedWidth =5940
                    LayoutCachedHeight =1515
                    RowStart =2
                    RowEnd =2
                    LayoutGroup =4
                    GroupTable =14
                End
                Begin Label
                    TextFontCharSet =186
                    TextAlign =1
                    TextFontFamily =34
                    Left =5880
                    Top =1515
                    Width =60
                    Height =285
                    Name ="Label163"
                    Caption =" "
                    GroupTable =14
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =5880
                    LayoutCachedTop =1515
                    LayoutCachedWidth =5940
                    LayoutCachedHeight =1800
                    RowStart =3
                    RowEnd =3
                    LayoutGroup =4
                    GroupTable =14
                End
                Begin Label
                    TextFontCharSet =186
                    TextAlign =1
                    TextFontFamily =34
                    Left =5880
                    Top =1800
                    Width =60
                    Height =285
                    Name ="Label164"
                    Caption =" "
                    GroupTable =14
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =5880
                    LayoutCachedTop =1800
                    LayoutCachedWidth =5940
                    LayoutCachedHeight =2085
                    RowStart =4
                    RowEnd =4
                    LayoutGroup =4
                    GroupTable =14
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    TextFontCharSet =186
                    TextAlign =2
                    TextFontFamily =34
                    Left =870
                    Top =150
                    Width =1635
                    Height =240
                    FontSize =10
                    FontWeight =700
                    BorderColor =4210752
                    Name ="Label165"
                    Caption ="Remit Information:"
                    LayoutCachedLeft =870
                    LayoutCachedTop =150
                    LayoutCachedWidth =2505
                    LayoutCachedHeight =390
                End
                Begin Label
                    TextFontCharSet =186
                    TextAlign =1
                    TextFontFamily =34
                    Left =5880
                    Top =630
                    Width =60
                    Height =315
                    Name ="Label166"
                    Caption =" "
                    GroupTable =14
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =5880
                    LayoutCachedTop =630
                    LayoutCachedWidth =5940
                    LayoutCachedHeight =945
                    LayoutGroup =4
                    GroupTable =14
                End
                Begin TextBox
                    TextFontCharSet =186
                    TextAlign =1
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Left =420
                    Top =2520
                    Width =13845
                    Height =840
                    FontSize =10
                    Name ="Label105"
                    ControlSource ="=\"<i><b>Upon receipt of total sale price, the equipment titles will be released"
                        " and sent with the executed bill of sale via overnight courier</b>.<br>In additi"
                        "on, the appropriate contact at the equipment location will be notified to releas"
                        "e the units to your authorized representative.<br><b>Purchaser is responsible fo"
                        "r any taxes due, if any.\""
                    TextFormat =1

                    LayoutCachedLeft =420
                    LayoutCachedTop =2520
                    LayoutCachedWidth =14265
                    LayoutCachedHeight =3360
                End
            End
        End
    End
End
