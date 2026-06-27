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
    Width =15240
    DatasheetFontHeight =11
    ItemSuffix =448
    Left =5280
    Top =2715
    ShortcutMenuBar ="PrintMe"
    RecSrcDt = Begin
        0xa74686b8f33ce540
    End
    RecordSource ="SELECT vw_RentInvoice.* FROM vw_RentInvoice WHERE (((vw_RentInvoice.BillID)=[For"
        "ms]![Billing]![BillID]));"
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
        Begin EmptyCell
            Height =240
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin BreakLevel
            GroupHeader = NotDefault
            GroupFooter = NotDefault
            ControlSource ="Schedule"
        End
        Begin BreakLevel
            GroupHeader = NotDefault
            GroupFooter = NotDefault
            ControlSource ="UnitGroup"
        End
        Begin BreakLevel
            ControlSource ="UnitGroup"
        End
        Begin BreakLevel
            ControlSource ="Unitnum"
        End
        Begin BreakLevel
            SortOrder = NotDefault
            ControlSource ="BilledRent"
        End
        Begin FormHeader
            KeepTogether = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =2130
            Name ="ReportHeader"
            Begin
                Begin Image
                    Left =120
                    Top =480
                    Width =3528
                    Name ="Auto_Logo0"
                    Picture ="Smalllogo.gif"
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

                    LayoutCachedLeft =120
                    LayoutCachedTop =480
                    LayoutCachedWidth =3648
                    LayoutCachedHeight =1920
                    TabIndex =9
                End
                Begin Line
                    BorderWidth =1
                    Left =4020
                    Top =1365
                    Width =0
                    Height =576
                    BorderColor =8210719
                    Name ="Line437"
                    LayoutCachedLeft =4020
                    LayoutCachedTop =1365
                    LayoutCachedWidth =4020
                    LayoutCachedHeight =1941
                End
                Begin Line
                    BorderWidth =1
                    Left =4020
                    Top =1365
                    Width =4320
                    BorderColor =8210719
                    Name ="Line436"
                    LayoutCachedLeft =4020
                    LayoutCachedTop =1365
                    LayoutCachedWidth =8340
                    LayoutCachedHeight =1365
                End
                Begin Line
                    BorderWidth =1
                    Left =4020
                    Top =90
                    Width =0
                    Height =720
                    BorderColor =8210719
                    Name ="Line435"
                    LayoutCachedLeft =4020
                    LayoutCachedTop =90
                    LayoutCachedWidth =4020
                    LayoutCachedHeight =810
                End
                Begin Line
                    BorderWidth =1
                    Left =4020
                    Top =90
                    Width =4320
                    BorderColor =8210719
                    Name ="Line434"
                    LayoutCachedLeft =4020
                    LayoutCachedTop =90
                    LayoutCachedWidth =8340
                    LayoutCachedHeight =90
                End
                Begin TextBox
                    TextAlign =1
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Left =4410
                    Top =1785
                    Width =6165
                    Height =300
                    ColumnOrder =6
                    FontSize =12
                    TabIndex =6
                    ForeColor =-2147483615
                    Name ="Text1"
                    ControlSource ="=\"Rent Bill\" & \" for Period <Strong>\" & [BillFrom] & \"</Strong> through <St"
                        "rong>\" & [BillTo]"
                    TextFormat =1

                    LayoutCachedLeft =4410
                    LayoutCachedTop =1785
                    LayoutCachedWidth =10575
                    LayoutCachedHeight =2085
                End
                Begin TextBox
                    Visible = NotDefault
                    TextAlign =2
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Left =10620
                    Top =810
                    Width =1065
                    ColumnOrder =7
                    FontSize =8
                    TabIndex =7
                    ForeColor =10855845
                    Name ="BillID"
                    ControlSource ="BillID"

                    LayoutCachedLeft =10620
                    LayoutCachedTop =810
                    LayoutCachedWidth =11685
                    LayoutCachedHeight =1050
                End
                Begin TextBox
                    TextAlign =1
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Left =4410
                    Top =1440
                    Width =6165
                    Height =360
                    ColumnOrder =4
                    FontSize =14
                    FontWeight =700
                    TabIndex =4
                    Name ="Text228"
                    ControlSource ="=UCase([MLNO]) & \" | Schedule \" & [schedule]"

                    LayoutCachedLeft =4410
                    LayoutCachedTop =1440
                    LayoutCachedWidth =10575
                    LayoutCachedHeight =1800
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    TextAlign =1
                    Left =4110
                    Width =720
                    Height =195
                    FontSize =8
                    FontWeight =700
                    BorderColor =-2147483615
                    Name ="Label5"
                    Caption =" Bill To:"
                    LayoutCachedLeft =4110
                    LayoutCachedWidth =4830
                    LayoutCachedHeight =195
                End
                Begin Label
                    TextAlign =1
                    Left =120
                    Width =2310
                    Height =675
                    FontSize =32
                    FontWeight =700
                    Name ="Label231"
                    Caption ="INVOICE"
                    LayoutCachedLeft =120
                    LayoutCachedWidth =2430
                    LayoutCachedHeight =675
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    TextAlign =1
                    Left =4110
                    Top =1290
                    Width =720
                    Height =195
                    FontSize =8
                    FontWeight =700
                    BorderColor =-2147483615
                    Name ="Label232"
                    Caption =" Lease:"
                    LayoutCachedLeft =4110
                    LayoutCachedTop =1290
                    LayoutCachedWidth =4830
                    LayoutCachedHeight =1485
                End
                Begin TextBox
                    Visible = NotDefault
                    TextAlign =2
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Left =10635
                    Top =540
                    Width =720
                    ColumnOrder =5
                    FontSize =8
                    TabIndex =5
                    ForeColor =10855845
                    Name ="clientID"
                    ControlSource ="clientID"

                    LayoutCachedLeft =10635
                    LayoutCachedTop =540
                    LayoutCachedWidth =11355
                    LayoutCachedHeight =780
                End
                Begin TextBox
                    OldBorderStyle =1
                    TextAlign =2
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =11670
                    Top =1620
                    Width =2535
                    Height =315
                    ColumnOrder =3
                    TabIndex =3
                    Name ="SchBnkClNO"
                    ControlSource ="AccountNumber"
                    StatusBarText ="Assignee Client No."
                    GroupTable =24
                    TopPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =11670
                    LayoutCachedTop =1620
                    LayoutCachedWidth =14205
                    LayoutCachedHeight =1935
                    RowStart =5
                    RowEnd =5
                    LayoutGroup =3
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =24
                End
                Begin TextBox
                    OldBorderStyle =1
                    TextAlign =2
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =11670
                    Top =390
                    Width =2535
                    Height =330
                    ColumnOrder =0
                    FontSize =14
                    FontWeight =700
                    Name ="Text410"
                    ControlSource ="BillNo"
                    GroupTable =24
                    TopPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =11670
                    LayoutCachedTop =390
                    LayoutCachedWidth =14205
                    LayoutCachedHeight =720
                    RowStart =1
                    RowEnd =1
                    LayoutGroup =3
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =24
                End
                Begin TextBox
                    OldBorderStyle =1
                    TextAlign =2
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Left =11670
                    Top =720
                    Width =2535
                    Height =315
                    ColumnOrder =1
                    FontSize =10
                    TabIndex =1
                    Name ="Text411"
                    ControlSource ="=\"Invoice Date: <strong>\" & [Billdate] & \"</strong>\""
                    GroupTable =24
                    TopPadding =0
                    BottomPadding =0
                    TextFormat =1

                    LayoutCachedLeft =11670
                    LayoutCachedTop =720
                    LayoutCachedWidth =14205
                    LayoutCachedHeight =1035
                    RowStart =2
                    RowEnd =2
                    LayoutGroup =3
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =24
                End
                Begin TextBox
                    OldBorderStyle =1
                    TextAlign =2
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Left =11670
                    Top =1035
                    Width =2535
                    Height =300
                    ColumnOrder =2
                    FontSize =10
                    TabIndex =2
                    Name ="Text412"
                    ControlSource ="=\"Due Date:<b> \" & [DueOn] & \"</b>\""
                    GroupTable =24
                    TopPadding =0
                    BottomPadding =0
                    TextFormat =1

                    LayoutCachedLeft =11670
                    LayoutCachedTop =1035
                    LayoutCachedWidth =14205
                    LayoutCachedHeight =1335
                    RowStart =3
                    RowEnd =3
                    LayoutGroup =3
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =24
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    TextAlign =2
                    Left =11670
                    Top =105
                    Width =2535
                    Height =285
                    FontWeight =700
                    BorderColor =-2147483615
                    ForeColor =16777215
                    Name ="Label414"
                    Caption ="  Invoice Number:"
                    GroupTable =24
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =11670
                    LayoutCachedTop =105
                    LayoutCachedWidth =14205
                    LayoutCachedHeight =390
                    LayoutGroup =3
                    BackThemeColorIndex =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =24
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    TextAlign =2
                    Left =11670
                    Top =1335
                    Width =2535
                    Height =285
                    FontWeight =700
                    Name ="Label428"
                    Caption ="  Account No.:"
                    GroupTable =24
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =11670
                    LayoutCachedTop =1335
                    LayoutCachedWidth =14205
                    LayoutCachedHeight =1620
                    RowStart =4
                    RowEnd =4
                    LayoutGroup =3
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =24
                End
                Begin TextBox
                    TextAlign =1
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =4125
                    Top =210
                    Width =6510
                    Height =1050
                    ColumnOrder =8
                    TabIndex =8
                    Name ="Text439"
                    ControlSource ="=[BillingAddress]"
                    TextFormat =1

                    LayoutCachedLeft =4125
                    LayoutCachedTop =210
                    LayoutCachedWidth =10635
                    LayoutCachedHeight =1260
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
                    Width =2895
                    Height =255
                    FontSize =10
                    FontWeight =700
                    BorderColor =-2147483615
                    ForeColor =16777215
                    Name ="group"
                    ControlSource ="=\" ▼ Group: \" & [UnitGroup]"

                    LayoutCachedTop =30
                    LayoutCachedWidth =2895
                    LayoutCachedHeight =285
                    BackThemeColorIndex =2
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
            Height =587
            BreakLevel =1
            Name ="GroupHeader0"
            Begin
                Begin Label
                    SpecialEffect =5
                    TextAlign =3
                    Left =12840
                    Top =300
                    Width =1365
                    Height =272
                    FontSize =9
                    FontWeight =700
                    Name ="Label96"
                    Caption ="Total Payment:"
                    GroupTable =1
                    LayoutCachedLeft =12840
                    LayoutCachedTop =300
                    LayoutCachedWidth =14205
                    LayoutCachedHeight =572
                    ColumnStart =9
                    ColumnEnd =9
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    SpecialEffect =5
                    TextAlign =2
                    Left =10650
                    Top =300
                    Width =1065
                    Height =272
                    FontSize =9
                    FontWeight =700
                    Name ="Label91"
                    Caption ="Tax:"
                    GroupTable =1
                    LayoutCachedLeft =10650
                    LayoutCachedTop =300
                    LayoutCachedWidth =11715
                    LayoutCachedHeight =572
                    ColumnStart =7
                    ColumnEnd =7
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    SpecialEffect =5
                    TextAlign =3
                    Left =9720
                    Top =300
                    Width =870
                    Height =272
                    FontSize =9
                    FontWeight =700
                    Name ="Label146"
                    Caption ="Tax Rate:"
                    GroupTable =1
                    LayoutCachedLeft =9720
                    LayoutCachedTop =300
                    LayoutCachedWidth =10590
                    LayoutCachedHeight =572
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    SpecialEffect =5
                    TextAlign =2
                    Left =7860
                    Top =300
                    Width =1109
                    Height =272
                    FontSize =9
                    FontWeight =700
                    Name ="Label69"
                    Caption ="Rent:"
                    GroupTable =1
                    LayoutCachedLeft =7860
                    LayoutCachedTop =300
                    LayoutCachedWidth =8969
                    LayoutCachedHeight =572
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    SpecialEffect =5
                    Left =7155
                    Top =300
                    Width =645
                    Height =272
                    FontSize =9
                    FontWeight =700
                    Name ="Label206"
                    Caption ="Status:"
                    GroupTable =1
                    LayoutCachedLeft =7155
                    LayoutCachedTop =300
                    LayoutCachedWidth =7800
                    LayoutCachedHeight =572
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    SpecialEffect =5
                    Left =3720
                    Top =300
                    Width =3375
                    Height =272
                    FontSize =9
                    FontWeight =700
                    Name ="Label13"
                    Caption ="Description:"
                    GroupTable =1
                    LayoutCachedLeft =3720
                    LayoutCachedTop =300
                    LayoutCachedWidth =7095
                    LayoutCachedHeight =572
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    SpecialEffect =5
                    Left =1560
                    Top =300
                    Width =2100
                    Height =272
                    FontSize =9
                    FontWeight =700
                    Name ="Label12"
                    Caption ="VIN:"
                    GroupTable =1
                    LayoutCachedLeft =1560
                    LayoutCachedTop =300
                    LayoutCachedWidth =3660
                    LayoutCachedHeight =572
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    SpecialEffect =5
                    TextAlign =2
                    Left =720
                    Top =300
                    Width =780
                    Height =272
                    FontSize =9
                    FontWeight =700
                    Name ="Label11"
                    Caption ="Unit:"
                    GroupTable =1
                    LayoutCachedLeft =720
                    LayoutCachedTop =300
                    LayoutCachedWidth =1500
                    LayoutCachedHeight =572
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    SpecialEffect =5
                    TextAlign =2
                    Left =9030
                    Top =300
                    Width =630
                    Height =272
                    FontSize =9
                    FontWeight =700
                    Name ="Label238"
                    Caption ="Taxed?"
                    GroupTable =1
                    LayoutCachedLeft =9030
                    LayoutCachedTop =300
                    LayoutCachedWidth =9660
                    LayoutCachedHeight =572
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    SpecialEffect =5
                    TextAlign =2
                    Left =11775
                    Top =300
                    Width =1005
                    Height =272
                    FontSize =9
                    FontWeight =700
                    Name ="FeeLabel"
                    Caption ="    "
                    GroupTable =1
                    LayoutCachedLeft =11775
                    LayoutCachedTop =300
                    LayoutCachedWidth =12780
                    LayoutCachedHeight =572
                    ColumnStart =8
                    ColumnEnd =8
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Line
                    Left =504
                    Top =270
                    Width =0
                    Height =285
                    BorderColor =-2147483615
                    Name ="Line245"
                    LayoutCachedLeft =504
                    LayoutCachedTop =270
                    LayoutCachedWidth =504
                    LayoutCachedHeight =555
                End
                Begin TextBox
                    FontUnderline = NotDefault
                    TextFontCharSet =186
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Left =285
                    Top =15
                    Width =5535
                    Height =255
                    ColumnWidth =2745
                    FontSize =10
                    FontWeight =700
                    Name ="LocationNamex"
                    ControlSource ="=\"Location: \" & [LocationName] & \" - \" & [LState]"

                    LayoutCachedLeft =285
                    LayoutCachedTop =15
                    LayoutCachedWidth =5820
                    LayoutCachedHeight =270
                    ForeThemeColorIndex =2
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            Height =322
            Name ="Detail"
            Begin
                Begin TextBox
                    TextFontCharSet =163
                    TextAlign =2
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =720
                    Top =30
                    Width =780
                    Height =270
                    FontSize =9
                    Name ="UnitUnitNum"
                    ControlSource ="UnitNum"
                    GroupTable =1

                    LayoutCachedLeft =720
                    LayoutCachedTop =30
                    LayoutCachedWidth =1500
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
                    Left =1560
                    Top =30
                    Width =2100
                    Height =270
                    ColumnWidth =2130
                    FontSize =9
                    TabIndex =1
                    Name ="UnitVIN"
                    ControlSource ="VIN"
                    GroupTable =1

                    LayoutCachedLeft =1560
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
                    ControlSource ="UnitDesc"
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
                    Width =1109
                    Height =270
                    FontSize =9
                    TabIndex =3
                    Name ="UnitsbillRent"
                    ControlSource ="BilledRent"
                    Format ="$#,##0.00;($#,##0.00)"
                    GroupTable =1

                    LayoutCachedLeft =7860
                    LayoutCachedTop =30
                    LayoutCachedWidth =8969
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
                    Left =10650
                    Top =30
                    Width =1065
                    Height =270
                    FontSize =9
                    TabIndex =4
                    Name ="Text139"
                    ControlSource ="UnitTaxes"
                    Format ="$#,##0.00;($#,##0.00)"
                    GroupTable =1

                    LayoutCachedLeft =10650
                    LayoutCachedTop =30
                    LayoutCachedWidth =11715
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
                    Left =9720
                    Top =30
                    Width =870
                    Height =270
                    ColumnWidth =2160
                    FontSize =9
                    TabIndex =7
                    Name ="taxrate"
                    ControlSource ="=IIf([BilledRent]>0,IIf([UnitTaxExempt]=-1,\"--  \",[TaxRateUnit]),\"--  \")"
                    Format ="Percent"
                    GroupTable =1

                    LayoutCachedLeft =9720
                    LayoutCachedTop =30
                    LayoutCachedWidth =10590
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
                    Left =12840
                    Top =30
                    Width =1365
                    Height =270
                    FontSize =9
                    TabIndex =8
                    Name ="LineTotal"
                    ControlSource ="UnitTotalRent"
                    Format ="$#,##0.00;($#,##0.00)"
                    GroupTable =1

                    LayoutCachedLeft =12840
                    LayoutCachedTop =30
                    LayoutCachedWidth =14205
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
                    Left =9030
                    Top =30
                    Width =630
                    Height =270
                    FontSize =9
                    TabIndex =6
                    Name ="Text236"
                    ControlSource ="=IIf([BilledRent]>0,IIf([UnitTaxExempt]=-1,\"No\",\"Yes\"),\"--\")"
                    GroupTable =1

                    LayoutCachedLeft =9030
                    LayoutCachedTop =30
                    LayoutCachedWidth =9660
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
                    Visible = NotDefault
                    TextFontCharSet =163
                    TextAlign =3
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =11775
                    Top =30
                    Width =1005
                    Height =270
                    FontSize =9
                    TabIndex =5
                    Name ="FSLFeeX1"
                    ControlSource ="=IIf([BilledFees]=0,\"--   \",CCur([BilledFees]))"
                    Format ="$#,##0.00;($#,##0.00)"
                    GroupTable =1

                    LayoutCachedLeft =11775
                    LayoutCachedTop =30
                    LayoutCachedWidth =12780
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
                    Width =0
                    Height =317
                    BorderColor =-2147483615
                    Name ="Line244"
                    LayoutCachedLeft =504
                    LayoutCachedWidth =504
                    LayoutCachedHeight =317
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
                    Width =12528
                    BorderColor =-2147483615
                    Name ="Line243"
                    LayoutCachedLeft =504
                    LayoutCachedTop =135
                    LayoutCachedWidth =13032
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
                    TabIndex =2
                    BorderColor =-2147483615
                    ForeColor =-2147483615
                    Name ="Text247"
                    ControlSource ="=Sum([BilledRent])"
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
                    Width =1064
                    FontSize =9
                    TabIndex =3
                    BorderColor =-2147483615
                    ForeColor =-2147483615
                    Name ="Text248"
                    ControlSource ="=Sum([UnitTaxes])"
                    Format ="$#,##0.00;($#,##0.00)"
                    ControlTipText ="TaxAmt Total"

                    LayoutCachedLeft =10649
                    LayoutCachedWidth =11713
                    LayoutCachedHeight =240
                End
                Begin TextBox
                    Visible = NotDefault
                    OldBorderStyle =1
                    TextFontCharSet =163
                    TextAlign =3
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =11775
                    Top =15
                    Width =1008
                    FontSize =9
                    TabIndex =4
                    BorderColor =-2147483615
                    ForeColor =-2147483615
                    Name ="FSLFeeX2"
                    ControlSource ="=Sum([BilledFees])"
                    Format ="$#,##0.00;($#,##0.00)"

                    LayoutCachedLeft =11775
                    LayoutCachedTop =15
                    LayoutCachedWidth =12783
                    LayoutCachedHeight =255
                End
                Begin TextBox
                    OldBorderStyle =1
                    TextFontCharSet =163
                    TextAlign =3
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =12840
                    Width =1365
                    FontSize =9
                    TabIndex =1
                    BorderColor =-2147483615
                    ForeColor =-2147483615
                    Name ="Text235"
                    ControlSource ="=Sum([UnitTotalRent])"
                    Format ="$#,##0.00;($#,##0.00)"
                    ControlTipText ="LineTot Total"

                    LayoutCachedLeft =12840
                    LayoutCachedWidth =14205
                    LayoutCachedHeight =240
                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextFontCharSet =186
                    TextAlign =2
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =9180
                    Width =1155
                    FontSize =7
                    FontWeight =700
                    TabIndex =5
                    BorderColor =-2147483615
                    ForeColor =8355711
                    Name ="Text215"
                    ControlSource ="=\" Tax Type: \" & [sttaxtype] & \"\""
                    TextFormat =1

                    LayoutCachedLeft =9180
                    LayoutCachedWidth =10335
                    LayoutCachedHeight =240
                End
                Begin TextBox
                    OldBorderStyle =1
                    TextFontCharSet =186
                    TextAlign =3
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =3660
                    Width =4050
                    FontSize =9
                    FontWeight =700
                    TabIndex =6
                    BorderColor =-2147483615
                    Name ="Text136"
                    ControlSource ="=Sum([countMe]) & IIf(Sum([countMe])>1,\" units at \",\" unit at \") & [Location"
                        "Name] & \" - Sub Total ►\""

                    LayoutCachedLeft =3660
                    LayoutCachedWidth =7710
                    LayoutCachedHeight =240
                    ForeThemeColorIndex =4
                    ForeShade =50.0
                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =300
            Name ="GroupFooter2"
            Begin
                Begin TextBox
                    OldBorderStyle =1
                    TextFontCharSet =163
                    TextAlign =3
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =12840
                    Top =60
                    Width =1365
                    FontSize =9
                    BorderColor =-2147483615
                    ForeColor =16777215
                    Name ="AccessTotalsLineTot1"
                    ControlSource ="=Sum([UnitTotalRent])"
                    Format ="$#,##0.00;($#,##0.00)"
                    ControlTipText ="LineTot Total"

                    LayoutCachedLeft =12840
                    LayoutCachedTop =60
                    LayoutCachedWidth =14205
                    LayoutCachedHeight =300
                    BackThemeColorIndex =2
                End
                Begin TextBox
                    OldBorderStyle =1
                    TextFontCharSet =163
                    TextAlign =3
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =10648
                    Top =60
                    Width =1065
                    FontSize =9
                    TabIndex =1
                    BorderColor =-2147483615
                    ForeColor =16777215
                    Name ="Text187"
                    ControlSource ="=Sum([UnitTaxes])"
                    Format ="$#,##0.00;($#,##0.00)"
                    ControlTipText ="LineTot Total"

                    LayoutCachedLeft =10648
                    LayoutCachedTop =60
                    LayoutCachedWidth =11713
                    LayoutCachedHeight =300
                    BackThemeColorIndex =2
                End
                Begin TextBox
                    OldBorderStyle =1
                    TextFontCharSet =163
                    TextAlign =3
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =7845
                    Top =60
                    Width =1110
                    FontSize =9
                    TabIndex =2
                    BorderColor =-2147483615
                    ForeColor =16777215
                    Name ="Text188"
                    ControlSource ="=Sum([BilledRent])"
                    Format ="$#,##0.00;($#,##0.00)"
                    ControlTipText ="LineTot Total"

                    LayoutCachedLeft =7845
                    LayoutCachedTop =60
                    LayoutCachedWidth =8955
                    LayoutCachedHeight =300
                    BackThemeColorIndex =2
                End
                Begin TextBox
                    Visible = NotDefault
                    OldBorderStyle =1
                    TextFontCharSet =163
                    TextAlign =3
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =11773
                    Top =60
                    Width =1008
                    FontSize =9
                    TabIndex =3
                    BorderColor =-2147483615
                    ForeColor =16777215
                    Name ="FSLFeeX3"
                    ControlSource ="=Sum([BilledFees])"
                    Format ="$#,##0.00;($#,##0.00)"

                    LayoutCachedLeft =11773
                    LayoutCachedTop =60
                    LayoutCachedWidth =12781
                    LayoutCachedHeight =300
                    BackThemeColorIndex =2
                End
                Begin TextBox
                    TextFontCharSet =163
                    TextAlign =3
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Left =3705
                    Top =60
                    Width =4050
                    FontSize =9
                    FontWeight =700
                    TabIndex =4
                    Name ="Text138"
                    ControlSource ="=Sum([countMe]) & IIf(Sum([countMe])>1,\" units\",\" unit\") & \" - Group \" & ["
                        "UnitGroup] & \" - Sub Total ►\""

                    LayoutCachedLeft =3705
                    LayoutCachedTop =60
                    LayoutCachedWidth =7755
                    LayoutCachedHeight =300
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
                    Width =2938
                    FontSize =9
                    Name ="Text128"
                    ControlSource ="=Date() & \"@ \" & Time()"

                    LayoutCachedWidth =2938
                    LayoutCachedHeight =240
                End
                Begin TextBox
                    TextFontCharSet =161
                    TextAlign =3
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Left =12150
                    Width =2938
                    FontSize =9
                    TabIndex =1
                    Name ="Text130"
                    ControlSource ="=\"Page \" & [Page] & \" of \" & [Pages]"

                    LayoutCachedLeft =12150
                    LayoutCachedWidth =15088
                    LayoutCachedHeight =240
                End
            End
        End
        Begin FormFooter
            KeepTogether = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =4305
            Name ="ReportFooter"
            Begin
                Begin Image
                    PictureAlignment =0
                    Left =180
                    Top =1620
                    Width =3528
                    Name ="Image314"
                    Picture ="FALogo.jpg"
                    GroupTable =25
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    ImageData = Begin
                        0xffd8ffe000104a46494600010101006000600000ffdb00430003020203020203 ,
                        0x03030304030304050805050404050a070706080c0a0c0c0b0a0b0b0d0e12100d ,
                        0x0e110e0b0b1016101113141515150c0f171816141812141514ffdb0043010304 ,
                        0x0405040509050509140d0b0d1414141414141414141414141414141414141414 ,
                        0x141414141414141414141414141414141414141414141414141414141414ffc0 ,
                        0x0011080036014703012200021101031101ffc4001f0000010501010101010100 ,
                        0x000000000000000102030405060708090a0bffc400b510000201030302040305 ,
                        0x0504040000017d01020300041105122131410613516107227114328191a10823 ,
                        0x42b1c11552d1f02433627282090a161718191a25262728292a3435363738393a ,
                        0x434445464748494a535455565758595a636465666768696a737475767778797a ,
                        0x838485868788898a92939495969798999aa2a3a4a5a6a7a8a9aab2b3b4b5b6b7 ,
                        0xb8b9bac2c3c4c5c6c7c8c9cad2d3d4d5d6d7d8d9dae1e2e3e4e5e6e7e8e9eaf1 ,
                        0xf2f3f4f5f6f7f8f9faffc4001f01000301010101010101010100000000000001 ,
                        0x02030405060708090a0bffc400b5110002010204040304070504040001027700 ,
                        0x0102031104052131061241510761711322328108144291a1b1c109233352f015 ,
                        0x6272d10a162434e125f11718191a262728292a35363738393a43444546474849 ,
                        0x4a535455565758595a636465666768696a737475767778797a82838485868788 ,
                        0x898a92939495969798999aa2a3a4a5a6a7a8a9aab2b3b4b5b6b7b8b9bac2c3c4 ,
                        0xc5c6c7c8c9cad2d3d4d5d6d7d8d9dae2e3e4e5e6e7e8e9eaf2f3f4f5f6f7f8f9 ,
                        0xfaffda000c03010002110311003f00f5dd6b52b8bfd62f6e6476124d33c8dcf7 ,
                        0x2c49aa7f6897fe7a37e75db7c5ef02dcf833c5f7aa62616373234d6d201f2942 ,
                        0x738fa8ce3f0f715c2d7f3f6269d4a35a74ea7c49ea7c1d48ca137196e49f6897 ,
                        0xfe7a37e747da25ff009e8df9d4745735d9049f6897fe7a37e747da24ff009e8d ,
                        0xf9d474e41b9d47a9c517607d290ea5770fecd33cb7b33c934b13428ee727634d ,
                        0xb40fcb3f957cefa76b9a86912892c6f6e2d1c1cee865643fa1afa9e3bbf0ef86 ,
                        0x7e10f8760f14a79b633430feef6336e90af98381e9c9ae57fe12af837ff40eff ,
                        0x00c9797fc6bef730c17b6f617c4460e308af79d9dfb9ee6228f3f25ea24d25bb ,
                        0xd4a1f09ff680bd6d4adf49f124bf688266091deb0c3c6c4f1bbd57dfa8f7edf4 ,
                        0x557827fc255f06bfe81dff0092f2ff008d76b67f1f7c157122423519601d034b ,
                        0x6ee147e38afa0cab191c3d374b178a84fb3e6d7d1dff0003bb0b59538f2d5aa9 ,
                        0xf6d4f46ae57e2b6adfd83f0c7c59a86ff2cdbe957322b7a3089b1fae2ba0d375 ,
                        0x4b3d66cd2eac2ea2bcb67fbb2c2e194fe22bca7f6b8d61345fd9e3c63239dbe7 ,
                        0x5ba5b0fac922ae3f5afac8b5249a7a1ea5eeae8fcbbf04f82f54f889e2eb0f0e ,
                        0xe8b1acda9df48520491c229214b1c93c0e01af69ff00860df8b5ff0040bb0ffc ,
                        0x1845fe35cdfec97e2af0e7827e35697ae78a7515d334db3866713bc6ee3cc319 ,
                        0x5518504ff11ed5fa07a6fed63f0a357d42d6c6cfc5f04f77752ac30c4b6b700b ,
                        0xbb1014731f72456adb11f0ff00c73f87fac7c1bf81be0af08ebd1c36faadc6ad ,
                        0x7da94d0c32ac836848638ce471d377e75eebff0004d9d1c43e0df18ea98e6e2f ,
                        0xa1b61ff6ce32dffb56bceffe0a39ac25d7c4df0ee9cad97b4d2f7b2fa6f91bfa ,
                        0x28af7bfd813494d3fe00c372176bdf6a371331f5da420ffd029741f53e90afca ,
                        0x7fdb27576d63f68cf169dfbd2de48ad93db642808ffbeb35fab15f8e1f19f595 ,
                        0xf107c5cf186a2a72971ab5cba9f6f35b1fa6294419e8ba07ec4ff147c49a1e9f ,
                        0xab5969966d677d6f1dcc25ef635628ea194904f070457a2fc19fd92fc7ff000b ,
                        0x7e2269de2ef1158d9c1a4e8b1dc5ecb247791bb06482428428393f3edaf67f0e ,
                        0xfedc9f09741f0fe99a625ceadb2ced62b75ff403d1102ff7bdabb5f88df1534b ,
                        0xf19fecbfe2cf1768724e96175a55d4703dc47e5bee3ba2e9fef74a77607e60e8 ,
                        0x1a06a7f103c5d69a4e9f189f55d52e4451233050d23b7727a726bdb3fe1837e2 ,
                        0xd7fd02ec3ff06117f8d79cfc01f1668de05f8bde1cf106bf2cd1697a7dc1b891 ,
                        0xade2f31f211b68033fdec57def07ede7f0aee668e28ee3562f230551f603d49c ,
                        0x7ad377e823c37c27fb24fc47f0afc21f1fe94fa55bcbad6b8d636f04515e4641 ,
                        0x86394c92927381caa77e727d2bc37e217ecc9f103e17f86a5d7bc45a4c765a6c ,
                        0x72244d20b98dcee638030ac4d7eb6ab6e507a679af94bfe0a2fac0b4f849a1e9 ,
                        0xf9c3de6acaf8f55489f3fab0a49b19f0dfc32f84be26f8bdabdce99e18b117d7 ,
                        0x76f01b8955a558c2a6e0b9cb103ab0afabbf64bfd973c73f0d7e305aebfe28d1 ,
                        0xe3b4b0b6b4984728b98e422565da385627a16aabff0004d6d203eb3e37d508e6 ,
                        0x3b7b6b653fef33b1ff00d016beeea1be801457cd9fb44fed597bf08fc75a7e83 ,
                        0xa2e9b6fa9436d1c773addc4caec2d6292455403690031073ce7ef2f15f47c132 ,
                        0x5cc31cd19dd1c8a1d5bd411906a46494515e2dfb377c56d7fe2b49e39b9d5fc8 ,
                        0x163a6eb5258e9e218b66235cf53fc470579a407b4d1451400514514005145711 ,
                        0xe0ff008a767e34f1e78c3c3567653a7fc2352430cf7aecbe5cd248a58aa81cfc ,
                        0xb820e7bd0076f4514500145145001451450067eb7a0e9de24b16b3d4ece2bdb6 ,
                        0x6e764ab9c1f507a83ee2bcdaeff669f09dc4c5e292fad94f3e5a4aac07d0b293 ,
                        0xf99af59a2b871181c2e29debd3527e68c6a51a75759c6e78d5dfecdfe11b0b69 ,
                        0x2e2e751be8208c6e79249235551ea4ecac7ff8541f0d3fe869ff00c9c83ff89a ,
                        0xf73d4f4db7d634eb9b1bb8c4b6d711b45221eea460d7c9bf10be0bf883c217d3 ,
                        0x3db4336a5a633131dc42a5881e8e0743fa57c9e6d83a3808aa9470919c7aefa7 ,
                        0xfc03cac5528504a50a49a3bfff008541f0d3fe869ffc9c83ff0089a747f07be1 ,
                        0xb33aaa78a0b3138005dc39cffdf35f3cb79d1b1562c18755279addf02d94ba97 ,
                        0x8c746b7c332c97710607382378cfe95f2b4f31c2d49a82c1c3576ea7991c4539 ,
                        0x34bd923d9bf69d962d37c3fe1bd262cec52e541feea2aa8fe75f3bd7b9fed50d ,
                        0x2b7883454287c95b562add8b173b87e185fcebc3bcb6fee9ac73f9736635176b ,
                        0x2fc1118e77c44bfae8368a7796dfdd34796dfdd35f3c701dbfc27f88579e07f1 ,
                        0x2dbb0998e9d3b84b9849f9594f1bbea3a8fa63bd76bff0502d696c3e02a5a839 ,
                        0x3a86a76f10c7a287933ff8e0af17b7b79269911109663802baaff828bea4f6fe ,
                        0x0df016905b0cf2cd70ebdf291a20ff00d0dabf4de11ad5270ab464ef18d9af2b ,
                        0xdffc8fa2cae7271941ec8f9eff00665fd9e17f682d6b5ab397567d1e1d3add26 ,
                        0x332c1e6ee666202e370c7009fc2bea3f87dfb00d9781bc6fa1f885bc5b25f8d3 ,
                        0x2f22bbfb335905121460c149de71c8f4ac8ff826be9062f0df8d7532b8135d5b ,
                        0xdb06ff00711d8ffe8c15f6757e80db3dd47e5efedd1acaeadfb44eb30af22c6d ,
                        0xedad7f1f295cff00e875f74feca7a58d23f67bf0545e5f96d2597da1863a991d ,
                        0x9f3f930afcdffda43563ae7c78f1c5e03b87f6acd129ff006636f2c7e882bd03 ,
                        0xc31fb747c47f0ae8ba56916b0e8ad61a75bc76b146f66d931a2850090fd703ad ,
                        0x55b411fa5bacdf0d2f47bebd3d2de09263ff0001527fa57e38f84b463f107e25 ,
                        0x695a6c81d46ada9c70c9e5fde0b2480311ee0135f777867f6b087e307c03f889 ,
                        0x752d8ae91e24d23489da5b78d8b4522ba32a3a13c8e4e0839c6472735f03f81b ,
                        0xc617be00f18697e23d3e3865bdd3a75b8896e14b2161d320114a207dfdff000e ,
                        0xedf875ff00415d7bfeff00c3ff00c6ead7ed49a169ff0008ff00645b9f0be98f ,
                        0x2b59a35b58c2d3105db330909620019f95bb5785e8dff0516f1cdadea3ea5a1e ,
                        0x8f7d6b9f9e2892489b1df0db8e0fe06bb1fdb0be3269bf14bf671f086ada32c9 ,
                        0x1daeafa996920971be178a370e871d70cdd7b8c1ef46b7d46783fec9ff0004f4 ,
                        0xaf8e7f102ff48d6a7bbb7d3ed74f7ba2f66caafbc3a2a8cb0231f31eddabebbd ,
                        0x27fe09fbf0f349d52cef9351d6e57b6992611c9344558ab038388fa1c57c45f0 ,
                        0x53e3b6bff02752d4af740b6b19e7bf896190dec45f6aa9cfcb8618c9fe55f457 ,
                        0xc38ff8289eacdae5adb78cf44b36d2e57092de69c19248413f7f6962180ee383 ,
                        0xfca9bb88fbbebe1bff008294eb599bc11a481caadcdd1fc4c6a3f91afb82dee2 ,
                        0x3bab78a785d6586450e8ea721948c823db15f9d5ff000512d65efbe31697a775 ,
                        0x8ec7498c8ff79e4727f402a63b8d9ec7ff0004e3d25ad7e18f88efd936fdab54 ,
                        0x08ad8fbca912ff005735f4df8dbc5d63e03f096ade21d49fcbb2d3addae24f56 ,
                        0xc0e147b93803dc8af1cfd86f477d27f677d19e4183797371723e9e6141ff00a0 ,
                        0x573ffb4f6a171f163e21f84be0c69533a477d2aea5ae4d09e61b64e554fbe016 ,
                        0xe7bf97eb46ec0e5fc33f0af50f1c7ecd7f12bc65ae5bf9be27f1a4526ac8add6 ,
                        0x38613e65ba0f6f9723d8afa57a6f817e39d97877f64fd23c73a8b8b97b0d356d ,
                        0x9e2de374d7319f25533eacca0fb039a55fd8e7c2696c2d97c4be325b709e5885 ,
                        0x75b6081718db8db8c638c57cab751dfe9bfb3afc40f0461a7ff843fc5b1cb710 ,
                        0xaa924dab174dc7dbcc4ce7b669ee07d01e1ff03fc59f15f82e4f883ad7c4dbef ,
                        0x0e6af35b36a36ba2d9c0a6c6de2d9bd1248c9c36475ee33d4915e43e0fd6fc51 ,
                        0xf0fbf654d0fc77a46a77562c3c5725dea31db1c0b9b7661110e3b8df181ff02a ,
                        0xf5cf8cdfb48e91e2cf83dab697f0ea66d6352b8d21e6bc9618d923d2ed047fbc ,
                        0x69588f95c8ca2a8e771f6ad1f03fc3b4f117ec3969e1f440d2de68525dc5b973 ,
                        0xfbd6669d0ffdf58a00ebff0068df8b175e0bf8636ade1c767f12f89248ac3465 ,
                        0x8f05fcc971f38fa29e0fa95ae265f1878ef5af1069ff0008bc1daf6ed6348b18 ,
                        0xe4f13f8c6f544ef0bb0c98e253c1724f19e9d38c135c8feca91ea3f1c7c4be1f ,
                        0xf15eb703ae8fe05d261d234e8e4e44b79b30f37d426dfc4a7a53be0bfc54f0ff ,
                        0x00c19f8a1f196cfc757cda5eab71ab1bc85a489ddae62dd21554c03938752077 ,
                        0xdded480daf09c3e2df02fed5da2782dfc7faf789f4a97499352bf8f55943ae71 ,
                        0x22a851fc237043c7ad637c25f1b7c4ef8e1e22f1de85a4f88ae345d0e0d6e779 ,
                        0x35f7512cb6f6fb888ad6d94f0090092dd801dcf31fc1dd7f59f137ed39f11bc6 ,
                        0xbaf5849a5cda7e81e65be9f71feb2daddc46f0ab8ecc514311d8b1aefbf611b5 ,
                        0xdbf0464d41a3d92ea3aadcdc336396c155e7f1534c0c4f827ad78de4f1dfc53f ,
                        0x8637be36bcbf7d223492c3c457312cd716c5b19c87c86e1870c48054e2b90fd9 ,
                        0x4fc03e2ef162f8a7c590fc47bed32093c40e2f552d227fed0f2f0eceec7eee43 ,
                        0x91c70335d57ecfb7526ade39fda0fc4d1446494df3dac21072c63137ca3f25fc ,
                        0xeb9ffd95fc79e1ff00f867dbaf04596aaade38d523d51e3d3d51bcd32796e549 ,
                        0x38c0caaae0e6803b8d37c57e35fda67c41ab7fc229e209bc13f0f34cb86b45d5 ,
                        0x6ce20d7ba9caa3e631b1fb88320e47a8eb92071df0f7e297897e1b5b7c70bdbe ,
                        0xf126a5e31d3bc22d15ae9e357937b34ed23a02cdd71b82e71da97f670fda03c3 ,
                        0x7e09f80ba7f876ca19af7c750cf3dac1e1d8606f3ee6e5e46643d3017046e627 ,
                        0xe50a73d2b2ff006729bc376bf04be255ff00c50d42ded6c7c41ae4f6b792c8e4 ,
                        0x19dc22b1f2f1c960eccc3038c6680353c6da47c4af09fc131f152e7e2aead1f8 ,
                        0x88c105f3696a231a70599d76c4b111b490ae3923923a77ad9d4fe257c40d3f4b ,
                        0xf067816c3c52ba978efc576cba9dceb57f650c09a3d9b4618e2340016187c161 ,
                        0xdbdc63c8bc65a9786f4d8fc01612f8b7c61e23f01d96af01bab2d674d92ded61 ,
                        0xb51f77e631af998e0609e06702bbaf8eda0f846c7f6a2d175af88965f68f026a ,
                        0xfa3adb5ade6f91608e65ce0318c838c76e987cf40681166ebc57e28f863f16bc ,
                        0x3da2f847e236a1f156f6fe2b93a9e8f7aeb3c706c8f72b8642447cff000e7b60 ,
                        0xf5a2bd3be18fc46f859a7f8be0f0a7c32d023be69a2696f750d0ec80b7b75552 ,
                        0x57ce98e0b162368e4f38a290cf76a28a2a461451450054bad22c6f7fe3e2cade ,
                        0x7ffaeb12b7f31505bf86b48b59926834ab18664395923b745653ea081c569515 ,
                        0x9ba706eee2ae4f2adec57bcd3ed751555bbb586e954e42cd18703f31553fe117 ,
                        0xd1bfe81363ff0080c9fe15a7450e9c24eed20e54f74667fc22fa37fd026c7ff0 ,
                        0x193fc28ff845f46ffa04d8ff00e0327f8569d14bd953fe55f7072c7b19abe1ad ,
                        0x22360cba5592b29c822dd011fa52eade1cd275e68cea7a5d96a263c843776e92 ,
                        0xedcf5c6e0715a3455c6318fc2ac3492d8a5a5e8ba76870b43a6d85ae9f0b36f6 ,
                        0x8ed615894b63192140e7815768a2a8660cde02f0cdc4cf2cbe1cd26596462cf2 ,
                        0x3d8c45989ea492bc9a67fc2bcf0aff00d0b3a3ff00e0045ffc4d7434500635bf ,
                        0x82fc3d671dc470683a6411dc27973247671a891739dac02f233d8d57ff008577 ,
                        0xe15ffa16747ffc008bff0089ae86b26dbc55a5dd5eeab691dd2f9da5e0dd0604 ,
                        0x08c15dd9cf718eb8e95129c6365276b89b4b72a7fc2bcf0aff00d0b3a3ff00e0 ,
                        0x045ffc4d4efe0bf0f49671da3e83a635a46e64481ace3288c7ab05db804fad56 ,
                        0x8fc79a5c960d77b6f163de891abda48af397cecf2c11f36707a52cde3ad3e26b ,
                        0x64582fe69a788cc218ace46915436d2597191cf1cd63f59a36bf3a23da43b8bf ,
                        0xf0af3c2bff0042ce8fff0080117ff1347fc2bbf0affd0b3a3ffe0045ff00c4d4 ,
                        0x527c40d3e1d5934c6b5d4bed8e32a9f629395c805b38fba091935ababebf63a1 ,
                        0x3590be9c41f6c9d6da12c38323024027b67154abd2926d4969bf90d4e2eeefb1 ,
                        0x7a289218d238d1638d005545180a074007a565ea5e11d0b59ba373a868ba75f5 ,
                        0xc1014cd73691c8f81d064826a79b5db2b7d6adb4979bfd3ee2269922009f9148 ,
                        0x0493dbafe86a947e36d1e4b1d5eed2eb741a53bc776c14fc85464fd47b8f434d ,
                        0xd6a7176725ff000caeff000d439e2b766b58d8db6996b1dad9dbc5696d18c243 ,
                        0x0204451d7014702a28f47b0875297508ec6dd2fe550925d2c4a25751d017c648 ,
                        0xe0719ed5147e20b09e4d3562b85946a08d25b347cac8a143139fa1150e9fe2cd ,
                        0x2b525d51a0bb429a64ad15d337ca232a324e4f6c77f634fdb53bdb997f4aff00 ,
                        0x96a3e68f735ea82e81a5acd792ae9b6825bc1b6e5c40bba71e8e71f30fad6659 ,
                        0xf8ef4bbe8e59516ee381216b859a6b4911248d464b2b1183c537fe13fd2e3d38 ,
                        0xdfdcaddd8d9ef8d04d756cf1a9f30e148c8e47a9ed59fd6a8dafcebef27da43b ,
                        0x97b4df08e85a3d84f6361a2e9f6365719f3adadad638e3933d7728001fc6b46d ,
                        0xad61b3b78edede18e08235089146a1555474000e00aa5a87886c34bbfb0b2b99 ,
                        0xc25cdf33ac118192db54b31f6000eb5976ff0010b4bb8d366d4765e45611c5e7 ,
                        0x9ba96d5d632990320e39ebdaaa588a516d4a4aebfe1ff2686ea4568d9b9a7697 ,
                        0x65a3db982c2cedeca0dc5fcab789635dc7a9c00064d54bdf0ae8ba96a906a777 ,
                        0xa3d85d6a307115e4d6a8f347feeb9191f81a4bef1358d8ac3cc9726785ae225b ,
                        0x58ccad2202b92a17afdf5a8bc3be2eb2f1479a6ca2bc548f20c93db3c6a483b4 ,
                        0x8048e4820e453f6d4f9f93995c39e37e5bea5ffec7b0fb4dcdc7d8adfed17482 ,
                        0x39e5f2977caa0602b9c65863b1a7e9fa6da6936896b636b0d95b27dd86de311a ,
                        0x2e79e140c0ac8b9f1c69567ab4ba7ccf347245224324c606f251dc028a5f1804 ,
                        0xe475f5a7daf8c6cafb509ad6186f2430caf04930b67f295d33b817c638c54fd6 ,
                        0x295f97995f617b485ed73474fd1ec349f3fec3656f67e7b9925fb3c4b1f98e7a ,
                        0xb360727dcd54d33c23a168b7f717da768ba7585edc1266b9b5b48e392427a966 ,
                        0x5009fc6a3d37c65a56ad7969696f739b8bab417d0c6ca417889c647e3daa48fc ,
                        0x55a64926ae8b7209d2b1f6b6c1c47f296ebdf807a552af49aba92d7cfcaff96a ,
                        0x3e78f463acfc2ba2e9fab4faa5ae8f616da9cffeb6f21b5449a4cf5dce064fe2 ,
                        0x6bc5fe3e7c19d6f58f11f82bc55e0bd334bbf9bc37752dcc9a05d6d821ba3215 ,
                        0x25c1fbbbc15ea7be0f6c1f626f19e90be1fb6d6bed59b1ba0be432a317919ba2 ,
                        0xaa8192d9ed8ec69b0f8cb4e9a18ddc5d5b33dcc76823b9b678dfcc7fba304743 ,
                        0xebd297d6295d7bcbbefd05cf0ee782f8e342f8affb456930f85359f09d9fc3ef ,
                        0x0bcd3c526a57536a097b712a23870912a80072a393e839ea0fd073f8674abcd1 ,
                        0x61d22f74fb7d434e86348d6def22595085000c86041e95043e33d22e34bb9d42 ,
                        0x2ba125b5b4e6da62aa72920709b48fa91f9e6996de34b0bcbfb8b5862bc93ecf ,
                        0x2bc12ce2d9fc9574fbc0be31c51f58a5a7bcb5d839e3dcbfa3787f4bf0edafd9 ,
                        0xb49d36d34bb6ce7c9b381614cfae14014566693e3dd27569ada346b8b7374a5e ,
                        0xddaeadde25980193b188c1e39c67a51550ad4ea2bc2499519464ae99d1514515 ,
                        0xa941451450014514500145145001451450014514500145145001451450053d67 ,
                        0x511a46917b7cc86416d0bcdb477daa4e3f4af308f4dbcf09dae8babdecf06a51 ,
                        0xdfac96f776eb0f9665f3d5a53b9f27761863a0e09a28af171ff1737f2abaff00 ,
                        0xc097dff338ebef7edfe662d8f8b7497f0bdf5f35bea32f8661b8b7860b492e49 ,
                        0xb8b49b04ee89f390a32a002debf4a8f50f19c7e1fb8f0eea3ad5c6a773f6bb09 ,
                        0x3c992ce60929433064129f97710b807df34515f31eda7c97becbff006efeaddb ,
                        0xa1e673cad7feb73d1bed4b79f12345b8405526d1a69006ea0192235278ff004d ,
                        0x83589f40b1b95df04f78c8c3a1ff00532608f707047d28a2bea6ca74ea297592 ,
                        0xff00db4f5378cafdd7e872bad59df782eead755beb95d475b7b7ba06e15768c2 ,
                        0xa2244b8ec06771f72c6a8f88f41baf0369325bcb7715e457fa4cd66e2283cb3e ,
                        0x646a64591be63b89cb8278ea28a2bc8ad4e318d56becf2db57f6ac9fadfcce49 ,
                        0x452e6f2b7e3b96bc431cbe03d62c6ead195b4dfb25d5d5b5ae3982631a8603fd ,
                        0x82595b1d8eea2fbc177da32da690da84371fdb163269e5960318f3501995dfe6 ,
                        0x3bb277827ae1a8a2aa54a0e759748b8dbfedeb27f7ad3cba0f955e6bb35f8dae ,
                        0x41ff000b21fc6925fd95989acd2db4cbb17d6d26c311611b2a98c81bbef7ae38 ,
                        0x038aa9a1f8ab4af1978662d3618f50992def6c63b85d46512060cdb485e4e07c ,
                        0xa7f3a28ae1a75ea5692751dee9dfee46319ca4fde7bdff0023a593c2b71a1ea5 ,
                        0xa44f7d782fe7177f65b5931831db2c329553eac4fde3df02b88f0af8d34ed53c ,
                        0x157fa1c23516bb86c51a5fb54a1e0f96440422e781f37a74a28ae8c57ee6aa84 ,
                        0x344d4bf28f72ea7b92b47cff00247a0e9ba2cbe1ef88d6565148ada49b3b99ad ,
                        0x2139dd6e59e3df18ff00633c8f4c91d315aff0dffe45b7ff00afdbaffd1ef451 ,
                        0x5ece1a2a388715b272ff00db0eca6ad36979fe871be20b630eb5e25d5a496496 ,
                        0xc74fd42d65b9d3cb6239d0449ce3fbcadb5864e0edc1aa5e15f1b69d6be38d57 ,
                        0x436fed26bfb9d42ef6fef47d98062cc094cf5c0f4eb4515e3d493a7898f2f597 ,
                        0xe7297f91c926e3515bbfeac9fc37a24fad5be8ed63325aea96ba1dacd697122e ,
                        0x551bcc9010c3bab0241144fa2b596ad79e17b59f649a8cd68b3dd489bb7ed8e4 ,
                        0x9656619192e5718f4268a2aa34e2e8d395b56d2f934d3fc3aee3515c917dc7df ,
                        0x0b9f09dedecb3bc57c9a15d2eadb123f295a29d1d1d517270cac198738e4f229 ,
                        0x96fe2e7f1a2dbdfc524e74dfedeb14b78ae8209236dcc5c7ca3a72b8e49e28a2 ,
                        0xb39ce54eb7b18bf76ef4f4724bf04bee25b719722db5fd486f3447b7f08ff6bd ,
                        0x948b019af8c17f091f2ce82f0ec6ff007d4e067d091e951f877c65a7d9f8db5d ,
                        0xd05c6a26fee351bd0bb65ff4519dc72573d783da8a2a64fd956a5c9a5d43f1bd ,
                        0xc4fdd9c6de5fa95fc3dadc1f15bc3ba47852cd65b13630c725cddcb8c8d89b40 ,
                        0x8c0ce7248e491c668a28ae9c0e16963a97b4c44799ad3b6892ed634a14e35a1c ,
                        0xd515dffc03ffd9
                    End

                    LayoutCachedLeft =180
                    LayoutCachedTop =1620
                    LayoutCachedWidth =3708
                    LayoutCachedHeight =3060
                    TabIndex =6
                    LayoutGroup =4
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =25
                End
                Begin Line
                    BorderWidth =1
                    Left =3330
                    Top =1560
                    Width =0
                    Height =720
                    BorderColor =8210719
                    Name ="Line355"
                    LayoutCachedLeft =3330
                    LayoutCachedTop =1560
                    LayoutCachedWidth =3330
                    LayoutCachedHeight =2280
                End
                Begin Line
                    BorderWidth =1
                    Left =4125
                    Top =1530
                    Width =4320
                    BorderColor =8210719
                    Name ="Line354"
                    LayoutCachedLeft =4125
                    LayoutCachedTop =1530
                    LayoutCachedWidth =8445
                    LayoutCachedHeight =1530
                End
                Begin Line
                    OldBorderStyle =2
                    BorderWidth =3
                    BorderLineStyle =1
                    Top =1350
                    Width =14688
                    Name ="Line147"
                    LayoutCachedTop =1350
                    LayoutCachedWidth =14688
                    LayoutCachedHeight =1350
                End
                Begin Subform
                    Locked = NotDefault
                    OldBorderStyle =0
                    Top =60
                    Width =14580
                    Height =1232
                    Name ="InvoiceSummary"
                    SourceObject ="Report.RentInvoiceSummary"
                    LinkChildFields ="BillID"
                    LinkMasterFields ="BillID"

                    LayoutCachedTop =60
                    LayoutCachedWidth =14580
                    LayoutCachedHeight =1292
                End
                Begin TextBox
                    OldBorderStyle =1
                    TextAlign =2
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =11670
                    Top =1920
                    Width =2535
                    Height =330
                    FontSize =14
                    FontWeight =700
                    TabIndex =2
                    Name ="Text304"
                    ControlSource ="BillNo"
                    GroupTable =15
                    TopPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =11670
                    LayoutCachedTop =1920
                    LayoutCachedWidth =14205
                    LayoutCachedHeight =2250
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =15
                End
                Begin TextBox
                    OldBorderStyle =1
                    TextAlign =2
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Left =11670
                    Top =2250
                    Width =2535
                    Height =315
                    FontSize =10
                    TabIndex =3
                    Name ="Text308"
                    ControlSource ="=\"Invoice Date: <strong>\" & [Billdate] & \"</strong>\""
                    GroupTable =15
                    TopPadding =0
                    BottomPadding =0
                    TextFormat =1

                    LayoutCachedLeft =11670
                    LayoutCachedTop =2250
                    LayoutCachedWidth =14205
                    LayoutCachedHeight =2565
                    RowStart =2
                    RowEnd =2
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =15
                End
                Begin TextBox
                    OldBorderStyle =1
                    TextAlign =2
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Left =11670
                    Top =2565
                    Width =2535
                    Height =300
                    FontSize =10
                    TabIndex =4
                    Name ="Text309"
                    ControlSource ="=\"Due Date:<b> \" & [BillDue] & \"</b>\""
                    GroupTable =15
                    TopPadding =0
                    BottomPadding =0
                    TextFormat =1

                    LayoutCachedLeft =11670
                    LayoutCachedTop =2565
                    LayoutCachedWidth =14205
                    LayoutCachedHeight =2865
                    RowStart =3
                    RowEnd =3
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =15
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    TextAlign =1
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Left =4125
                    Top =1635
                    Width =7515
                    Height =1950
                    FontSize =12
                    TabIndex =1
                    Name ="RemitToz"
                    ControlSource ="=[RemitTo] & \"<Br /><Br />Rent Bill for Period<b> \" & [BillFrom] & \"</b> thro"
                        "ugh <b>\" & [BillTo] & \"</b><br /><em>Please Reference:</em> Invoice : <b> \" &"
                        " [BillNo] & \"</b> Account No.: <b> \" & [AccountNumber] & \"</b>\""
                    StatusBarText ="default company name"
                    GroupTable =15
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    TextFormat =1

                    LayoutCachedLeft =4125
                    LayoutCachedTop =1635
                    LayoutCachedWidth =11640
                    LayoutCachedHeight =3585
                    RowEnd =5
                    LayoutGroup =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =15
                End
                Begin Label
                    TextAlign =1
                    Left =180
                    Top =3060
                    Width =3528
                    Height =270
                    FontSize =10
                    FontWeight =700
                    ForeColor =7500402
                    Name ="Label263"
                    Caption =" Telephone: 954-615-4400"
                    GroupTable =25
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =180
                    LayoutCachedTop =3060
                    LayoutCachedWidth =3708
                    LayoutCachedHeight =3330
                    RowStart =1
                    RowEnd =1
                    LayoutGroup =4
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =25
                End
                Begin TextBox
                    OldBorderStyle =1
                    TextAlign =3
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Left =11670
                    Top =3150
                    Width =2535
                    Height =435
                    FontSize =16
                    FontWeight =700
                    TabIndex =5
                    Name ="Text221"
                    ControlSource ="=Sum([UnitTotalRent])"
                    Format ="$#,##0.00;($#,##0.00)"
                    ControlTipText ="LineTot Total"
                    GroupTable =15
                    TopPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =11670
                    LayoutCachedTop =3150
                    LayoutCachedWidth =14205
                    LayoutCachedHeight =3585
                    RowStart =5
                    RowEnd =5
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =15
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    TextAlign =1
                    Left =4125
                    Top =1440
                    Width =765
                    Height =210
                    FontSize =8
                    FontWeight =700
                    BorderColor =4210752
                    Name ="Label150"
                    Caption =" Remit To:"
                    LayoutCachedLeft =4125
                    LayoutCachedTop =1440
                    LayoutCachedWidth =4890
                    LayoutCachedHeight =1650
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    TextAlign =1
                    Left =11670
                    Top =1635
                    Width =2535
                    Height =285
                    FontWeight =700
                    BorderColor =-2147483615
                    ForeColor =16777215
                    Name ="Label376"
                    Caption ="  Invoice Number:"
                    GroupTable =15
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =11670
                    LayoutCachedTop =1635
                    LayoutCachedWidth =14205
                    LayoutCachedHeight =1920
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =2
                    BackThemeColorIndex =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =15
                End
                Begin Label
                    OldBorderStyle =1
                    TextAlign =1
                    Left =11670
                    Top =2865
                    Width =2535
                    Height =285
                    FontWeight =700
                    Name ="Label387"
                    Caption ="  Total Payment Amount:"
                    GroupTable =15
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =11670
                    LayoutCachedTop =2865
                    LayoutCachedWidth =14205
                    LayoutCachedHeight =3150
                    RowStart =4
                    RowEnd =4
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =15
                End
            End
        End
    End
End
CodeBehindForm
' See "RentInvoice.cls"
