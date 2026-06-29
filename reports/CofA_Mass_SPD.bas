Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    AutoResize = NotDefault
    PopUp = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    TabularCharSet =238
    DateGrouping =1
    GrpKeepTogether =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =11364
    DatasheetFontHeight =11
    ItemSuffix =896
    Left =5400
    Top =2730
    ShortcutMenuBar ="PrintMe"
    RecSrcDt = Begin
        0xa2efe4cfefd4e440
    End
    RecordSource ="CofA_Report_Mass"
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
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
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
        Begin EmptyCell
            Height =240
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin BreakLevel
            GroupHeader = NotDefault
            ControlSource ="MLNo"
        End
        Begin BreakLevel
            GroupHeader = NotDefault
            ControlSource ="SchNo"
        End
        Begin BreakLevel
            GroupHeader = NotDefault
            GroupFooter = NotDefault
            KeepTogether =1
            ControlSource ="grpcd"
        End
        Begin BreakLevel
            ControlSource ="vin4"
        End
        Begin BreakLevel
            ControlSource ="unitunitnum"
        End
        Begin BreakLevel
            GroupHeader = NotDefault
            ControlSource ="UnitID"
        End
        Begin FormHeader
            KeepTogether = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =0
            Name ="ReportHeader"
        End
        Begin PageHeader
            Height =924
            Name ="PageHeaderSection"
            Begin
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1860
                    Width =9375
                    Height =315
                    Name ="Header"
                    ControlSource ="=IIf([Page]<>\"1\",\"<B>Exhibit A  </B>\",\"\") & \"<B>Certificate Of Acceptance"
                        "</B><Br />MASTER EQUIPMENT LEASE AGREEMENT NO:  <B>\" & [MLNo] & \"</B><Br />SCH"
                        "EDULE  NO:  <B>\" & [SchNo] & \"</B>  GROUP:  <B>\" & [SGrpNO]"
                    TextFormat =1

                    LayoutCachedLeft =1860
                    LayoutCachedWidth =11235
                    LayoutCachedHeight =315
                End
                Begin Image
                    Width =2160
                    Height =864
                    Name ="Auto_Logo0"
                    Picture ="FALogo_Small.jpg"
                    ImageData = Begin
                        0xffd8ffe000104a46494600010101006000600000ffdb00430003020203020203 ,
                        0x03030304030304050805050404050a070706080c0a0c0c0b0a0b0b0d0e12100d ,
                        0x0e110e0b0b1016101113141515150c0f171816141812141514ffdb0043010304 ,
                        0x0405040509050509140d0b0d1414141414141414141414141414141414141414 ,
                        0x141414141414141414141414141414141414141414141414141414141414ffc0 ,
                        0x0011080060012703012200021101031101ffc4001f0000010501010101010100 ,
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
                        0xfaffda000c03010002110311003f00fd53a28a2800a28a2800a28a2800a28a28 ,
                        0x00a28a2800a28a2800a28a2800a28a2803c03c7dfb456afe19f186a9a569f616 ,
                        0x135ada4be50927572e5801bb38603ef67b573fff000d4de24ffa05e95ff7c4bf ,
                        0xfc5d798f8d266b8f18ebb2b9cbbdfcec7f191ab1abf19c46758ff6d3e5aad2bb ,
                        0xfccf8fa98cafceed23d9ff00e1a9bc49ff0040bd2bfef897ff008ba3fe1a9bc4 ,
                        0x9ff40bd2bfef897ff8baf18a2b0fedacc3fe7f323eb988fe73d9ff00e1a9bc49 ,
                        0xff0040bd2bfef897ff008ba3fe1a9bc49ff40bd2bfef897ff8baf18a28fedacc ,
                        0x3fe7f30fae623f9cf67ff86a6f127fd02f4aff00be25ff00e2e8ff0086a6f127 ,
                        0xfd02f4affbe25ffe2ebc628a3fb6b30ff9fcc3eb988fe73d9ffe1a9bc49ff40b ,
                        0xd2bfef897ff8ba3fe1a9bc49ff0040bd2bfef897ff008baf18a28fedacc3fe7f ,
                        0x30fae623f9cf67ff0086a6f127fd02f4affbe25ffe2ebd33e0ff00c693f11af2 ,
                        0xeb4ebeb28ecb50862f3d1a16263910100f07904123b9ce7dabe4caf61fd97ed1 ,
                        0xa6f1f5e4ff00c30d83e7ea5d001fcff2af5b29cdf1d571b4e9cea3926ecd33ab ,
                        0x0b8baf2ad18ca574cf78f8a5e329fc07e0cbbd5ed63866ba8de348e39c1284b3 ,
                        0x8073820f4cf7af1087f6a7f10ab7efb49d31d7d1048a7f5735dc7ed457de4781 ,
                        0xf4fb607e69ef949ff7551cff00322be5daeecff34c561b1bece85471492fbcdf ,
                        0x1d8aab4eb72c256b23e91d0ff6a8d3ae1d5356d16e2cc1e0cb6b28947d704290 ,
                        0x3f3af5bf0cf8cb46f195a1b8d1f508af517efaa921d33fde53823f115f09d5fd ,
                        0x0f5dbff0dea515fe9975259ddc472b24671f811dc1ee0f06b8f07c4f8aa524b1 ,
                        0x2b9e3f73ff002feb731a399558bfde6abf13ef7a2b81f847f14a0f891a2b9955 ,
                        0x6df57b5005cc0bd0e7a3affb27d3b1e3d09efabf4dc3e229e2a946b52778b3e9 ,
                        0x29d48d48a9c5e8c28a28ae8340a28a2800a28a2800a28a2800a28a2800a28a28 ,
                        0x00a28a2800a28a2800a28a2800a28a2803e25f8ada3be87f11bc416acbb41bb7 ,
                        0x993fdd93e75fd1857275f507ed09f0b67f1459c7afe95099b51b38f64f020f9a ,
                        0x68864e40eecb93c7707d803f2fd7e239c60a782c5ce2d7badb6bd1ff0096c7c5 ,
                        0xe2e8ba355a7b3d828a28af14e30a28a2800a28a2800a28a2800afa03f651b1cc ,
                        0xbe23bd23a2c30a9faef27f92d7cff5f517ecbb63e4781efee48e67be603fdd54 ,
                        0x41fcc9afa6e1da7cf98c1f64dfe16fd4f4b2f8f3621795cddf8c5f0bef7e2645 ,
                        0xa5456b7f0d9259b48cfe7216dc5b6818c7a60fe75e67ff000ca9aaff00d076cf ,
                        0xfefd3d77df123e3c43f0f7c487483a3b6a0cb0a4ad2adc88f1bb3c6369ec07e7 ,
                        0x5cb7fc35841ff42d49ff0081a3ff008dd7d563ff00b0eae2672c54bdfd9fc5d3 ,
                        0x4e9a1ea57fa94aa3755fbdf332bfe195355ffa0ed9ff00dfa7a3fe195355ff00 ,
                        0xa0ed9ffdfa7ad5ff0086b083fe85a93ff0347ff1ba3fe1ac20ff00a16a4ffc0d ,
                        0x1ffc6eb83d9f0dff0037fe97fe461cb9777fccd6f863f03757f87de2c83556d6 ,
                        0x2de7b7f2de29a18e36064561c0e7d1829fc2bda2bc07fe1ac20ffa16a4ff00c0 ,
                        0xd1ff00c6e9d1fed5d6c586ff000dccabdcade027f2d82bd9c1e6793e069fb2a1 ,
                        0x52caf7da4ff3475d1c4e128c79612d3e67be515e65e1bfda1bc23afc890cf713 ,
                        0x68f3b718be40a99ff7c1200f76c57a5452a4d1a491bac91b80caea72181e841f ,
                        0x4afa4c3e2a862a3cd426a4bc8f469d58555783b8fa28a2ba8d428a2bc1bc5bfb ,
                        0x6afc34f05f89b54d0751b9d4bedfa6dc3dacfe4d99741229c30073cf22803de6 ,
                        0x8af9c3fe1bf3e13ffcfceadff8007fc68ff86fcf84ff00f3f3ab7fe001ff001a ,
                        0x76607d1f4578d699fb59780756f879acf8d609b501a1e957315a5c3bda112192 ,
                        0x42a142ae79fbc335ce58fede1f0b351bdb7b4827d59a79e458a35fb011966200 ,
                        0x1d7d4d16607d1145145200a2bcabe2e7ed2de0af827ad59e95e259ef12f2eedf ,
                        0xed31adadb1946cdc5724e78e54fe55c2ff00c37e7c27ff009f9d5bff00000ff8 ,
                        0xd3b303e8fa2be70ff86fcf84ff00f3f3ab7fe001ff001ad8f07feda5f0dfc71e ,
                        0x28d3340d2e7d51f51d4675b78164b22abb98e064e78145981eef45781f89bf6d ,
                        0xcf863e13f11ea9a25f5d6a46f74db992d27f26c8ba7991b15600e791906b33fe ,
                        0x1bf3e13ffcfceadff8007fc68b303e8fa2be70ff0086fcf84fff003f3ab7fe00 ,
                        0x1ff1aeabc19fb58f80bc79a4f89752d2e5d41ad3c3f62750be79ad0a6d8c6785 ,
                        0xe79638381ed45981ecb5e6be3ef80fa078d6692f21dda3ea6fcb5c5ba828e7d5 ,
                        0xd3804fb8209ee6bcebfe1bf3e13ffcfceadff8007fc68ff86fcf84ff00f3f3ab ,
                        0x7fe001ff001ae5c46168e2e1ecebc3991954a70aab966ae8c3d5bf667f15d8c8 ,
                        0xdf6392cb528bf84c72f96df886000fccd667fc33d78e7fe81517fe05c5ff00c5 ,
                        0x575fff000df9f09ffe7e756ffc003fe35a3e1dfdb7be19f8abc41a668ba7cdaa ,
                        0xc97fa8dcc7696e8d62403248c154139e0648af9a970b6064ee9c97cd7ea99e73 ,
                        0xcb2837a3679fff00c33d78e7fe81517fe05c5ffc551ff0cf5e39ff00a0545ff8 ,
                        0x1717ff00155f5f5709e2cf8cfe1cf05eb52e97a9b5d25d46aae7cb80b290c320 ,
                        0x839e6b92b70ee59868f3d6a928af36bfc8ca797e1a9abce4d2f55fe47cf9ff00 ,
                        0x0cf5e39ffa0545ff0081717ff1547fc33d78e7fe81517fe05c5ffc557b5ffc34 ,
                        0x97833fe7adeffe031ff1a3fe1a4bc19ff3d6f7ff00018ff8d70ff65e45ff0041 ,
                        0x1ff9347fc8c7eab82ff9f9f8aff23c53fe19ebc73ff40a8bff0002e2ff00e2a8 ,
                        0xff00867af1cffd02a2ff00c0b8bff8aaf6bff8692f067fcf5bdffc063fe347fc ,
                        0x3497833fe7adeffe031ff1a3fb2f22ff00a08ffc9a3fe41f55c17fcfcfc57f91 ,
                        0xe29ff0cf5e39ff00a0545ff81717ff00155f457c1ff0ade7837c0561a6ea11ac ,
                        0x37caf23cc8ac18025c91c8e0fcb8ac0ff8692f067fcf5bdffc063fe35e996379 ,
                        0x1ea1636f7516ef2a78d654dc3070c01191f8d7b994e032ec3d6955c1d4e776b3 ,
                        0xd53dfd12ec76e16861e9cdca8caecf8f7e3a5f7dbfe29eb8c0e56378e11ff018 ,
                        0xd41fd735c156e78e6fbfb4fc69aedde7226be9997e9bce3f4c561d7e5d8ca9ed ,
                        0x71352a77937f89f315a5cd5252eed8514515c6641451450015e8bf09fe305ffc ,
                        0x3fd423b6b9792ef4291b12db1393164fdf8f3d0fb743f5c11e754574e1f11570 ,
                        0xb5155a2ecd1a53a92a525383b33f402ceee1bfb586e6de459ade645923914e43 ,
                        0x2919047d4515e53fb34f88a4d5bc0b2d84cdb9f4db831a7fd7361b97f52e3e80 ,
                        0x515fb960f12b19878575f697fc3fe27db51a8ab53535d4f5cafc5ef88dac7fc2 ,
                        0x43f10bc4faa6edff006dd52eae437aef959bfad7ec578bb571e1ff0009eb5aa1 ,
                        0x3b458d94f73bbd3646cdfd2bf150b162493927a9af4226acf77f857fb1a78e3e ,
                        0x2f7826cbc53a35f6876da75e348b125f5c4a92fc8ec8c485898632a7bd75bff0 ,
                        0xee9f899ff416f0bffe065c7ff18afb27f659d1ff00b0ff0067af02db631bf4e5 ,
                        0xb9ff00bfacd2ff00ecf5eab4b982c7e7a7c61f859aafecf7fb24a785f59b9b3b ,
                        0x8d5758f152dcc8da7c8ef19885b9c0cb2a9e0c4bdbbd780fc09d1ffb7be34781 ,
                        0xac48dc926b3685c7fb0b2ab37e80d7d5ff00f052cd636e9fe04d2958fef25bbb ,
                        0xa75edf288954ff00e3cd5e11fb14e8ff00daff00b4878572331da8b8b96ff80c ,
                        0x126dff00c78ad57411faa3451456651f99ff00f0501d63fb4be3fbdb6ec8d3f4 ,
                        0xbb6b6c7a677cbffb54579d7c11fd9efc4bf1f2e35787c3b73a6da9d2d2279db5 ,
                        0x19648c1f30b050bb11b27e46eb8abdfb59eb1fdb9fb4578e2e376e11de8b51ed ,
                        0xe5469163ff001cafa53fe09cb269da3f83bc63a85ddf5adac9757f0c01669551 ,
                        0x888e32d9e4f4fde9ad3a13d4f34ff8774fc4cffa0b785fff00032e3ff8c57a0f ,
                        0xc03fd88fc69f0cfe2e787bc51af6a3a0cfa5e9b2492c91d9dc4cf29262754c06 ,
                        0x89470cca4f3d01afb23fe12ad13fe83161ff008149fe3543e216b6ba1fc3bf12 ,
                        0xeae8ea56d34ab9ba570783b6166073f854dd8ec7e3b78bb563aff8b35ad4c9c9 ,
                        0xbdbd9ee73ebbe466feb5ecff000c7f62cf1dfc56f03e9be29d26fb43b6d3affc ,
                        0xc30c77d71324b8491a324858987250e39e98af02afd67fd9c6fb47f0e7c08f03 ,
                        0x58c9a9d8c320d2a199e36b94055a41e6104678397354f411f1effc3ba7e267fd ,
                        0x05bc2fff0081971ffc62bd4fe1bfec7fe36f00fc21f89da0c977a35c6bde2682 ,
                        0xd6dad1edee64f29234763207668c119573d01ce3b57d756fe22d2aea64860d4e ,
                        0xce695ce1638ee11989f600d68d4dd8ec7e6aff00c3bdfe297fcf5d07ff00039f ,
                        0xff008dd7cf5e29f0edd7847c4daae857ad13dee9b752d9ced036e43246e55b69 ,
                        0xc0c8c835fb5c4e393c0afc54f17eac75ef166b7a993b8dedf4f739f5df2337f5 ,
                        0xa69dc47a5fc21fd957c6df1afc3136bde1e6d352c22ba6b43f6db968dcbaaa31 ,
                        0x20053c61c7eb5edbf04ff620f1ef81fe2bf863c41adc9a39d334dbc5ba945bdd ,
                        0x33c9f2825700a0cfcdb7bd7bb7ec33a3ff0065fece3a04a46d6beb8bab93ff00 ,
                        0x7f9907e918af7ea4d8c2bc8be3f7c2d9fc65a7c3ac697179baad8a147857ef4f ,
                        0x0e7381eaca492077c9ef8af5da2b871985a78da32a15767fd5ccab528d68384b ,
                        0x667e7d329562ac08607041ea292bec0f881f03741f1d4ef7abbb4ad51fef5d5b ,
                        0xa82b21f574e8c7dc107debc7758fd993c5362cc6c67b2d4e3edb6431bfe21860 ,
                        0x7e75f9462f87f1d8693e48f3c7baff002dcf97ab80ad4de8aebc8f21a2bb6bbf ,
                        0x82fe36b36224f0f5cb63fe78b249ff00a0b1acf7f867e2d8faf86b553feed9c8 ,
                        0x7f90af1a583c4c7495292f9338dd1a8b78bfb8e7adaddaeae22853efc8e117ea ,
                        0x4e2bef59e44d1f4991c7115ac04fe0abff00d6af913c0ff0d7c48de31d11aef4 ,
                        0x0d4aded16f6179659ad5d1550382492471c0afafb52b25d4b4fbab473b52e226 ,
                        0x88b0ec1811fd6befb85f0f52953ad3946cdd92bf95ff00ccf772ca728c67268f ,
                        0x812491a6919dce5989627d49a6d769ac7c1df1868f7d2db36857774118859ad6 ,
                        0x332a38ec415f5f7e6a97fc2b1f177fd0b5aa7fe023ff00857c0cb07898b6a54e ,
                        0x57f46784e8d44ece2fee398a2ba7ff008563e2effa16b54ffc047ff0a3fe158f ,
                        0x8bbfe85ad53ff011ff00c2a7eab88ff9f6fee62f6553f95fdc7314574fff000a ,
                        0xc7c5dff42d6a9ff808ff00e147fc2b1f177fd0b5aa7fe023ff00851f55c47fcf ,
                        0xb7f730f6553f95fdc7314574ff00f0ac7c5dff0042d6a9ff00808ffe147fc2b1 ,
                        0xf177fd0b5aa7fe023ff851f55c47fcfb7f730f6553f95fdc7aff00ec9fbbcaf1 ,
                        0x413f7375b63ebfbdcff4a2ba6fd9cfc237fe18f0bea326a567358dddd5d710ce ,
                        0x851b62a800e0fb96a2bf62c8e9ca8e5f4a13567afe2db3ebb05170c3c53feb52 ,
                        0xe7ed37ac7f617c00f1ddd6ed85b4b96d811eb2e22ffd9ebf22abf4f7f6f1d63f ,
                        0xb2ff00675d56dc36d3a85edadafd7120971ff90abf377c0da4ff00c241e36f0f ,
                        0xe978cfdb750b7b6c6339df22aff5af7e3b1d6c75bf8f7c4d676f1410788b5682 ,
                        0x08942471477d2aaa28180000d8000ed5edbfb1d6bfe24f157ed0de18b5bcd7f5 ,
                        0x4bab387ed17134335e48e8c120720152d83f36dafd2ffecbb2ff009f483fefd2 ,
                        0xff00853e2b1b681c3c76f146ff00de5400d2e61d8fcf5ff828e6b1f6af8b5e1f ,
                        0xd395b2b67a3ac84760d24d267f445aadff0004e9d1fed9f19b58bf61f25968d2 ,
                        0x63fdf796203f40d5c87edc5ac7f6b7ed1fe228c3164b286d6d57f0811c8ffbe9 ,
                        0xcd7b3ffc134f48ff0091f75461ff003e76c87fefeb37fec94fa0ba9f71514565 ,
                        0x78af561a0785f58d4c9c0b2b39ae49f4d885bfa54147e3afc4ad63fe121f88be ,
                        0x29d53717fb6eab757018f70f2b37f5ab3a0fc21f1c78ab4b8753d1bc1fae6aba ,
                        0x74db8477765a7cb2c4fb58a9c32a9070411f515c9b3166249c927249afd6bfd9 ,
                        0x5348fec4fd9e3c0b6f8c6fd3c5cf4c7fad7697ff0067ad1bb127e6927ecfdf13 ,
                        0x5d957fe15ff89464e32da54e07fe835fa3bfb49ca9e0afd977c536f195548349 ,
                        0x8f4e5ec30e520c0fc1abd9ebe6efdbfb58fecdfd9f66b6c81fda1a9db5b63d71 ,
                        0xba5ffda5537bb19f99c88d23aaa82ccc7000ea4d77bff0a07e267fd13ef137fe ,
                        0x0a67ff00e26b1fe17e8fff00090fc4af0a697b778bcd5ad6dcafb34ca0fe86bf ,
                        0x676a9bb08fcdafd927e07f8d348f8fde19d4b5ef08eb7a4e9964679e4b9beb09 ,
                        0x618c30824080b32819dc56bf4968a2a1bb9473df11357ff847fe1ff89b54cedf ,
                        0xb169973739f4d91337f4afc5dafd6bfdab357fec4fd9dfc75719c6fb0fb37fdf ,
                        0xd758bff67afc968e369645441b9d8e028ee4d5444cfd7bfd9cf47fec3f811e03 ,
                        0xb52bb5bfb22de665f4322090feae6bd1ab3fc3da5ae87a0699a720c259db456e ,
                        0x3e88817fa5657c49f1d58fc33f02eb5e26d44ffa2e9d6ed2ecce0c8fd1107bb3 ,
                        0x1551f5a819f34fed6df1e3c69a2f8a2e741f8777ef64fe1ad3c6a9aedcc31a49 ,
                        0xb43c91a24677291c070c7d437b57d45e0df1243e30f08e89aedbe043a9d94378 ,
                        0x801e82440d8fc338af9ebe00fc24bcf137c0cf19eb5e224f33c4df11e1b9b9b8 ,
                        0x9241ca452238800cf41f3171e9b97d2b43f637f88d6cdfb32a5d6ab31893c2ad ,
                        0x756d78ec3e648e3ccc323da3703fe035423e8eaf9cfc65e2ed6eeff6d4f04785 ,
                        0xec757bdb6d0e0d165bdbfb086e196199cadc6dde80e1b0445d45729f0dbc03ad ,
                        0xfed6705e78f7c6faf6b3a4f86ae6e248f42d0349bb3024712315f35c80773641 ,
                        0x19ea48278181589f077e1fc8ff00b59fc44f0fddebfa96bd67a47879b4c8352b ,
                        0xf9775d2a4a21c29718cb20791437fb3dba000fb4a8af9d3f625f166a37bf0d75 ,
                        0xdf0ff882f259f58f0c6b1716570f75297754ceec966e787f3473fddae2b40f8b ,
                        0xed71ae7c45f8e9abddddc9e0fd1f3a2f86b4b59d922bb9010a64099c65891f36 ,
                        0x0e048ffdca5619f60515f317863e0bebff0018bc351f8cbe2c78af5987ed90fd ,
                        0xb2dfc3ba3dc9b3b4b1808dca180e59f6e0e4f23a126bcbbe18f8a357d1ff0061 ,
                        0x3f881ae5c6afa84b3cb7d25bd94b3dd3b3c31b35bc415093f2fccce78c739a2c ,
                        0x23eeea2be51f863f083c49f1f3e14f87efbc71e29d7746d2469905be99a5e957 ,
                        0x5e5b48ab1aafdaae5d81323bb0dcaa780a47ad647ecf5aa6b1f103f657f16db6 ,
                        0xb3e2bd76c97c3b7d74b6fabe977861ba6821812511f98c18edcb371e840c8c51 ,
                        0x619eebf043e2d5dfc5a9bc6d3b5adbc1a5e8faecda558cb0eedd3a4601def93d ,
                        0x4ee07803ad7a857c63fb1cfc0e8b54f867a278e2e7c65e2ad2b3a8cb7cfa6d8e ,
                        0xa621b09fc9976932c7b4960de561b9190315d4f87eebc4ff00b5e6b9ab5fc5e2 ,
                        0x0d47c25f0ab4fba6b3b587497f26ef5675fbced27554e471c8e718c824311f52 ,
                        0xd15f1c7c23d261f863fb557c43d234ad4b58b8f0c7873c3ff6a92d6eaf649849 ,
                        0x29481ce41e09f99f1c7157be0e7c2f7fda63c1173f107c7be26d6daeb54b8b81 ,
                        0x6369a76a0d6f6fa6c48e50796838c82a793d80ce4e4d2b01f5cd15f0d7c17f88 ,
                        0xdacf807e08f89fe29788fc5de22f11dc69f79368fa4e9b7fa833d8ddb6d8c472 ,
                        0xec6cb13b99b2777011bbd68eb3e12f0d47f0f9bc55f173e29ea6de36bdb53791 ,
                        0x58e9fab7966c9dd77471436e99e41c03c6339ec334ec33ed4a2bcabf65ed53c5 ,
                        0x3acfc0ef0d5e78c5ae24d6a48e4265bb07cf921f31bca6933c9629b793c91827 ,
                        0x9268a903c5bfe0a45ac7d9fe1cf8534b0c47dab556b8dbea238997ff006a8af9 ,
                        0x33f65cd23fb6ff00684f025be376cd492e7a67fd5032ff00ec95fa99e33f86be ,
                        0x16f888b68be26d06cb5b1685cdb8bc883f97bb1bb19e99dabf9565f86be07f80 ,
                        0x3c1dad5beafa278474ad3353b7dde55d5b5b8591372956c1ed95247e3557d047 ,
                        0x734514548cfc7efda1b58feddf8e7e3bbc0c594eb1731293dd52428bfa28acef ,
                        0x04fc60f19fc37b3b8b4f0cf88af346b6b8904b2c76cc00760319391e95fa9979 ,
                        0xfb39fc31d42ee7bab9f03e8b3dc4eed2cb2bdb02cecc72493ea49a87fe199fe1 ,
                        0x5ffd085a1ffe02ad5f31363f3d7c33fb66fc5bf0ddf4739f1436ab0ab65ed753 ,
                        0x8239a3907a1380c07fbac0d7d83ac7ed0563f16bf641f1af8aad223a7de2e9d3 ,
                        0xe9b7b69bb77913ba84c29eea44aa41f7c7506bd27fe199fe15ff00d085a1ff00 ,
                        0xe02ad6c587c19f03697e1dd4741b4f0b69b6fa36a4c8f79631c004539520a965 ,
                        0xef8207e54ae867e3857ace8dfb567c56f0fe9363a5e9fe2fb8b6b0b1812dade1 ,
                        0x5b68088e3450aaa331e78000e6bf47bfe199fe15ff00d085a1ff00e02ad1ff00 ,
                        0x0ccff0affe842d0fff0001569f32158f8abe18fedfbe3cf0eeb16c9e2e783c51 ,
                        0xa333859ffd1e382e634ee6368c2a923ae181cf4c8ea3d2ff00e0a1be30b3d73e ,
                        0x1b7c3d3a75c2dcd86ad7126a504abd248d615dac3f09ff005afa2bfe199fe15f ,
                        0xfd085a1ffe02ad6beb1f05bc0be20d334ad3b52f0ae997b63a546d0d8dbcd006 ,
                        0x4b7438caa0ec0ed1f952ba19f903e1fd7aff00c2fadd8eafa5dc1b4d46c6659e ,
                        0xde7550c63914e5580208e0fa8af568ff006c6f8c713861e36b82473f359db30f ,
                        0xc8c75fa23ff0ccff000aff00e842d0ff00f01568ff008667f857ff00421687ff ,
                        0x0080ab4f9901f3efecc3fb6f6a5e36f15d8f847c7715b1bbd41c4363abdb208b ,
                        0x74a7eec72a0e32c7e5565c738047391f67579de9ff00b3bfc34d2750b6beb3f0 ,
                        0x4e8f6d796d2acd0cd1db00d1ba905581ec4100d7a254b19f38fedf9ab7f66fec ,
                        0xf7756f9c7f686a56b6df5c1697ff006957e7a7c29d1ffe120f8a1e11d336ef17 ,
                        0x7abda40c3d9a6507f4cd7ebdf8c3c09e1ef881a7c561e24d22d75ab386513c70 ,
                        0x5e461d56400a8600f7c311f89ae7746f803f0e7c3baadaea7a6f83348b2d42d6 ,
                        0x412c1710db0578dc74607b114d3b08f40af94ff690ba97e397c66f0a7c18d3e6 ,
                        0x9174bb761ab788a680e0a46065533d8ed39e72332c67b57d59591a7f84742d27 ,
                        0x5abdd62c745d3ecf56be18babfb7b5449e7e41f9e40373741d4f6a433c523fd8 ,
                        0x93c07146a89ab78a9114615575760001d87cb5e05e1dd0ae7c00ff00b48fc28d ,
                        0x31a79b6e9ada9584721dd2b4280330e9f333453463df6d7dff00590be11d093c ,
                        0x46fe205d174f5d79e3f29b5416a82e593006d32e3711800633d851711f3a7c08 ,
                        0xfda3bc0fe19fd9e7c2f6915fc777e23b4b51629e1db51baf6e2e81202ac639c3 ,
                        0x920eee9f3727208acafd8a6d758d47e25fc65d73c42633adb6a51da5cf92db91 ,
                        0x250f399114ff007548551ec057d1fa1fc2ef0778675a9b57d27c2da3e99aa4b9 ,
                        0x2f7969631c729cf5f980c8cf7c75ad6d23c35a4787e5be974bd2ecb4d92fa76b ,
                        0x9bb7b4b7488dc4a7ac921503731eec79a00f87be2bde6bbf0dfe3a7c4df04786 ,
                        0x60717bf12e2b2fb0b2f011a562b3484f6196b919ed9cf6aecff6bef8769f0e7f ,
                        0x660f08e87a3c0d3e8da0ea96bf6c0a31e62f9728691ffde95c67de415f55dcf8 ,
                        0x5745bcf105aebb71a458cfadda4662b7d4a4b646b8850eecaa4846e51f33700f ,
                        0xf11f5ab9a869d69abd8cf657f6b0ded9dc218e6b7b88c491c8a7aab291820fa1 ,
                        0xa7703e74f8dbfb49685aff00c2bd674df877a9c3aeeb57ba54f348d699dba759 ,
                        0xac45a69a638fddb05caaa9e77b28c5789f8ba36f0cff00c13cbc276601136b5a ,
                        0x9aee5eedba79e55ffc7635afb6743f863e10f0ce977da6e95e18d22c34fbf431 ,
                        0x5ddb41651ac772841056418f9c60918391c9f5a9efbe1f785f54d16c748bcf0d ,
                        0xe9377a4d8b2b5a58cd6313c16e402018d0aed520120600eb45c0a9a932f817e1 ,
                        0x6dd32feed346d19c8c1e82180f7ff80d7cb1f056d67d03fe09ff00e30bdb7473 ,
                        0x2ea16fa94bf2024ed23c866fc1509fc2becbbeb1b6d52c6e2cef2de2bbb3b88d ,
                        0xa19ade740f1c88c30caca7820824107ae6aa697e1bd2743d15347d3b4bb3b0d2 ,
                        0x515916c2dadd23802b125804036e096248c73934867ccbf06bc6de1cd53f6548 ,
                        0x7c09e1ed76d2f3c633f85f51316976ef99c4c5642ca54746df20c03c91920100 ,
                        0xd677ecbbfb42f81bc07fb38e93a75cea0ade22d3e4b888e870296bcbb99e7768 ,
                        0xd624eafb83a0cf40739e95f4b7857e187843c0f7535cf87bc31a4e8b73302af3 ,
                        0x58d9c71395272572a338ce38e9c0a34ff85fe0fd2bc472f882cbc2fa45aeb723 ,
                        0x33bea10d946b3166fbcdb80ce4f73d4e4d1711f2ff00ec957b70df10fe3878b3 ,
                        0xc7325a69d711dd476da835c4abe44077cde6445cf1b576a2fe02b88f1d5e7c35 ,
                        0xf09f80fc6363f0efe28f8996c2ea39a483c396714e6c0cadc3209cc18098cf1b ,
                        0xc6e0002c73cfbcfed53f046fbc5df0f8af81f43b06befed88f56d534e8234b76 ,
                        0xd58287ddbd8637be5b3f31c9e79ce018bfe17d78be7f0ea68be15f823e27d375 ,
                        0xaf27ecf045a85aa5b69b6ad8c03e6640645f4c2e718e334c0f31f8c7a3e9fe2a ,
                        0xfd85fc1d75e09b7373a3e8f2dadede411ae5c044963b8de073912c8598fa02dd ,
                        0x39aeefc2de33fd99fc0be1783c4fa447e1ab69842264b708b71a9231c1d811cb ,
                        0x481f381d877ce39af4bfd9bfe10cdf067e12e9fe1ad4278ef2fd9e4b9bc31f31 ,
                        0x89243ca2e7a80001ee413deb7b4ef82de00d23581aad8f82f41b4d455b7adc43 ,
                        0xa744ae8dfde5c2f07dc517036bc1be261e32f0b699adae9d7ba52df4226167a8 ,
                        0xc6239e307a07504e3d7af4228adaa2a467ffd9
                    End

                    LayoutCachedWidth =2160
                    LayoutCachedHeight =864
                    TabIndex =3
                End
                Begin TextBox
                    Visible = NotDefault
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Left =10230
                    Top =840
                    Width =144
                    Height =14
                    TabIndex =1
                    Name ="SGrpID"
                    ControlSource ="SGrpID"

                    LayoutCachedLeft =10230
                    LayoutCachedTop =840
                    LayoutCachedWidth =10374
                    LayoutCachedHeight =854
                End
                Begin TextBox
                    Visible = NotDefault
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Left =9360
                    Top =840
                    Width =144
                    Height =14
                    TabIndex =2
                    Name ="clientID"
                    ControlSource ="clientID"

                    LayoutCachedLeft =9360
                    LayoutCachedTop =840
                    LayoutCachedWidth =9504
                    LayoutCachedHeight =854
                End
            End
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =12015
            Name ="GroupHeader2"
            Begin
                Begin Label
                    TextAlign =3
                    TextFontFamily =34
                    Top =180
                    Width =1080
                    Height =285
                    Name ="Label6"
                    Caption ="Lessor: "
                    GroupTable =49
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedTop =180
                    LayoutCachedWidth =1080
                    LayoutCachedHeight =465
                    LayoutGroup =3
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =49
                End
                Begin Label
                    TextAlign =1
                    Left =1080
                    Top =468
                    Width =3435
                    Height =288
                    Name ="Label8"
                    Caption ="401 East Las Olas Blvd., Suite 1720"
                    GroupTable =49
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =1080
                    LayoutCachedTop =468
                    LayoutCachedWidth =4515
                    LayoutCachedHeight =756
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =3
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =49
                End
                Begin Label
                    TextAlign =1
                    Left =1080
                    Top =756
                    Width =3435
                    Height =288
                    Name ="Label9"
                    Caption ="Fort Lauderdale, FL 33301"
                    GroupTable =49
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =1080
                    LayoutCachedTop =756
                    LayoutCachedWidth =4515
                    LayoutCachedHeight =1044
                    RowStart =2
                    RowEnd =2
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =3
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =49
                End
                Begin Line
                    BorderWidth =1
                    Left =45
                    Top =1440
                    Width =10800
                    Name ="Line77"
                    LayoutCachedLeft =45
                    LayoutCachedTop =1440
                    LayoutCachedWidth =10845
                    LayoutCachedHeight =1440
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    TextFontCharSet =186
                    TextAlign =1
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Left =5772
                    Top =180
                    Width =5412
                    Height =864
                    TabIndex =1
                    Name ="AddressBox"
                    ControlSource ="=[Lessee]"
                    GroupTable =49
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =5772
                    LayoutCachedTop =180
                    LayoutCachedWidth =11184
                    LayoutCachedHeight =1044
                    RowEnd =2
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =3
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =49
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    TextFontCharSet =186
                    TextAlign =3
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Left =4512
                    Top =180
                    Width =1260
                    Height =864
                    Name ="LesseeBox"
                    ControlSource ="=\"Lessee: \" & IIf(IsNull([clientconm2]),Null,Chr(13) & Chr(10) & \"Co-Lessee: "
                        "\")"
                    GroupTable =49
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =4512
                    LayoutCachedTop =180
                    LayoutCachedWidth =5772
                    LayoutCachedHeight =1044
                    RowEnd =2
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =3
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =49
                    Begin
                        Begin Label
                            TextAlign =1
                            Left =1080
                            Top =180
                            Width =3435
                            Height =285
                            Name ="Label7"
                            Caption ="Fleet Advantage, LLC"
                            GroupTable =49
                            LeftPadding =0
                            TopPadding =0
                            RightPadding =0
                            BottomPadding =0
                            LayoutCachedLeft =1080
                            LayoutCachedTop =180
                            LayoutCachedWidth =4515
                            LayoutCachedHeight =465
                            ColumnStart =1
                            ColumnEnd =1
                            LayoutGroup =3
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =49
                        End
                    End
                End
                Begin EmptyCell
                    Top =468
                    Width =1080
                    Height =288
                    Name ="EmptyCell385"
                    GroupTable =49
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedTop =468
                    LayoutCachedWidth =1080
                    LayoutCachedHeight =756
                    RowStart =1
                    RowEnd =1
                    LayoutGroup =3
                    GroupTable =49
                End
                Begin EmptyCell
                    Top =756
                    Width =1080
                    Height =288
                    Name ="EmptyCell386"
                    GroupTable =49
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedTop =756
                    LayoutCachedWidth =1080
                    LayoutCachedHeight =1044
                    RowStart =2
                    RowEnd =2
                    LayoutGroup =3
                    GroupTable =49
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    TextAlign =1
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Left =5772
                    Top =1044
                    Width =5412
                    Height =300
                    TabIndex =3
                    Name ="SubLesseeName"
                    ControlSource ="=IIf(IsNull([subleasee]),Null,[subleasee])"
                    GroupTable =49
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =5772
                    LayoutCachedTop =1044
                    LayoutCachedWidth =11184
                    LayoutCachedHeight =1344
                    RowStart =3
                    RowEnd =3
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =3
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =49
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    TextAlign =3
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Left =4512
                    Top =1044
                    Width =1260
                    Height =300
                    TabIndex =2
                    Name ="SubLesseeLabel"
                    ControlSource ="=IIf(IsNull([subleasee]),Null,\"Sublessee: \")"
                    GroupTable =49
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =4512
                    LayoutCachedTop =1044
                    LayoutCachedWidth =5772
                    LayoutCachedHeight =1344
                    RowStart =3
                    RowEnd =3
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =3
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =49
                End
                Begin Label
                    TextAlign =1
                    TextFontFamily =34
                    Left =36
                    Top =1512
                    Width =11292
                    Height =264
                    FontWeight =700
                    Name ="Label17"
                    Caption ="Lessee, through its duly authorized representative, hereby certifies to Lessor t"
                        "hat:"
                    GroupTable =41
                    LayoutCachedLeft =36
                    LayoutCachedTop =1512
                    LayoutCachedWidth =11328
                    LayoutCachedHeight =1776
                    LayoutGroup =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =41
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    TextAlign =1
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Left =36
                    Top =3120
                    Width =11292
                    Height =870
                    TabIndex =7
                    Name ="Text22"
                    ControlSource ="=\"4. Lessee acknowledges that signature of this document (a) constitutes uncond"
                        "itional acceptance of the Exhibit A Equipment under and subject to the terms of "
                        "the Master Equipment Lease Agreement No. \" & [MLNo] & \" dated as of  \" & [mlM"
                        "O] & \" \" & [MLDAY] & \", \" & [MLYR] & \" by and between Lessor and Lessee (as"
                        " may be amended from time to time) and the Schedule thereunder relating to the E"
                        "xhibit A Equipment, (b) that such acceptance is not on a trial basis, and (c) au"
                        "thorizes commencement of the Lease Term under the Master Equipment Lease Agreeme"
                        "nt and Schedule.\""
                    GroupTable =41

                    LayoutCachedLeft =36
                    LayoutCachedTop =3120
                    LayoutCachedWidth =11328
                    LayoutCachedHeight =3990
                    RowStart =4
                    RowEnd =4
                    LayoutGroup =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =41
                End
                Begin TextBox
                    SpecialEffect =5
                    BorderWidth =1
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1308
                    Top =8064
                    Width =8304
                    Height =348
                    TabIndex =16
                    Name ="contactFullName"
                    ControlSource ="FullNm"
                    GroupTable =51
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =1308
                    LayoutCachedTop =8064
                    LayoutCachedWidth =9612
                    LayoutCachedHeight =8412
                    RowStart =4
                    RowEnd =4
                    ColumnStart =1
                    ColumnEnd =3
                    LayoutGroup =4
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =51
                    Begin
                        Begin Label
                            TextAlign =3
                            Top =8064
                            Width =1305
                            Height =348
                            Name ="Label71"
                            Caption ="Name:"
                            GroupTable =51
                            LeftPadding =0
                            TopPadding =0
                            RightPadding =0
                            BottomPadding =0
                            LayoutCachedTop =8064
                            LayoutCachedWidth =1305
                            LayoutCachedHeight =8412
                            RowStart =4
                            RowEnd =4
                            LayoutGroup =4
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =51
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =5
                    BorderWidth =1
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1308
                    Top =8712
                    Width =8304
                    Height =348
                    TabIndex =17
                    Name ="contactTitle"
                    ControlSource ="contactTitle"
                    GroupTable =51
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =1308
                    LayoutCachedTop =8712
                    LayoutCachedWidth =9612
                    LayoutCachedHeight =9060
                    RowStart =6
                    RowEnd =6
                    ColumnStart =1
                    ColumnEnd =3
                    LayoutGroup =4
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =51
                    Begin
                        Begin Label
                            TextAlign =3
                            Top =8712
                            Width =1305
                            Height =348
                            Name ="Label72"
                            Caption ="Title:"
                            GroupTable =51
                            LeftPadding =0
                            TopPadding =0
                            RightPadding =0
                            BottomPadding =0
                            LayoutCachedTop =8712
                            LayoutCachedWidth =1305
                            LayoutCachedHeight =9060
                            RowStart =6
                            RowEnd =6
                            LayoutGroup =4
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =51
                        End
                    End
                End
                Begin TextBox
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1308
                    Top =6612
                    Width =8304
                    Height =408
                    TabIndex =11
                    Name ="Text91"
                    ControlSource ="clientCompanyName"
                    GroupTable =51
                    LeftPadding =0
                    RightPadding =0

                    LayoutCachedLeft =1308
                    LayoutCachedTop =6612
                    LayoutCachedWidth =9612
                    LayoutCachedHeight =7020
                    ColumnStart =1
                    ColumnEnd =3
                    LayoutGroup =4
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =51
                    Begin
                        Begin Label
                            TextFontCharSet =186
                            TextAlign =3
                            TextFontFamily =34
                            Top =6612
                            Width =1305
                            Height =408
                            Name ="Label92"
                            Caption ="Lessee:"
                            GroupTable =51
                            LeftPadding =0
                            RightPadding =0
                            LayoutCachedTop =6612
                            LayoutCachedWidth =1305
                            LayoutCachedHeight =7020
                            LayoutGroup =4
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =51
                        End
                    End
                End
                Begin Label
                    TextFontCharSet =204
                    TextAlign =2
                    TextFontFamily =34
                    Left =1308
                    Top =8412
                    Width =4932
                    Height =300
                    FontSize =8
                    ForeColor =5855577
                    Name ="Label95"
                    Caption ="Please Print"
                    GroupTable =51
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =1308
                    LayoutCachedTop =8412
                    LayoutCachedWidth =6240
                    LayoutCachedHeight =8712
                    RowStart =5
                    RowEnd =5
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =4
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =51
                End
                Begin Label
                    TextFontCharSet =204
                    TextAlign =2
                    TextFontFamily =34
                    Left =1308
                    Top =9060
                    Width =4932
                    Height =300
                    FontSize =8
                    ForeColor =5855577
                    Name ="Label97"
                    Caption ="Please Print"
                    GroupTable =51
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =1308
                    LayoutCachedTop =9060
                    LayoutCachedWidth =6240
                    LayoutCachedHeight =9360
                    RowStart =7
                    RowEnd =7
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =4
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =51
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    TextAlign =1
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1308
                    Top =7080
                    Width =8304
                    Height =300
                    TabIndex =13
                    Name ="Text127"
                    ControlSource ="clientCoNm2"
                    GroupTable =51
                    LeftPadding =0
                    RightPadding =0

                    LayoutCachedLeft =1308
                    LayoutCachedTop =7080
                    LayoutCachedWidth =9612
                    LayoutCachedHeight =7380
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =3
                    LayoutGroup =4
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =51
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    TextFontCharSet =186
                    TextAlign =3
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Top =7080
                    Width =1305
                    Height =300
                    TabIndex =12
                    Name ="Text112"
                    ControlSource ="=IIf(IsNull([clientconm2]),Null,\"Co-Lessee:\")"
                    GroupTable =51
                    LeftPadding =0
                    RightPadding =0

                    LayoutCachedTop =7080
                    LayoutCachedWidth =1305
                    LayoutCachedHeight =7380
                    RowStart =1
                    RowEnd =1
                    LayoutGroup =4
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =51
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    TextFontCharSet =163
                    TextAlign =1
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1308
                    Top =11652
                    Width =8304
                    Height =300
                    FontSize =18
                    FontWeight =700
                    TabIndex =31
                    ForeColor =255
                    Name ="Text128"
                    ControlSource ="=IIf(IsNull([DST1]) Or IsNull([DCITY]),\"DOMICILE LOCATION REQUIRED\",Null)"
                    GroupTable =51
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =1308
                    LayoutCachedTop =11652
                    LayoutCachedWidth =9612
                    LayoutCachedHeight =11952
                    RowStart =15
                    RowEnd =15
                    ColumnStart =1
                    ColumnEnd =3
                    LayoutGroup =4
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =51
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    SpecialEffect =5
                    BorderWidth =1
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1308
                    Top =7416
                    Width =4932
                    Height =348
                    TabIndex =14
                    Name ="Text146"
                    GroupTable =51
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =1308
                    LayoutCachedTop =7416
                    LayoutCachedWidth =6240
                    LayoutCachedHeight =7764
                    RowStart =2
                    RowEnd =2
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =4
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =51
                    Begin
                        Begin Label
                            TextAlign =3
                            Top =7416
                            Width =1305
                            Height =348
                            Name ="Label69"
                            Caption ="By:"
                            GroupTable =51
                            LeftPadding =0
                            TopPadding =0
                            RightPadding =0
                            BottomPadding =0
                            LayoutCachedTop =7416
                            LayoutCachedWidth =1305
                            LayoutCachedHeight =7764
                            RowStart =2
                            RowEnd =2
                            LayoutGroup =4
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =51
                        End
                    End
                End
                Begin Label
                    TextFontCharSet =163
                    TextAlign =2
                    TextFontFamily =34
                    Left =1308
                    Top =7764
                    Width =4932
                    Height =300
                    FontSize =9
                    ForeColor =5855577
                    Name ="Label149"
                    Caption ="Signature"
                    GroupTable =51
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =1308
                    LayoutCachedTop =7764
                    LayoutCachedWidth =6240
                    LayoutCachedHeight =8064
                    RowStart =3
                    RowEnd =3
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =4
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =51
                End
                Begin EmptyCell
                    Top =7764
                    Width =1305
                    Height =300
                    Name ="EmptyCell445"
                    GroupTable =51
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedTop =7764
                    LayoutCachedWidth =1305
                    LayoutCachedHeight =8064
                    RowStart =3
                    RowEnd =3
                    LayoutGroup =4
                    GroupTable =51
                End
                Begin EmptyCell
                    Top =9060
                    Width =1305
                    Height =300
                    Name ="EmptyCell478"
                    GroupTable =51
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedTop =9060
                    LayoutCachedWidth =1305
                    LayoutCachedHeight =9360
                    RowStart =7
                    RowEnd =7
                    LayoutGroup =4
                    GroupTable =51
                End
                Begin EmptyCell
                    Top =8412
                    Width =1305
                    Height =300
                    Name ="EmptyCell481"
                    GroupTable =51
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedTop =8412
                    LayoutCachedWidth =1305
                    LayoutCachedHeight =8712
                    RowStart =5
                    RowEnd =5
                    LayoutGroup =4
                    GroupTable =51
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    SpecialEffect =5
                    BorderWidth =1
                    TextAlign =1
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1308
                    Top =10356
                    Width =8304
                    Height =348
                    TabIndex =26
                    Name ="Text487"
                    ControlSource ="=IIf(IsNull([subleasee]),Null,\" \")"
                    GroupTable =51
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =1308
                    LayoutCachedTop =10356
                    LayoutCachedWidth =9612
                    LayoutCachedHeight =10704
                    RowStart =11
                    RowEnd =11
                    ColumnStart =1
                    ColumnEnd =3
                    LayoutGroup =4
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =51
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    SpecialEffect =5
                    BorderWidth =1
                    TextAlign =1
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1308
                    Top =11004
                    Width =8304
                    Height =348
                    TabIndex =29
                    Name ="Text489"
                    ControlSource ="=IIf(IsNull([subleasee]),Null,\" \")"
                    GroupTable =51
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =1308
                    LayoutCachedTop =11004
                    LayoutCachedWidth =9612
                    LayoutCachedHeight =11352
                    RowStart =13
                    RowEnd =13
                    ColumnStart =1
                    ColumnEnd =3
                    LayoutGroup =4
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =51
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    SpecialEffect =5
                    BorderWidth =1
                    TextAlign =1
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1308
                    Top =9708
                    Width =4932
                    Height =348
                    TabIndex =21
                    Name ="Text498"
                    ControlSource ="=IIf(IsNull([subleasee]),Null,\" \")"
                    GroupTable =51
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =1308
                    LayoutCachedTop =9708
                    LayoutCachedWidth =6240
                    LayoutCachedHeight =10056
                    RowStart =9
                    RowEnd =9
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =4
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =51
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    TextAlign =1
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1308
                    Top =9360
                    Width =8304
                    Height =348
                    TabIndex =19
                    Name ="Text500"
                    ControlSource ="=IIf(IsNull([subleasee]),Null,[subleasee])"
                    GroupTable =51
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =1308
                    LayoutCachedTop =9360
                    LayoutCachedWidth =9612
                    LayoutCachedHeight =9708
                    RowStart =8
                    RowEnd =8
                    ColumnStart =1
                    ColumnEnd =3
                    LayoutGroup =4
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =51
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    TextAlign =3
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Top =9360
                    Width =1305
                    Height =348
                    TabIndex =18
                    Name ="Text501"
                    ControlSource ="=IIf(IsNull([subleasee]),Null,\"Sublessee: \")"
                    GroupTable =51
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedTop =9360
                    LayoutCachedWidth =1305
                    LayoutCachedHeight =9708
                    RowStart =8
                    RowEnd =8
                    LayoutGroup =4
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =51
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    TextAlign =3
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Top =9708
                    Width =1305
                    Height =348
                    TabIndex =20
                    Name ="Label486"
                    ControlSource ="=IIf(IsNull([subleasee]),Null,\"By: \")"
                    GroupTable =51
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedTop =9708
                    LayoutCachedWidth =1305
                    LayoutCachedHeight =10056
                    RowStart =9
                    RowEnd =9
                    LayoutGroup =4
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =51
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    TextAlign =2
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1308
                    Top =10056
                    Width =4932
                    Height =300
                    FontSize =9
                    TabIndex =24
                    ForeColor =5855577
                    Name ="Label499"
                    ControlSource ="=IIf(IsNull([subleasee]),Null,\"Signature\")"
                    GroupTable =51
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =1308
                    LayoutCachedTop =10056
                    LayoutCachedWidth =6240
                    LayoutCachedHeight =10356
                    RowStart =10
                    RowEnd =10
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =4
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =51
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    TextAlign =3
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Top =10356
                    Width =1305
                    Height =348
                    TabIndex =25
                    Name ="Label488"
                    ControlSource ="=IIf(IsNull([subleasee]),Null,\"Name:\")"
                    GroupTable =51
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedTop =10356
                    LayoutCachedWidth =1305
                    LayoutCachedHeight =10704
                    RowStart =11
                    RowEnd =11
                    LayoutGroup =4
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =51
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    TextAlign =3
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Top =11004
                    Width =1305
                    Height =348
                    TabIndex =28
                    Name ="Label490"
                    ControlSource ="=IIf(IsNull([subleasee]),Null,\"Title:\")"
                    GroupTable =51
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedTop =11004
                    LayoutCachedWidth =1305
                    LayoutCachedHeight =11352
                    RowStart =13
                    RowEnd =13
                    LayoutGroup =4
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =51
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    BorderWidth =1
                    TextAlign =2
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1308
                    Top =10704
                    Width =4932
                    Height =300
                    FontSize =8
                    TabIndex =27
                    ForeColor =5855577
                    Name ="Label493"
                    ControlSource ="=IIf(IsNull([subleasee]),Null,\"Please Print\")"
                    GroupTable =51
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =1308
                    LayoutCachedTop =10704
                    LayoutCachedWidth =6240
                    LayoutCachedHeight =11004
                    RowStart =12
                    RowEnd =12
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =4
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =51
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    BorderWidth =1
                    TextAlign =2
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1308
                    Top =11352
                    Width =4932
                    Height =300
                    FontSize =8
                    TabIndex =30
                    ForeColor =5855577
                    Name ="Label495"
                    ControlSource ="=IIf(IsNull([subleasee]),Null,\"Please Print\")"
                    GroupTable =51
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =1308
                    LayoutCachedTop =11352
                    LayoutCachedWidth =6240
                    LayoutCachedHeight =11652
                    RowStart =14
                    RowEnd =14
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =4
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =51
                End
                Begin EmptyCell
                    Top =10056
                    Width =1305
                    Height =300
                    Name ="EmptyCell548"
                    GroupTable =51
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedTop =10056
                    LayoutCachedWidth =1305
                    LayoutCachedHeight =10356
                    RowStart =10
                    RowEnd =10
                    LayoutGroup =4
                    GroupTable =51
                End
                Begin EmptyCell
                    Top =10704
                    Width =1305
                    Height =300
                    Name ="EmptyCell561"
                    GroupTable =51
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedTop =10704
                    LayoutCachedWidth =1305
                    LayoutCachedHeight =11004
                    RowStart =12
                    RowEnd =12
                    LayoutGroup =4
                    GroupTable =51
                End
                Begin EmptyCell
                    Top =11352
                    Width =1305
                    Height =300
                    Name ="EmptyCell567"
                    GroupTable =51
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedTop =11352
                    LayoutCachedWidth =1305
                    LayoutCachedHeight =11652
                    RowStart =14
                    RowEnd =14
                    LayoutGroup =4
                    GroupTable =51
                End
                Begin EmptyCell
                    Top =11652
                    Width =1305
                    Height =300
                    Name ="EmptyCell578"
                    GroupTable =51
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedTop =11652
                    LayoutCachedWidth =1305
                    LayoutCachedHeight =11952
                    RowStart =15
                    RowEnd =15
                    LayoutGroup =4
                    GroupTable =51
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    TextAlign =1
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Left =36
                    Top =5964
                    Width =11292
                    Height =480
                    TabIndex =10
                    Name ="Label33"
                    ControlSource ="=\"7.\011Lessee is not in default under the Master Equipment Lease Agreement or "
                        "the Schedule, and all of Lessee’s statements and promises contained in the Maste"
                        "r Equipment Lease Agreement and Schedule remain true and correct in all material"
                        " respects.\""
                    GroupTable =41

                    LayoutCachedLeft =36
                    LayoutCachedTop =5964
                    LayoutCachedWidth =11328
                    LayoutCachedHeight =6444
                    RowStart =7
                    RowEnd =7
                    LayoutGroup =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =41
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    TextAlign =1
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Left =36
                    Top =5052
                    Width =11292
                    Height =852
                    TabIndex =9
                    Name ="Label32"
                    ControlSource ="=\"6.\011The Exhibit A Equipment shall at all times be the sole and exclusive pr"
                        "operty of Lessor and Lessee shall have no right, title or interest in or to the "
                        "Exhibit A Equipment except the right to use the same upon the terms and conditio"
                        "ns herein so long, and only so long, as no Event of Default exists or is continu"
                        "ing under the Transaction Documents.\""
                    GroupTable =41

                    LayoutCachedLeft =36
                    LayoutCachedTop =5052
                    LayoutCachedWidth =11328
                    LayoutCachedHeight =5904
                    RowStart =6
                    RowEnd =6
                    LayoutGroup =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =41
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    TextAlign =1
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Left =36
                    Top =4056
                    Width =11292
                    Height =936
                    TabIndex =8
                    Name ="Label31"
                    ControlSource ="=\"5. Lessee's execution of this Certificate of Acceptance with respect to the E"
                        "xhibit A Equipment conclusively evidences that the Lessee has inspected and acce"
                        "pted each item of Exhibit A Equipment, and that the Lessee shall not reject, ret"
                        "urn or revoke this Agreement and any Schedule as to any such item of Exhibit A E"
                        "quipment. Lessee will not contest the authority of any person executing a Certif"
                        "icate of Acceptance on Lessee's behalf. Each item of Exhibit A Equipment is leas"
                        "ed in 'AS IS', 'WHERE IS' condition.\""
                    GroupTable =41

                    LayoutCachedLeft =36
                    LayoutCachedTop =4056
                    LayoutCachedWidth =11328
                    LayoutCachedHeight =4992
                    RowStart =5
                    RowEnd =5
                    LayoutGroup =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =41
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    TextAlign =1
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Left =36
                    Top =2820
                    Width =11292
                    TabIndex =6
                    Name ="Label20"
                    ControlSource ="=\"3. The Exhibit A Equipment is of a size, design, capacity and manufacture acc"
                        "eptable to Lessee and suitable for Lessee's purposes.\""
                    GroupTable =41

                    LayoutCachedLeft =36
                    LayoutCachedTop =2820
                    LayoutCachedWidth =11328
                    LayoutCachedHeight =3060
                    RowStart =3
                    RowEnd =3
                    LayoutGroup =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =41
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    TextAlign =1
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Left =36
                    Top =2328
                    Width =11292
                    Height =432
                    TabIndex =5
                    Name ="Label18"
                    ControlSource ="=\"2. All of the Exhibit A Equipment has been inspected and is determined to be "
                        "(a) complete, (b) properly installed, (c) functioning, and (d) in good working o"
                        "rder and in compliance with  all applicable  specifications.\""
                    GroupTable =41

                    LayoutCachedLeft =36
                    LayoutCachedTop =2328
                    LayoutCachedWidth =11328
                    LayoutCachedHeight =2760
                    RowStart =2
                    RowEnd =2
                    LayoutGroup =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =41
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    TextAlign =1
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Left =36
                    Top =1836
                    Width =11292
                    Height =432
                    TabIndex =4
                    Name ="Label19"
                    ControlSource ="=\"1. The Equipment described in<u> Exhibit A</u> attached hereto and incorporat"
                        "ed herein by this reference (hereinafter referred to as the “Exhibit A Equipment"
                        "”) has been duly delivered to the location identified in <u>Exhibit A</u> and wh"
                        "ere it will be garaged or used.\""
                    GroupTable =41
                    TextFormat =1

                    LayoutCachedLeft =36
                    LayoutCachedTop =1836
                    LayoutCachedWidth =11328
                    LayoutCachedHeight =2268
                    RowStart =1
                    RowEnd =1
                    LayoutGroup =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =41
                End
                Begin PageBreak
                    Top =12015
                    Name ="PageBreak654"
                End
                Begin EmptyCell
                    Left =6240
                    Top =8412
                    Width =3372
                    Height =300
                    Name ="EmptyCell669"
                    GroupTable =51
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =6240
                    LayoutCachedTop =8412
                    LayoutCachedWidth =9612
                    LayoutCachedHeight =8712
                    RowStart =5
                    RowEnd =5
                    ColumnStart =2
                    ColumnEnd =3
                    LayoutGroup =4
                    GroupTable =51
                End
                Begin EmptyCell
                    Left =6240
                    Top =10704
                    Width =3372
                    Height =300
                    Name ="EmptyCell670"
                    GroupTable =51
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =6240
                    LayoutCachedTop =10704
                    LayoutCachedWidth =9612
                    LayoutCachedHeight =11004
                    RowStart =12
                    RowEnd =12
                    ColumnStart =2
                    ColumnEnd =3
                    LayoutGroup =4
                    GroupTable =51
                End
                Begin EmptyCell
                    Left =6240
                    Top =11352
                    Width =3372
                    Height =300
                    Name ="EmptyCell671"
                    GroupTable =51
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =6240
                    LayoutCachedTop =11352
                    LayoutCachedWidth =9612
                    LayoutCachedHeight =11652
                    RowStart =14
                    RowEnd =14
                    ColumnStart =2
                    ColumnEnd =3
                    LayoutGroup =4
                    GroupTable =51
                End
                Begin EmptyCell
                    Left =6240
                    Top =9060
                    Width =3372
                    Height =300
                    Name ="EmptyCell672"
                    GroupTable =51
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =6240
                    LayoutCachedTop =9060
                    LayoutCachedWidth =9612
                    LayoutCachedHeight =9360
                    RowStart =7
                    RowEnd =7
                    ColumnStart =2
                    ColumnEnd =3
                    LayoutGroup =4
                    GroupTable =51
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    SpecialEffect =5
                    BorderWidth =1
                    TextAlign =1
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Left =6876
                    Top =7416
                    Width =2736
                    Height =348
                    TabIndex =15
                    Name ="Text676"
                    GroupTable =51
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =6876
                    LayoutCachedTop =7416
                    LayoutCachedWidth =9612
                    LayoutCachedHeight =7764
                    RowStart =2
                    RowEnd =2
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =4
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =51
                    Begin
                        Begin Label
                            BorderWidth =1
                            TextAlign =3
                            TextFontFamily =34
                            Left =6240
                            Top =7416
                            Width =630
                            Height =348
                            Name ="Label673"
                            Caption ="Date: "
                            GroupTable =51
                            LeftPadding =0
                            TopPadding =0
                            RightPadding =0
                            BottomPadding =0
                            LayoutCachedLeft =6240
                            LayoutCachedTop =7416
                            LayoutCachedWidth =6870
                            LayoutCachedHeight =7764
                            RowStart =2
                            RowEnd =2
                            ColumnStart =2
                            ColumnEnd =2
                            LayoutGroup =4
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =51
                        End
                    End
                End
                Begin EmptyCell
                    Left =6240
                    Top =7764
                    Width =630
                    Height =300
                    Name ="EmptyCell683"
                    GroupTable =51
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =6240
                    LayoutCachedTop =7764
                    LayoutCachedWidth =6870
                    LayoutCachedHeight =8064
                    RowStart =3
                    RowEnd =3
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =4
                    GroupTable =51
                End
                Begin EmptyCell
                    Left =6876
                    Top =7764
                    Width =2736
                    Height =300
                    Name ="EmptyCell686"
                    GroupTable =51
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =6876
                    LayoutCachedTop =7764
                    LayoutCachedWidth =9612
                    LayoutCachedHeight =8064
                    RowStart =3
                    RowEnd =3
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =4
                    GroupTable =51
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    BorderWidth =1
                    TextAlign =3
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Left =6240
                    Top =9708
                    Width =630
                    Height =348
                    TabIndex =22
                    Name ="Text687"
                    ControlSource ="=IIf(IsNull([subleasee]),Null,\"Date: \")"
                    GroupTable =51
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =6240
                    LayoutCachedTop =9708
                    LayoutCachedWidth =6870
                    LayoutCachedHeight =10056
                    RowStart =9
                    RowEnd =9
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =4
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =51
                End
                Begin EmptyCell
                    Left =6240
                    Top =10056
                    Width =3372
                    Height =300
                    Name ="EmptyCell688"
                    GroupTable =51
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =6240
                    LayoutCachedTop =10056
                    LayoutCachedWidth =9612
                    LayoutCachedHeight =10356
                    RowStart =10
                    RowEnd =10
                    ColumnStart =2
                    ColumnEnd =3
                    LayoutGroup =4
                    GroupTable =51
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    SpecialEffect =5
                    BorderWidth =1
                    TextAlign =1
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Left =6876
                    Top =9708
                    Width =2736
                    Height =348
                    TabIndex =23
                    Name ="Text692"
                    GroupTable =51
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =6876
                    LayoutCachedTop =9708
                    LayoutCachedWidth =9612
                    LayoutCachedHeight =10056
                    RowStart =9
                    RowEnd =9
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =4
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =51
                End
                Begin EmptyCell
                    Left =1080
                    Top =1044
                    Width =3435
                    Height =300
                    Name ="EmptyCell867"
                    GroupTable =49
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =1080
                    LayoutCachedTop =1044
                    LayoutCachedWidth =4515
                    LayoutCachedHeight =1344
                    RowStart =3
                    RowEnd =3
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =3
                    GroupTable =49
                End
                Begin EmptyCell
                    Top =1044
                    Width =1080
                    Height =300
                    Name ="EmptyCell868"
                    GroupTable =49
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedTop =1044
                    LayoutCachedWidth =1080
                    LayoutCachedHeight =1344
                    RowStart =3
                    RowEnd =3
                    LayoutGroup =3
                    GroupTable =49
                End
            End
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            RepeatSection = NotDefault
            Height =450
            BreakLevel =1
            Name ="GroupHeader1"
            Begin
                Begin TextBox
                    FontUnderline = NotDefault
                    TextAlign =2
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Left =2025
                    Width =8220
                    Height =375
                    FontSize =14
                    FontWeight =700
                    Name ="Label658"
                    ControlSource ="=IIf([Page]>\"2\",\"Exhibit A Equipment - Cont'd\",\"Exhibit A Equipment\")"

                    LayoutCachedLeft =2025
                    LayoutCachedWidth =10245
                    LayoutCachedHeight =375
                End
            End
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            RepeatSection = NotDefault
            Height =483
            BreakLevel =2
            Name ="GroupHeader0"
            Begin
                Begin Rectangle
                    Left =75
                    Width =11271
                    Height =483
                    BackColor =14277081
                    Name ="Box845"
                    LayoutCachedLeft =75
                    LayoutCachedWidth =11346
                    LayoutCachedHeight =483
                    BackShade =85.0
                End
                Begin TextBox
                    Visible = NotDefault
                    IMESentenceMode =3
                    Left =10590
                    Width =180
                    Height =0
                    TabIndex =1
                    Name ="grpcd"
                    ControlSource ="grpcd"

                    LayoutCachedLeft =10590
                    LayoutCachedWidth =10770
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =34
                    Left =8640
                    Width =1620
                    Height =480
                    FontSize =9
                    FontWeight =700
                    Name ="AcceptedDateLabel"
                    Caption ="Lessor's Cost Estimated Per Unit"
                    GroupTable =53
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =8640
                    LayoutCachedWidth =10260
                    LayoutCachedHeight =480
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =5
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =53
                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1092
                    Width =1968
                    Height =480
                    FontSize =9
                    FontWeight =700
                    Name ="Text699"
                    ControlSource ="=Chr(10) & IIf(IsNull([ChildVin]),\"VIN\",\"VIN / Serial Number\")"
                    GroupTable =53
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =1092
                    LayoutCachedWidth =3060
                    LayoutCachedHeight =480
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =5
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =53
                    Begin
                        Begin Label
                            TextAlign =1
                            TextFontFamily =34
                            Left =72
                            Width =1020
                            Height =480
                            FontSize =9
                            FontWeight =700
                            Name ="Label694"
                            Caption ="\015\012Unit No."
                            GroupTable =53
                            LeftPadding =0
                            TopPadding =0
                            RightPadding =0
                            BottomPadding =0
                            LayoutCachedLeft =72
                            LayoutCachedWidth =1092
                            LayoutCachedHeight =480
                            LayoutGroup =5
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =53
                        End
                    End
                End
                Begin EmptyCell
                    Left =11280
                    Width =0
                    Height =480
                    Name ="EmptyCell814"
                    GroupTable =53
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =11280
                    LayoutCachedWidth =11280
                    LayoutCachedHeight =480
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =5
                    GroupTable =53
                End
                Begin Label
                    TextAlign =1
                    TextFontFamily =34
                    Left =3060
                    Width =5580
                    Height =480
                    FontSize =9
                    FontWeight =700
                    Name ="Label848"
                    Caption ="\015\012 Year / Make  /  Model  / Description"
                    GroupTable =53
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =3060
                    LayoutCachedWidth =8640
                    LayoutCachedHeight =480
                    ColumnStart =2
                    ColumnEnd =3
                    LayoutGroup =5
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =53
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =34
                    Left =10260
                    Width =1020
                    Height =480
                    FontSize =9
                    FontWeight =700
                    Name ="Label887"
                    Caption ="Basic Rent Per Unit"
                    GroupTable =53
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =10260
                    LayoutCachedWidth =11280
                    LayoutCachedHeight =480
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =5
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =53
                End
            End
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =240
            BreakLevel =5
            Name ="GroupHeader3"
            Begin
                Begin TextBox
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1092
                    Width =1968
                    FontSize =9
                    TabIndex =1
                    Name ="Text179"
                    ControlSource ="=[UnitVin]"
                    GroupTable =53
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =1092
                    LayoutCachedWidth =3060
                    LayoutCachedHeight =240
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =5
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =53
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    TextFontCharSet =186
                    TextAlign =1
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1092
                    Top =240
                    Width =1968
                    Height =0
                    FontSize =9
                    TabIndex =7
                    Name ="Text180"
                    ControlSource ="=IIf(IsNull([ChildVin]),Null,[ChildVIN])"
                    GroupTable =53
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =1092
                    LayoutCachedTop =240
                    LayoutCachedWidth =3060
                    LayoutCachedHeight =240
                    RowStart =2
                    RowEnd =2
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =5
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =53
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Left =3060
                    Width =5580
                    FontSize =9
                    TabIndex =2
                    Name ="Text162"
                    ControlSource ="DescriptionLong"
                    GroupTable =53
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =3060
                    LayoutCachedWidth =8640
                    LayoutCachedHeight =240
                    RowStart =1
                    RowEnd =1
                    ColumnStart =2
                    ColumnEnd =3
                    LayoutGroup =5
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =53
                End
                Begin TextBox
                    Visible = NotDefault
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    TextAlign =1
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Left =72
                    Top =240
                    Width =1020
                    Height =0
                    FontSize =9
                    TabIndex =6
                    Name ="Text177"
                    ControlSource ="=IIf(IsNull([ChildVin]),Null,\"└───────\")"
                    GroupTable =53
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =72
                    LayoutCachedTop =240
                    LayoutCachedWidth =1092
                    LayoutCachedHeight =240
                    RowStart =2
                    RowEnd =2
                    LayoutGroup =5
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =53
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Left =3060
                    Top =240
                    Width =525
                    Height =0
                    FontSize =9
                    TabIndex =8
                    Name ="Text176"
                    ControlSource ="=IIf(IsNull([ChildVin]),Null,[ChildYear])"
                    GroupTable =53
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =3060
                    LayoutCachedTop =240
                    LayoutCachedWidth =3585
                    LayoutCachedHeight =240
                    RowStart =2
                    RowEnd =2
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =5
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =53
                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Left =8640
                    Width =1620
                    FontSize =9
                    TabIndex =3
                    Name ="TotalCost"
                    ControlSource ="TotalCost"
                    Format ="$#,##0.00;($#,##0.00)"
                    GroupTable =53
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =8640
                    LayoutCachedWidth =10260
                    LayoutCachedHeight =240
                    RowStart =1
                    RowEnd =1
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =5
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =53
                    CurrencySymbol ="$"
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    TextFontCharSet =186
                    TextAlign =1
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Left =3588
                    Top =240
                    Width =5052
                    Height =0
                    FontSize =9
                    TabIndex =9
                    Name ="Text173"
                    ControlSource ="=iif([ChildMakeModel]<>\"\",[ChildMakeModel] & \"  \" & [Ctype],null)"
                    GroupTable =53
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =3588
                    LayoutCachedTop =240
                    LayoutCachedWidth =8640
                    LayoutCachedHeight =240
                    RowStart =2
                    RowEnd =2
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =5
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =53
                End
                Begin TextBox
                    TextFontCharSet =186
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Left =72
                    Width =1020
                    FontSize =9
                    Name ="Text161"
                    ControlSource ="[UnitUnitNum]"
                    GroupTable =53
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =72
                    LayoutCachedWidth =1092
                    LayoutCachedHeight =240
                    RowStart =1
                    RowEnd =1
                    LayoutGroup =5
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =53
                End
                Begin TextBox
                    Visible = NotDefault
                    TextAlign =1
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Left =11280
                    Width =0
                    ColumnOrder =4
                    FontSize =9
                    TabIndex =5
                    Name ="UnitID"
                    ControlSource ="UnitID"
                    GroupTable =53
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =11280
                    LayoutCachedWidth =11280
                    LayoutCachedHeight =240
                    RowStart =1
                    RowEnd =1
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =5
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =53
                End
                Begin EmptyCell
                    Left =8640
                    Top =240
                    Width =1620
                    Height =0
                    Name ="EmptyCell835"
                    GroupTable =53
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =8640
                    LayoutCachedTop =240
                    LayoutCachedWidth =10260
                    LayoutCachedHeight =240
                    RowStart =2
                    RowEnd =2
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =5
                    GroupTable =53
                End
                Begin EmptyCell
                    Left =11280
                    Top =240
                    Width =0
                    Height =0
                    Name ="EmptyCell837"
                    GroupTable =53
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =11280
                    LayoutCachedTop =240
                    LayoutCachedWidth =11280
                    LayoutCachedHeight =240
                    RowStart =2
                    RowEnd =2
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =5
                    GroupTable =53
                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Left =10260
                    Width =1020
                    FontSize =9
                    TabIndex =4
                    Name ="UNITRENT"
                    ControlSource ="Rentx"
                    Format ="$#,##0.00;($#,##0.00)"
                    GroupTable =53
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =10260
                    LayoutCachedWidth =11280
                    LayoutCachedHeight =240
                    RowStart =1
                    RowEnd =1
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =5
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =53
                    CurrencySymbol ="$"
                End
                Begin EmptyCell
                    Left =10260
                    Top =240
                    Width =1020
                    Height =0
                    Name ="EmptyCell893"
                    GroupTable =53
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =10260
                    LayoutCachedTop =240
                    LayoutCachedWidth =11280
                    LayoutCachedHeight =240
                    RowStart =2
                    RowEnd =2
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =5
                    GroupTable =53
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =0
            Name ="Detail"
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =675
            BreakLevel =2
            Name ="GroupFooter0"
            Begin
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    TextAlign =1
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Left =4740
                    Top =312
                    Width =2505
                    FontSize =9
                    TabIndex =4
                    Name ="Text111"
                    ControlSource ="=[dcnty] & \" (\" & [TaxType] & \")\""
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =4740
                    LayoutCachedTop =312
                    LayoutCachedWidth =7245
                    LayoutCachedHeight =552
                    RowStart =2
                    RowEnd =2
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                    Begin
                        Begin Label
                            TextAlign =1
                            TextFontFamily =34
                            Left =3048
                            Top =312
                            Width =1692
                            Height =240
                            FontSize =9
                            FontWeight =700
                            Name ="Label150"
                            Caption ="County for Sale/Tax: "
                            GroupTable =1
                            LeftPadding =0
                            TopPadding =0
                            RightPadding =0
                            BottomPadding =0
                            LayoutCachedLeft =3048
                            LayoutCachedTop =312
                            LayoutCachedWidth =4740
                            LayoutCachedHeight =552
                            RowStart =2
                            RowEnd =2
                            ColumnStart =1
                            ColumnEnd =1
                            LayoutGroup =1
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =1
                        End
                    End
                End
                Begin Label
                    TextAlign =1
                    TextFontFamily =34
                    Left =3048
                    Top =72
                    Width =4197
                    Height =240
                    FontSize =9
                    FontWeight =700
                    Name ="Label113"
                    Caption ="Equipment to be garaged or used at:"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =3048
                    LayoutCachedTop =72
                    LayoutCachedWidth =7245
                    LayoutCachedHeight =312
                    ColumnStart =1
                    ColumnEnd =2
                    LayoutGroup =1
                    BackThemeColorIndex =1
                    BackShade =85.0
                    BorderThemeColorIndex =1
                    BorderShade =65.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin Label
                    TextAlign =1
                    TextFontFamily =34
                    Left =7248
                    Top =72
                    Width =4116
                    Height =240
                    FontSize =9
                    FontWeight =700
                    Name ="Label117"
                    Caption ="Equipment is registered to:"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =7248
                    LayoutCachedTop =72
                    LayoutCachedWidth =11364
                    LayoutCachedHeight =312
                    ColumnStart =3
                    ColumnEnd =4
                    LayoutGroup =1
                    BackThemeColorIndex =1
                    BackShade =85.0
                    BorderThemeColorIndex =1
                    BorderShade =65.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin Label
                    TextAlign =1
                    TextFontFamily =34
                    Left =72
                    Top =72
                    Width =2976
                    Height =240
                    FontSize =9
                    FontWeight =700
                    Name ="Label122"
                    Caption ="Equipment to be delivered to:"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =72
                    LayoutCachedTop =72
                    LayoutCachedWidth =3048
                    LayoutCachedHeight =312
                    LayoutGroup =1
                    BackThemeColorIndex =1
                    BackShade =85.0
                    BorderThemeColorIndex =1
                    BorderShade =65.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin TextBox
                    TextAlign =1
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Left =9216
                    Top =312
                    Width =2148
                    FontSize =9
                    TabIndex =5
                    Name ="TSt"
                    ControlSource ="TitledState"
                    StatusBarText ="lookup in lookup_state table"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =9216
                    LayoutCachedTop =312
                    LayoutCachedWidth =11364
                    LayoutCachedHeight =552
                    RowStart =2
                    RowEnd =2
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                    Begin
                        Begin Label
                            TextAlign =1
                            TextFontFamily =34
                            Left =7248
                            Top =312
                            Width =1968
                            Height =240
                            FontSize =9
                            FontWeight =700
                            Name ="Label114"
                            Caption ="Equipment is titled in:"
                            GroupTable =1
                            LeftPadding =0
                            TopPadding =0
                            RightPadding =0
                            BottomPadding =0
                            LayoutCachedLeft =7248
                            LayoutCachedTop =312
                            LayoutCachedWidth =9216
                            LayoutCachedHeight =552
                            RowStart =2
                            RowEnd =2
                            ColumnStart =3
                            ColumnEnd =3
                            LayoutGroup =1
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =1
                        End
                    End
                End
                Begin TextBox
                    IMESentenceMode =3
                    Top =15
                    Width =10800
                    Height =30
                    Name ="Text151"
                    ControlSource ="clientCompanyName"

                    LayoutCachedTop =15
                    LayoutCachedWidth =10800
                    LayoutCachedHeight =45
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    TextAlign =1
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =3048
                    Top =312
                    Width =4197
                    Height =0
                    FontSize =9
                    TabIndex =2
                    Name ="Text157"
                    ControlSource ="GarageLoc"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =3048
                    LayoutCachedTop =312
                    LayoutCachedWidth =7245
                    LayoutCachedHeight =312
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =2
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    TextAlign =1
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =7248
                    Top =312
                    Width =4116
                    Height =0
                    FontSize =9
                    TabIndex =3
                    Name ="Text159"
                    ControlSource ="RegLoc"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =7248
                    LayoutCachedTop =312
                    LayoutCachedWidth =11364
                    LayoutCachedHeight =312
                    RowStart =1
                    RowEnd =1
                    ColumnStart =3
                    ColumnEnd =4
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    TextAlign =1
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =72
                    Top =312
                    Width =2976
                    Height =0
                    FontSize =9
                    TabIndex =1
                    Name ="Text160"
                    ControlSource ="DeliveryLoc"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =72
                    LayoutCachedTop =312
                    LayoutCachedWidth =3048
                    LayoutCachedHeight =312
                    RowStart =1
                    RowEnd =1
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =72
                    Top =312
                    Width =2976
                    Name ="EmptyCell204"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =72
                    LayoutCachedTop =312
                    LayoutCachedWidth =3048
                    LayoutCachedHeight =552
                    RowStart =2
                    RowEnd =2
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Rectangle
                    Left =75
                    Top =615
                    Width =11181
                    Height =0
                    BackColor =14277081
                    Name ="Box844"
                    LayoutCachedLeft =75
                    LayoutCachedTop =615
                    LayoutCachedWidth =11256
                    LayoutCachedHeight =615
                    BackShade =85.0
                End
            End
        End
        Begin PageFooter
            Height =240
            Name ="PageFooterSection"
            Begin
                Begin TextBox
                    TextFontCharSet =163
                    TextAlign =3
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Left =7320
                    Width =3600
                    ColumnOrder =0
                    FontSize =9
                    Name ="Text66"
                    ControlSource ="=\"Page \" & [Page] & \" of \" & [Pages]"

                    LayoutCachedLeft =7320
                    LayoutCachedWidth =10920
                    LayoutCachedHeight =240
                End
            End
        End
        Begin FormFooter
            KeepTogether = NotDefault
            CanGrow = NotDefault
            Height =0
            Name ="ReportFooter"
        End
    End
End
CodeBehindForm
' See "CofA_Mass_SPD.cls"
