SELECT
  vw_SixKeys.clientID,
  vw_SixKeys.SchID,
  LRemit.locationId AS RemitAddressID,
  vw_SixKeys.Schedule,
  vw_SixKeys.MLNo,
  Format([MLDt], "mmmm d"", ""yyyy") AS MLDate,
  Clients.clientCompanyName AS Lessee,
  [clientCompanyName] & Chr(13)& Chr(10)& IIf(
    IsNull([clientCoNm2]),
    "",
    [clientCoNm2] & Chr(13)& Chr(10)
  )& [clientAddress1] & " " & [clientAddress2] & Chr(13)& Chr(10)& [clientcity] & ", " & [clientstate] & " " & [clientzip] AS LesseeAddress,
  Banks.BankName,
  [BankName] & Chr(13)& Chr(10)& [Address1] & Chr(13)& Chr(10)& IIf(
    Len([banks].[Address2])> 3,
    [Address2] & Chr(13)& Chr(10),
    ""
  )& [Banks].[City] & ", " & [banks].[State] & " " & [banks].[Zip] AS BanksAddress,
  [BankName] & ", " & [Address1] & ", " & IIf(
    Len([banks].[Address2])> 3,
    "Mail Stop " & [Address2] & ", ",
    ""
  )& [Banks].[City] & ", " & [banks].[State] & " " & [banks].[Zip] AS BankAddressInLine,
  [LRemit].[LocationStreet] & Chr(13)& Chr(10)& [LRemit].[LAddrLine1] & Chr(13)& Chr(10)& IIf(
    Len([LRemit].[LAddrLine2])> 3,
    [LRemit].[LAddrLine2] & Chr(13)& Chr(10),
    ""
  )& [LRemit].[LCity] & " " & [LRemit].[LState] & " " & Left([LRemit].[LZip], 5)& IIf(
    Len([lremit].[lzip])> 5,
    "-" & Right([lremit].[Lzip], 4),
    ""
  ) AS BankRemitAddress,
  GroupsNoinSchedule([vw_sixKeys].[schid]) AS GroupsIn,
  ScheduleGroupsInLine([vw_sixKeys].[schid]) AS ConsistOf,
  Format([SchedDt], "mmmm d"", ""yyyy") AS ScheduleDate
FROM
  (
    (
      (
        (
          (
            (
              vw_SixKeys
              INNER JOIN Units ON vw_SixKeys.UnitID = Units.UnitID
            )
            INNER JOIN SchGrp ON vw_SixKeys.SGrpID = SchGrp.SGrpID
          )
          INNER JOIN Schedule ON vw_SixKeys.SchID = Schedule.SchID
        )
        LEFT JOIN Banks ON vw_SixKeys.BankId = Banks.BankId
      )
      LEFT JOIN LocationsRemit AS LRemit ON Banks.BankId = LRemit.ClientREF
    )
    INNER JOIN Clients ON vw_SixKeys.clientID = Clients.clientID
  )
  INNER JOIN MstrLease ON vw_SixKeys.MLID = MstrLease.MLID
GROUP BY
  vw_SixKeys.clientID,
  vw_SixKeys.SchID,
  LRemit.locationId,
  vw_SixKeys.Schedule,
  vw_SixKeys.MLNo,
  Format([MLDt], "mmmm d"", ""yyyy"),
  Clients.clientCompanyName,
  [clientCompanyName] & Chr(13)& Chr(10)& IIf(
    IsNull([clientCoNm2]),
    "",
    [clientCoNm2] & Chr(13)& Chr(10)
  )& [clientAddress1] & " " & [clientAddress2] & Chr(13)& Chr(10)& [clientcity] & ", " & [clientstate] & " " & [clientzip],
  Banks.BankName,
  [BankName] & Chr(13)& Chr(10)& [Address1] & Chr(13)& Chr(10)& IIf(
    Len([banks].[Address2])> 3,
    [Address2] & Chr(13)& Chr(10),
    ""
  )& [Banks].[City] & ", " & [banks].[State] & " " & [banks].[Zip],
  [BankName] & ", " & [Address1] & ", " & IIf(
    Len([banks].[Address2])> 3,
    "Mail Stop " & [Address2] & ", ",
    ""
  )& [Banks].[City] & ", " & [banks].[State] & " " & [banks].[Zip],
  [LRemit].[LocationStreet] & Chr(13)& Chr(10)& [LRemit].[LAddrLine1] & Chr(13)& Chr(10)& IIf(
    Len([LRemit].[LAddrLine2])> 3,
    [LRemit].[LAddrLine2] & Chr(13)& Chr(10),
    ""
  )& [LRemit].[LCity] & " " & [LRemit].[LState] & " " & Left([LRemit].[LZip], 5)& IIf(
    Len([lremit].[lzip])> 5,
    "-" & Right([lremit].[Lzip], 4),
    ""
  ),
  Format([SchedDt], "mmmm d"", ""yyyy")
HAVING
  (
    (
      (vw_SixKeys.SchID)= [Forms]![ScheduleControl]![SchID]
    )
  );
