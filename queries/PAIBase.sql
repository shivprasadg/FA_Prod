SELECT
  Clients.clientCompanyName,
  LRemit.locationId AS RemitAddressID,
  MstrLease.MLNo,
  Schedule.SchNo AS Schedule,
  SchGrp.SGrpNO AS GroupNo,
  SchGrp.UnitsInGroup,
  Format([MLDt], "mmmm d"", ""yyyy") AS MLDate,
  Clients.clientCompanyName AS Lessee,
  [clientCompanyName] & Chr(13)& Chr(10)& IIf(
    IsNull([clientCoNm2]),
    "",
    [clientCoNm2] & Chr(13)& Chr(10)
  )& [clientAddress1] & " " & [clientAddress2] & Chr(13)& Chr(10)& [clientcity] & ", " & [clientstate] & " " & [clientzip] AS LesseeAddress,
  SchGrp.LeasePymtForGroup AS GroupRent,
  DatePart("d", [schgrp.BLCD])& IIf(
    DatePart("d", [schgrp.BLCD])= 1,
    "st",
    IIf(
      DatePart("d", [schgrp.BLCD])= 31,
      "st",
      IIf(
        CStr(
          DatePart("d", [schgrp.BLCD])
        ) Like '*2',
        "nd",
        IIf(
          DatePart("d", [schgrp.BLCD])= 3,
          "rd",
          "th"
        )
      )
    )
  )& " of each month [" & [schgrp].[blcd] & " ~ " & IIf(
    IsNull([SchGrp.BLCD]),
    "tbd",
    IIf(
      [LeasePaidAdvArr] = "ADV",
      DateAdd(
        "m", [schgrp].[LeaseTermPrimary] - 1,
        [schgrp.BLCD]
      ),
      DateAdd(
        "m", [schgrp].[LeaseTermPrimary],
        [schgrp.BLCD]
      )
    )
  )& "]" AS RentDueDates,
  DatePart("d", [schgrp.BLCD])& IIf(
    DatePart("d", [schgrp.BLCD])= 1,
    "st",
    IIf(
      DatePart("d", [schgrp.BLCD])= 31,
      "st",
      IIf(
        CStr(
          DatePart("d", [schgrp.BLCD])
        ) Like '*2',
        "nd",
        IIf(
          DatePart("d", [schgrp.BLCD])= 3,
          "rd",
          "th"
        )
      )
    )
  ) AS PayDate,
  Banks.BankShortName,
  BankPAI.ABANumber,
  BankPAI.AccountName,
  BankPAI.AccountNumber,
  BankPAI.BankName,
  BankPAI.BankAddress,
  [LRemit].[LocationStreet] & Chr(13)& Chr(10)& [LRemit].[LAddrLine1] & Chr(13)& Chr(10)& IIf(
    Len([LRemit].[LAddrLine2])> 3,
    [LRemit].[LAddrLine2] & Chr(13)& Chr(10),
    ""
  )& [LRemit].[LCity] & " " & [LRemit].[LState] & " " & Left([LRemit].[LZip], 5)& IIf(
    Len([lremit].[lzip])> 5,
    "-" & Right([lremit].[Lzip], 4),
    ""
  ) AS BankRemitAddress,
  IIf(
    Nz([SchBnkClNO], "")<> "",
    [SchBnkClNO] & "-",
    ""
  )& [SgrpBnkNo] AS Contract,
  GroupsNoinSchedule(
    [schedule].[schid], 0, False, True
  ) AS GroupsIn,
  Format([SchedDt], "mmmm d"", ""yyyy") AS ScheduleDate,
  Clients.clientID,
  Schedule.SchID,
  SchGrp.SgrpId,
  Banks.BankID
FROM
  (
    (
      (
        Clients
        INNER JOIN MstrLease ON Clients.clientID = MstrLease.Client
      )
      INNER JOIN Schedule ON MstrLease.MLID = Schedule.SchMLRef
    )
    INNER JOIN (
      SchGrp
      INNER JOIN (
        Banks
        LEFT JOIN LocationsRemit AS LRemit ON Banks.BankId = LRemit.ClientREF
      ) ON SchGrp.AssigneeBankID = Banks.BankID
    ) ON Schedule.SchID = SchGrp.SGrpSchID
  )
  INNER JOIN BankPAI ON Banks.BankID = BankPAI.BankID
WHERE
  (
    (
      (SchGrp.Assigned_Indicator)= "F"
    )
    AND (
      (SchGrp.blcd) Is Not Null
    )
  )
ORDER BY
  SchGrp.SGrpNO;
