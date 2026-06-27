SELECT
  CofA_Report_X.*,
  Locations.LocationName AS TitleLocName,
  [Locations].[LocationStreet] & IIf(
    IsNull([Locations].[LocSubLessee]),
    Null,
    Chr(13)& Chr(10)& [Locations].[LocSubLessee]
  )& Chr(13)& Chr(10)& [Locations].[LAddrLine1] & Chr(13)& Chr(10)& [Locations].[LCity] & "  " & [Locations].[LState] & " " & IIf(
    Len([Locations].[LZip])> 5,
    Left([Locations].[LZip], 5)& "-" & Right([Locations].[LZip], 4),
    [Locations].[LZip]
  ) AS TitleLoc,
  [Locations_1].[LocationStreet] & IIf(
    IsNull([Locations_1].[LocSubLessee]),
    Null,
    Chr(13)& Chr(10)& [Locations_1].[LocSubLessee]
  )& Chr(13)& Chr(10)& [Locations_1].[LAddrLine1] & Chr(13)& Chr(10)& [Locations_1].[LCity] & "  " & [Locations_1].[LState] & " " & IIf(
    Len([Locations_1].[LZip])> 5,
    Left([Locations_1].[LZip], 5)& "-" & Right([Locations_1].[LZip], 4),
    [Locations_1].[LZip]
  ) AS GarageLoc,
  [Locations_2].[LocationStreet] & IIf(
    IsNull([Locations_2].[LocSubLessee]),
    Null,
    Chr(13)& Chr(10)& [Locations_2].[LocSubLessee]
  )& Chr(13)& Chr(10)& [Locations_2].[LAddrLine1] & Chr(13)& Chr(10)& [Locations_2].[LCity] & "  " & [Locations_2].[LState] & " " & IIf(
    Len([Locations_2].[LZip])> 5,
    Left([Locations_2].[LZip], 5)& "-" & Right([Locations_2].[LZip], 4),
    [Locations_2].[LZip]
  ) AS RegLoc,
  [Locations_3].[LocationStreet] & IIf(
    IsNull([Locations_3].[LocSubLessee]),
    Null,
    Chr(13)& Chr(10)& [Locations_3].[LocSubLessee]
  )& Chr(13)& Chr(10)& [Locations_3].[LAddrLine1] & Chr(13)& Chr(10)& [Locations_3].[LCity] & "  " & [Locations_3].[LState] & " " & IIf(
    Len([Locations_3].[LZip])> 5,
    Left([Locations_3].[LZip], 5)& "-" & Right([Locations_3].[LZip], 4),
    [Locations_3].[LZip]
  ) AS DeliveryLoc,
  Locations.LState AS TitledState,
  Locations.LState AS TSt,
  Locations_1.LocSubLessee AS DSub,
  Locations_1.LocationStreet AS Dst1,
  Locations_1.LAddrLine1 AS DST2,
  Locations_1.LCity AS DCITY,
  IIf(
    [Locations_1].[LCounty] Like "*county*",
    SplitString(
      [Locations_1].[LCounty], " County",
      0
    ),
    [Locations_1].[LCounty]
  ) AS DCNTY,
  Locations_1.LState AS DST,
  Locations_1.LZip AS DzIP,
  lookup_States.StTaxType,
  IIf(
    [UnitTaxExempt] = Yes, "Exempt", [StTaxType]
  ) AS TaxType,
  Locations_2.LocSubLessee AS RSub,
  Locations_2.LocationStreet AS Rst1,
  Locations_2.LAddrLine1 AS Rst2,
  Locations_2.LCity AS RCity,
  Locations_2.LState AS RST,
  Locations_2.LZip AS RZIP,
  Locations_3.LocSubLessee AS DelSub,
  Locations_3.LocationStreet AS DelSt1,
  Locations_3.LAddrLine1 AS DelSt2,
  Locations_3.LCity AS DelCity,
  Locations_3.LState AS DelSt,
  Locations_3.LZip AS DelZip,
  CLASS.TITLED,
  ScheduleAmends.AmendmentNO,
  ScheduleAmends.AmendmentDate,
  Contacts.contactTitle,
  [contactFirst] & " " & [contactMiddle] & IIf(
    IsNull([contactMiddle]),
    Null,
    "."
  )& " " & [Contactlast] & " " & [contactsuffix] AS FullNm,
  "n/a" AS MLRMo,
  "n/a" AS MLRday,
  "n/a" AS MLRYr,
  "n/a" AS Shmo,
  "n/a" AS Schday,
  "n/a" AS SchYr,
  Locations_1.LTaxRate AS GarLOC,
  Locations_1.LCountry,
  Locations_1.LTaxRate,
  ContactsIncumbent.ClientSignTitle,
  ContactsIncumbent.ClientSignFullName,
  [Desc] & "   " & IIf(
    [sgrpplb] = False, Null, "  Mi/Hr "
  )& IIf([sgrpplb] = False, Null, [mileage])& IIf([sgrpplb] = False, Null, " / ")& IIf([sgrpplb] = False, Null, [hours])& "   " & [UnitVIN] AS UnitDesc,
  DLookUp(
    "StateLong",
    "Lookup_States",
    "StateAbbrev='" & IIf(
      [Locations_2].[LocationName] Like "IRP Reg*",
      Trim(
        SplitString(
          [Locations_2].[LocationName], "IRP Reg -",
          1
        )
      ),
      [Locations_2].[LState]
    )& "'"
  ) AS RegState,
  DateAdd(
    "m", [LeaseTermPrimary], [BLCD] - 1
  ) AS ProjEOL
FROM
  (
    Bulk_Update_WorkTable
    INNER JOIN (
      (
        (
          (
            (
              (
                (
                  (
                    CofA_Report_X
                    LEFT JOIN Locations ON CofA_Report_X.UnitTitleLocationRef = Locations.LocationID
                  )
                  LEFT JOIN Locations AS Locations_1 ON CofA_Report_X.UnitGarageLocationRef = Locations_1.LocationID
                )
                LEFT JOIN lookup_States ON Locations_1.LState = lookup_States.StateAbbrev
              )
              LEFT JOIN Locations AS Locations_2 ON CofA_Report_X.UnitRegLocationRef = Locations_2.LocationID
            )
            LEFT JOIN Locations AS Locations_3 ON CofA_Report_X.UnitDeliveryLocationRef = Locations_3.LocationID
          )
          LEFT JOIN (
            lookup_groupTypes
            LEFT JOIN CLASS ON lookup_groupTypes.grpClass = CLASS.CLASS
          ) ON CofA_Report_X.UnitType = lookup_groupTypes.groupType
        )
        LEFT JOIN Contacts ON CofA_Report_X.CofASigner = Contacts.contactID
      )
      LEFT JOIN ScheduleAmends ON CofA_Report_X.SchID = ScheduleAmends.ScheduleID
    ) ON Bulk_Update_WorkTable.UnitID = CofA_Report_X.UnitID
  )
  LEFT JOIN ContactsIncumbent ON CofA_Report_X.CofASigner = ContactsIncumbent.contactID
ORDER BY
  CofA_Report_X.vin4,
  Locations_3.LAddrLine1;
