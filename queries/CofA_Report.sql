SELECT
  CofA_Report_A.*,
  Locations.LState AS TSt,
  Locations_1.LocSubLessee AS DSub,
  Locations_1.LocationStreet AS Dst1,
  Locations_1.LAddrLine1 AS DST2,
  Locations_1.LCity AS DCITY,
  Locations_1.LCounty AS DCNTY,
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
  MLAmends.MLAmtNO,
  MLAmends.MLAmdDt,
  ScheduleAmends.AmendmentNO,
  ScheduleAmends.AmendmentDate,
  Contacts.contactTitle,
  [contactFirst] & " " & [contactMiddle] & IIf(
    IsNull([contactMiddle]),
    Null,
    "."
  )& " " & [Contactlast] & " " & [contactsuffix] AS FullNm,
  IIf(
    IsNull([MLAmdDt]),
    "N/A",
    MonthName(
      DatePart("m", [MLAmdDt])
    )
  ) AS MLRMo,
  IIf(
    IsNull([MLAmdDt]),
    "N/A",
    MonthName(
      DatePart("d", [MLAmdDt])
    )
  ) AS MLRday,
  IIf(
    IsNull([MLAmdDt]),
    "N/A",
    MonthName(
      DatePart("yyyy", [MLAmdDt])
    )
  ) AS MLRYr,
  IIf(
    IsNull([scheddt]),
    "N/A",
    MonthName(
      DatePart("mm", [scheddt])
    )
  ) AS Shmo,
  IIf(
    IsNull([scheddt]),
    "N/A",
    MonthName(
      DatePart("dd", [scheddt])
    )
  ) AS Schday,
  IIf(
    IsNull([scheddt]),
    "N/A",
    MonthName(
      DatePart("yyyy", [scheddt])
    )
  ) AS SchYr,
  Locations_1.LTaxRate AS GarLOC,
  Locations_1.LCountry,
  Locations_1.LTaxRate
FROM
  (
    (
      (
        (
          (
            (
              (
                (
                  CofA_Report_A
                  LEFT JOIN Locations ON CofA_Report_A.UnitTitleLocationRef = Locations.LocationID
                )
                LEFT JOIN Locations AS Locations_1 ON CofA_Report_A.UnitGarageLocationRef = Locations_1.LocationID
              )
              LEFT JOIN lookup_States ON Locations_1.LState = lookup_States.StateAbbrev
            )
            LEFT JOIN Locations AS Locations_2 ON CofA_Report_A.UnitRegLocationRef = Locations_2.LocationID
          )
          LEFT JOIN Locations AS Locations_3 ON CofA_Report_A.UnitDeliveryLocationRef = Locations_3.LocationID
        )
        LEFT JOIN (
          lookup_groupTypes
          LEFT JOIN CLASS ON lookup_groupTypes.grpClass = CLASS.CLASS
        ) ON CofA_Report_A.UnitType = lookup_groupTypes.groupType
      )
      LEFT JOIN Contacts ON CofA_Report_A.CofASigner = Contacts.contactID
    )
    LEFT JOIN MLAmends ON CofA_Report_A.MLID = MLAmends.MLRef
  )
  LEFT JOIN ScheduleAmends ON CofA_Report_A.SchID = ScheduleAmends.ScheduleID;
