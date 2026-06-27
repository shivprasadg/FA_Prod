SELECT
  vw_SixKeys.CompanyName AS Lessee,
  vw_SixKeys.MLNo AS [Leasee Master Lease Number],
  vw_SixKeys.Schedule AS [Leasee Schedule No],
  "Equipment File" AS [File Type],
  Date() AS DateSent,
  Units.UnitTitleNumber AS TitleNum,
  Locations.LState AS [State of Title],
  vw_SixKeys.VIN,
  [UnitYr] & "  " & [UnitMake] & " " & [unitmodel] & " " & [UnitType] & " " & [Axle] AS [Equipment Description],
  vw_SixKeys.PortfolioNumber AS [Portfolio Certificate No],
  Banks.BankName AS [Lender or Holder Name],
  IIf(
    [MLLiabIns] = "R"
    And [MlinsExp] > Date(),
    "Yes",
    "Yes"
  ) AS [Copy of Certificate of Insurance included],
  "N/A" AS [Original Bill of Sale Included],
  vw_SixKeys.AssetID AS [FA-Asset Num],
  Banks.BankId
FROM
  Schedule_Group_WorkTable
  INNER JOIN (
    (
      (
        (
          (
            vw_SixKeys
            LEFT JOIN TitleSentTrustDate ON vw_SixKeys.AssetID = TitleSentTrustDate.AssetId
          )
          INNER JOIN Units ON vw_SixKeys.AssetID = Units.UnitID
        )
        INNER JOIN Locations ON Units.UnitTitleLocationRef = Locations.LocationID
      )
      INNER JOIN Banks ON vw_SixKeys.BankId = Banks.BankId
    )
    INNER JOIN MstrLease ON vw_SixKeys.MLID = MstrLease.MLID
  ) ON Schedule_Group_WorkTable.ScheduleId = vw_SixKeys.SchID
WHERE
  (
    (
      (vw_SixKeys.PortfolioNumber) Is Not Null
    )
    AND (
      (vw_SixKeys.clientID) Is Not Null
    )
  )
ORDER BY
  vw_SixKeys.CompanyName,
  vw_SixKeys.Schedule;
