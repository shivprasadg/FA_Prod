SELECT
  Banks.BankId,
  Units.DrawDownId,
  vw_SixKeys.clientID,
  MstrLease.MLGuarantor,
  DrawDowns.DrawNumber,
  Banks.BankName,
  Clients.clientCompanyName AS Guaranty,
  vw_SixKeys.CompanyName,
  vw_SixKeys.Schedule,
  vw_SixKeys.UnitGroup,
  Count(vw_SixKeys.AssetID) AS QTY,
  vw_SixKeys.PrintedDocs AS UnitDesc,
  Units.UnitIfAmt AS UnitCost,
  Sum(Units.UnitIfAmt) AS TotalCost,
  LastMilesReportedAll.MPYAllowed,
  vw_SixKeys.LeaseTermPrimary,
  Format([UnitEstDelDate], "mmm-yyyy") AS EstDelivery,
  Units.UnitsRent,
  Sum(Units.[UnitsRent]) AS RentTotal
FROM
  (
    (
      (
        (
          (
            DrawDowns
            INNER JOIN vw_SixKeys ON DrawDowns.DrawDownId = vw_SixKeys.DrawDownId
          )
          INNER JOIN Units ON vw_SixKeys.UnitID = Units.UnitID
        )
        LEFT JOIN LastMilesReportedAll ON Units.UnitID = LastMilesReportedAll.UnitID
      )
      INNER JOIN Banks ON DrawDowns.DrawBankId = Banks.BankId
    )
    INNER JOIN MstrLease ON DrawDowns.MLID = MstrLease.MLID
  )
  LEFT JOIN Clients ON MstrLease.MLGuarantor = Clients.clientID
GROUP BY
  Banks.BankId,
  Units.DrawDownId,
  vw_SixKeys.clientID,
  MstrLease.MLGuarantor,
  DrawDowns.DrawNumber,
  Banks.BankName,
  Clients.clientCompanyName,
  vw_SixKeys.CompanyName,
  vw_SixKeys.Schedule,
  vw_SixKeys.UnitGroup,
  vw_SixKeys.PrintedDocs,
  Units.UnitIfAmt,
  LastMilesReportedAll.MPYAllowed,
  vw_SixKeys.LeaseTermPrimary,
  Format([UnitEstDelDate], "mmm-yyyy"),
  Units.UnitsRent
HAVING
  (
    (
      (Units.DrawDownId)= [Forms]![DrawDownManager]![DrawID]
    )
  )
ORDER BY
  vw_SixKeys.UnitGroup,
  Format([UnitEstDelDate], "mmm-yyyy");
