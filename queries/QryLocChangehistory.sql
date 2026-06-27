SELECT
  Clients.clientCompanyName,
  MstrLease.MLNo,
  MstrLease.MLOrig,
  Schedule.SchNo,
  SchGrp.SGrpNO,
  Units.Unitvin,
  Units.unityr,
  Units.unitmake,
  Units.unitmodel,
  Units.unittype,
  First(
    qry_FinalLocHist.FirstOfFirstCHdt1
  ) AS FirstOfFirstOfFirstCHdt1,
  First(
    qry_FinalLocHist.LastOfLastCHdt2
  ) AS FirstOfLastOfLastCHdt2,
  First(
    qry_FinalLocHist.FirstOfFirstOfOldUnitloc
  ) AS FirstOfFirstOfFirstOfOldUnitloc,
  qry_FinalLocHist.FirstOfFirstOldStreet,
  First(
    qry_FinalLocHist.FirstOfFirstOldCity
  ) AS FirstOfFirstOfFirstOldCity,
  First(
    qry_FinalLocHist.FirstOfFirstOldCounty
  ) AS FirstOfFirstOfFirstOldCounty,
  First(
    qry_FinalLocHist.FirstOfFirstOldState
  ) AS FirstOfFirstOfFirstOldState,
  First(
    qry_FinalLocHist.FirstOfFirstOldZip
  ) AS FirstOfFirstOfFirstOldZip,
  First(
    qry_FinalLocHist.FirstOfFirstOldTaxStatus
  ) AS FirstOfFirstOfFirstOldTaxStatus,
  First(
    qry_FinalLocHist.FirstOfFirstoldTaxNote
  ) AS FirstOfFirstOfFirstoldTaxNote,
  First(
    qry_FinalLocHist.LastOfLastOfUnitsFormloc
  ) AS FirstOfLastOfLastOfUnitsFormloc,
  First(
    qry_FinalLocHist.LastOfLastNewStreet
  ) AS FirstOfLastOfLastNewStreet,
  First(
    qry_FinalLocHist.LastOfLastNewTaxStatus
  ) AS FirstOfLastOfLastNewTaxStatus,
  First(
    qry_FinalLocHist.LastOfLastNewTaxNote
  ) AS FirstOfLastOfLastNewTaxNote,
  First(
    qry_FinalLocHist.LastOfNewLocation
  ) AS FirstOfLastOfNewLocation,
  First(
    qry_FinalLocHist.LastOfLastNewCounty
  ) AS FirstOfLastOfLastNewCounty,
  First(
    qry_FinalLocHist.LastOfLastNewState
  ) AS FirstOfLastOfLastNewState,
  First(
    qry_FinalLocHist.LastOfLastNewZip
  ) AS FirstOfLastOfLastNewZip,
  First(
    qry_FinalLocHist.LastOfMinRptChDate
  ) AS FirstOfLastOfMinRptChDate,
  First(
    qry_FinalLocHist.LastOfMaxRptChDate
  ) AS FirstOfLastOfMaxRptChDate,
  First(Clients_1.clientCompanyName) AS [Assignee Bank]
FROM
  (
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
        LEFT JOIN Clients AS Clients_1 ON SchGrp.Assignee_Bank = Clients_1.ClientShNm
      ) ON Schedule.SchID = SchGrp.SGrpSchID
    )
    INNER JOIN Units ON SchGrp.SGrpID = Units.GroupREF
  )
  INNER JOIN qry_FinalLocHist ON Units.UnitID = qry_FinalLocHist.unitref
WHERE
  (
    (
      (Units.UnitID)= [forms]![UnitsForm]![unitid]
    )
  )
GROUP BY
  Clients.clientCompanyName,
  MstrLease.MLNo,
  MstrLease.MLOrig,
  Schedule.SchNo,
  SchGrp.SGrpNO,
  Units.Unitvin,
  Units.unityr,
  Units.unitmake,
  Units.unitmodel,
  Units.unittype,
  qry_FinalLocHist.FirstOfFirstOldStreet
ORDER BY
  First(Schedule.SchID);
