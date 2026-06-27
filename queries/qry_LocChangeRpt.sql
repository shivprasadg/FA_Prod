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
  LocChangeT.OLocationStreet,
  LocChangeT.OCity,
  LocChangeT.OCounty,
  LocChangeT.OSt,
  LocChangeT.Ozip,
  LocChangeT.NLocationStreet,
  LocChangeT.NCity,
  LocChangeT.NCounty,
  LocChangeT.NSt,
  LocChangeT.Nzip,
  LocChangeT.LocChDt,
  LocChangeT.LocRptDt,
  LocChangeT.OldTxstatus,
  LocChangeT.OTaxNotes,
  LocChangeT.Newtaxstatus,
  LocChangeT.NewTaxnotes,
  Clients_1.clientCompanyName AS [Assignee Bank],
  LocChangeT.ID
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
      LEFT JOIN Clients AS Clients_1 ON SchGrp.Assignee_Bank = Clients_1.ClientShNm
    ) ON Schedule.SchID = SchGrp.SGrpSchID
  )
  INNER JOIN (
    Units
    INNER JOIN LocChangeT ON Units.UnitID = LocChangeT.unitref
  ) ON SchGrp.SGrpID = Units.GroupREF
WHERE
  (
    (
      (LocChangeT.LocChDt)>= Date()
    )
    And (
      (Units.UnitID)= forms!UnitsForm!unitid
    )
  )
ORDER BY
  LocChangeT.LocChDt DESC,
  Schedule.SchID;
