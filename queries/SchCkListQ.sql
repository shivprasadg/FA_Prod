SELECT
  Clients.clientCompanyName,
  Clients.ClientShNm,
  MstrLease.MLNo,
  MstrLease.MLSent,
  MstrLease.MLBack,
  MstrLease.MLGuarantor,
  MstrLease.MLGuarBack,
  MstrLease.MLGuarSent,
  MstrLease.MLInsExp,
  MstrLease.MLINRecA,
  MstrLease.MLINSBLANKET,
  MstrLease.MLINSAccpt,
  Schedule.SchNo,
  SchGrp.SGrpNO,
  SchGrp.UnitsInGroup,
  SchGrp.OrigEquipCost,
  Schedule.SchOut,
  Schedule.SchIn,
  SchGrp.SGrpNO,
  Units.UnitID,
  Units.UnitUnitNum,
  Units.UnitVIN,
  Units.UnitAcceptDate,
  Units.UnitCOABack,
  Units.UnitVINLocked,
  Units.UnitInvoiceNum,
  Units.UnitInvoiceDate,
  Units.UnitFinalCost,
  Units.UnitVendorPytDate,
  Units.UnitMSORecd,
  Units.UnitIRP,
  Units.UnitTaxExempt,
  Units.UnitTaxNotes,
  Units.UnitTitleNumber,
  Units.UnitMake,
  Units.UnitModel,
  Units.UnitType,
  Sum(ChildParts.childFinalCost) AS SumOfchildFinalCost,
  UCase(
    [unityr] & " " & [unitmake] & " " & [unitmodel] & " " & [unittype]
  ) AS UDesc,
  Units.UnitTitleRecd,
  Locations.LState,
  Units.UnitID,
  lookup_States.StTaxType
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
      INNER JOIN SchGrp ON Schedule.SchID = SchGrp.SGrpSchID
    )
    INNER JOIN (
      lookup_States
      RIGHT JOIN (
        (
          (
            Units
            LEFT JOIN Locations ON Units.UnitGarageLocationRef = Locations.LocationID
          )
          LEFT JOIN Locations AS Locations_1 ON Units.UnitTitleLocationRef = Locations_1.LocationID
        )
        LEFT JOIN Locations AS Locations_2 ON Units.UnitRegLocationRef = Locations_2.LocationID
      ) ON lookup_States.StateAbbrev = Locations.LState
    ) ON SchGrp.SGrpID = Units.GroupREF
  )
  LEFT JOIN ChildParts ON Units.UnitID = ChildParts.UnitREF
GROUP BY
  Clients.clientCompanyName,
  Clients.ClientShNm,
  MstrLease.MLNo,
  MstrLease.MLSent,
  MstrLease.MLBack,
  MstrLease.MLGuarantor,
  MstrLease.MLGuarBack,
  MstrLease.MLGuarSent,
  MstrLease.MLInsExp,
  MstrLease.MLINRecA,
  MstrLease.MLINSBLANKET,
  MstrLease.MLINSAccpt,
  Schedule.SchNo,
  SchGrp.UnitsInGroup,
  SchGrp.OrigEquipCost,
  Schedule.SchOut,
  Schedule.SchIn,
  SchGrp.SGrpNO,
  Units.UnitUnitNum,
  Units.UnitVIN,
  Units.UnitAcceptDate,
  Units.UnitCOABack,
  Units.UnitVINLocked,
  Units.UnitInvoiceNum,
  Units.UnitInvoiceDate,
  Units.UnitFinalCost,
  Units.UnitVendorPytDate,
  Units.UnitMSORecd,
  Units.UnitIRP,
  Units.UnitTaxExempt,
  Units.UnitTaxNotes,
  Units.UnitTitleNumber,
  Units.UnitMake,
  Units.UnitModel,
  Units.UnitType,
  UCase(
    [unityr] & " " & [unitmake] & " " & [unitmodel] & " " & [unittype]
  ),
  Units.UnitTitleRecd,
  Locations.LState,
  Units.UnitID,
  lookup_States.StTaxType,
  SchGrp.SGrpNO,
  Units.UnitID,
  Units.UnitID,
  Right([unitvin], 6),
  Units.UnitUnitNum,
  Units.UnitVIN
HAVING
  (
    (
      (MstrLease.MLNo)= [ENTER MASTER LEASE ID]
    )
    AND (
      (Schedule.SchNo)= [ENTER SCHED NO]
    )
  )
ORDER BY
  Clients.clientCompanyName,
  Right([unitvin], 6),
  Units.UnitUnitNum,
  Units.UnitVIN;
