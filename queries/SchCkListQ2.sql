SELECT
  Clients.clientCompanyName,
  Clients.ClientShNm,
  MstrLease.MLNo,
  MstrLease.MLSent,
  MstrLease.MLBack,
  MstrLease.MLGuarantor,
  MstrLease.MLGuarBack,
  Schedule.SchID,
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
  Schedule.SchID,
  SchGrp.SGrpNO,
  Right([unitvin], 6) AS VIN4,
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
    [unityr] & " " & [Makeabrev] & " " & [unitmodel] & " " & [unittype] & " " & IIf([axle] = "tbd", Null, [axle])
  ) AS UDesc,
  Units.UnitTitleRecd,
  Locations.LState,
  Units.UnitID,
  lookup_States.StTaxType,
  IIf(
    IsNull([SumOfchildFinalCost]),
    0,
    [SumOfchildFinalCost]
  ) AS TotChild,
  IIf(
    IsNull([SumOfchildFinalCost]),
    Null,
    [SumOfchildFinalCost]
  )+ [unitfinalcost] AS TotUnit,
  ChildPartSummaryQuery.SumOfchildFinalCost,
  Units.UnitID,
  Units.unitunitnum2,
  [Last Title Entry].LastOfUnitTitleLogDate,
  [Last Title Entry].Action
FROM
  (
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
        (
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
        )
        LEFT JOIN ChildPartSummaryQuery ON Units.UnitID = ChildPartSummaryQuery.UnitREF
      ) ON SchGrp.SGrpID = Units.GroupREF
    )
    LEFT JOIN lookup_GroupMakes ON Units.UnitMake = lookup_GroupMakes.groupMake
  )
  LEFT JOIN [Last Title Entry] ON Units.UnitID = [Last Title Entry].UnitRef
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
  Schedule.SchID,
  SchGrp.SGrpNO,
  Right([unitvin], 6),
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
    [unityr] & " " & [Makeabrev] & " " & [unitmodel] & " " & [unittype] & " " & IIf([axle] = "tbd", Null, [axle])
  ),
  Units.UnitTitleRecd,
  Locations.LState,
  lookup_States.StTaxType,
  IIf(
    IsNull([SumOfchildFinalCost]),
    0,
    [SumOfchildFinalCost]
  ),
  IIf(
    IsNull([SumOfchildFinalCost]),
    Null,
    [SumOfchildFinalCost]
  )+ [unitfinalcost],
  ChildPartSummaryQuery.SumOfchildFinalCost,
  Units.UnitID,
  Units.unitunitnum2,
  [Last Title Entry].LastOfUnitTitleLogDate,
  [Last Title Entry].Action,
  Schedule.SchID,
  SchGrp.SGrpNO,
  Units.UnitID,
  Units.UnitUnitNum,
  Units.UnitVIN
ORDER BY
  Clients.clientCompanyName,
  SchGrp.SGrpNO,
  Right([unitvin], 6);
