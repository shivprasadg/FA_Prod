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
  Format(
    Nz(
      [Units].[UnitCOABack], [Units].[UnitCofACopy]
    ),
    "mm/dd/yy"
  ) AS UnitCOABack,
  Units.UnitVINLocked,
  Units.UnitInvoiceNum,
  Units.UnitInvoiceDate,
  IIf(
    [Vendorcost] > 0, [VendorCost], [UnitFinalCost]
  ) AS UnitCost,
  Units.UnitVendorPytDate,
  Units.UnitMSORecd,
  Units.UnitIRP,
  Units.UnitTaxExempt,
  Units.UnitTaxNotes,
  Units.UnitTitleNumber,
  Units.UnitMake,
  Units.UnitModel,
  Units.UnitType,
  [UnitYr] & " " & [UnitMake] & " " & [UnitModel] & IIf([Axle] = "tbd", Null, " " & [Axle])& IIf(
    Nz([UnitConfiguration], "")= "",
    Null,
    " " & [UnitConfiguration]
  )& " " & [UnitSubType] & " " & [UnitType] AS UnitDesc,
  Units.UnitTitleRecd,
  Locations.LState,
  Units.UnitID,
  lookup_States.StTaxType,
  IIf(
    IsNull([ChildPartCost]),
    0,
    [ChildPartCost]
  ) AS TotChild,
  Nz([ChildPartCost], 0)+ IIf(
    [Vendorcost] > 0, [VendorCost], [UnitFinalCost]
  ) AS TotUnit,
  ChildPart_Cost_OEC_summed.ChildPartCost AS SumOfchildFinalCost,
  Units.UnitID,
  Units.unitunitnum2,
  LastTitleLogEntry.LogDate AS LastOfUnitTitleLogDate,
  LastTitleLogEntry.Action,
  Banks.BankName AS Assignee,
  Units.VendorCost
FROM
  (
    (
      Schedule_Group_WorkTable
      INNER JOIN (
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
              LEFT JOIN ChildPart_Cost_OEC_summed ON Units.UnitID = ChildPart_Cost_OEC_summed.UnitID
            ) ON SchGrp.SGrpID = Units.GroupREF
          )
          LEFT JOIN lookup_GroupMakes ON Units.UnitMake = lookup_GroupMakes.groupMake
        )
        LEFT JOIN [Last Title Entry] ON Units.UnitID = [Last Title Entry].UnitRef
      ) ON Schedule_Group_WorkTable.ScheduleId = Schedule.SchID
    )
    LEFT JOIN Banks ON SchGrp.Assignee_Bank = Banks.BankShortName
  )
  LEFT JOIN LastTitleLogEntry ON Units.UnitID = LastTitleLogEntry.AssetId
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
  Format(
    Nz(
      [Units].[UnitCOABack], [Units].[UnitCofACopy]
    ),
    "mm/dd/yy"
  ),
  Units.UnitVINLocked,
  Units.UnitInvoiceNum,
  Units.UnitInvoiceDate,
  IIf(
    [Vendorcost] > 0, [VendorCost], [UnitFinalCost]
  ),
  Units.UnitVendorPytDate,
  Units.UnitMSORecd,
  Units.UnitIRP,
  Units.UnitTaxExempt,
  Units.UnitTaxNotes,
  Units.UnitTitleNumber,
  Units.UnitMake,
  Units.UnitModel,
  Units.UnitType,
  [UnitYr] & " " & [UnitMake] & " " & [UnitModel] & IIf([Axle] = "tbd", Null, " " & [Axle])& IIf(
    Nz([UnitConfiguration], "")= "",
    Null,
    " " & [UnitConfiguration]
  )& " " & [UnitSubType] & " " & [UnitType],
  Units.UnitTitleRecd,
  Locations.LState,
  lookup_States.StTaxType,
  Nz([ChildPartCost], 0)+ IIf(
    [Vendorcost] > 0, [VendorCost], [UnitFinalCost]
  ),
  ChildPart_Cost_OEC_summed.ChildPartCost,
  Units.UnitID,
  Units.unitunitnum2,
  LastTitleLogEntry.LogDate,
  LastTitleLogEntry.Action,
  Banks.BankName,
  Units.VendorCost,
  Schedule.SchID,
  SchGrp.SGrpNO,
  Units.UnitID,
  Units.UnitUnitNum,
  Units.UnitVIN
ORDER BY
  Clients.clientCompanyName,
  SchGrp.SGrpNO,
  Right([unitvin], 6);
