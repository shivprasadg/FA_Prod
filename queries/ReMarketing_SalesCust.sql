PARAMETERS [forms]![Rmkt Menu]![FromDate] DateTime,
[forms]![Rmkt Menu]![ToDate] DateTime;
SELECT
  Units.unitstatus,
  Clients.clientID,
  MstrLease.mlorig,
  Clients.clientshnm,
  MstrLease.mlno,
  Schedule.schno,
  Units.unitvin,
  Units.unitunitnum,
  Units.unityr,
  Units.unitmake,
  Units.unitmodel,
  Units.unittype,
  Rmkt.parkdt,
  Rmkt.BuyOutDate,
  lookup_GroupMakes.MakeAbrev,
  SchGrp.Assignee_Bank,
  Rmkt.BuyOutCost,
  Rmkt.SoldPrice,
  Rmkt.InvoicePaidDate,
  Rmkt.RmktTitleSent,
  Rmkt.RmktBOSSent,
  Clients_1.ClientShNm,
  ChildPartSummaryQuery.SumOfchildFinalCost,
  DatePart("yyyy", [InvoicePaidDate]) AS YR,
  Rmkt.MiParked,
  Rmkt.HrParked,
  Rmkt.BuyerRef,
  Rmkt.RmktSPRAdj,
  Rmkt.RmktAdjNote,
  Rmkt.RmkSm,
  Clients_1.[WH?],
  IIf(
    [adjbookvalue] > 0, [adjbookvalue],
    [BuyOutCost]
  ) AS Cost,
  Rmkt.TargetSale,
  Rmkt.InvoicePaidDate,
  lookup_groupTypes.GrpTypAbrev,
  Rmkt.RmktDeposit,
  Rmkt.RmktDepdt,
  IIf(
    [unitstatus] = "r"
    Or [unitstatus] = "k",
    "Pending Sale", "Sold"
  ) AS statussort,
  UnitVinDetails.series AS [Engine Series]
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
      (
        (
          (
            (
              Units
              LEFT JOIN Rmkt ON Units.UnitID = Rmkt.UnitRef
            )
            INNER JOIN lookup_GroupMakes ON Units.UnitMake = lookup_GroupMakes.groupMake
          )
          LEFT JOIN Clients AS Clients_1 ON Rmkt.BuyerRef = Clients_1.clientID
        )
        LEFT JOIN ChildPartSummaryQuery ON Units.UnitID = ChildPartSummaryQuery.UnitREF
      )
      INNER JOIN lookup_groupTypes ON Units.UnitType = lookup_groupTypes.groupType
    ) ON SchGrp.SGrpID = Units.GroupREF
  )
  LEFT JOIN UnitVinDetails ON Rmkt.UnitRef = UnitVinDetails.unitid
WHERE
  (
    (
      (Units.unitstatus)= "s"
      Or (Units.unitstatus)= "ks"
      Or (Units.unitstatus)= "rs"
    )
    AND (
      (Clients_1.ClientShNm) Is Not Null
    )
    AND (
      (Clients.ClientGroupID)<> 18
    )
  )
ORDER BY
  Clients.clientshnm,
  MstrLease.mlno,
  Schedule.schid;
