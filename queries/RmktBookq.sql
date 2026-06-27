SELECT
  Units.unitstatus,
  MstrLease.mlorig,
  Clients.clientshnm,
  MstrLease.mlno,
  Schedule.schno,
  Units.unitvin,
  Units.UnitUnitNum,
  Units.UnitYr,
  Units.UnitMake,
  Units.UnitModel,
  Units.UnitType,
  Units.UnitSubType,
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
  Clients.[WH?],
  lookup_groupTypes.grpClass,
  Clients_1.clientCompanyName
FROM
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
    LEFT JOIN lookup_groupTypes ON Units.UnitType = lookup_groupTypes.groupType
  ) ON SchGrp.SGrpID = Units.GroupREF
WHERE
  (
    (
      (Units.unitstatus)= "s"
    )
  )
ORDER BY
  Clients.clientshnm,
  MstrLease.mlno,
  Schedule.schid;
