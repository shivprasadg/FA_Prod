SELECT
  Clients.ClientShNm AS Client,
  MstrLease.MLNo,
  Schedule.SchNo AS Schedule,
  SchGrp.SGrpNO AS UnitGroup,
  SchGrp.Assignee_Bank AS Assignee,
  Units.UnitIFBank AS LoanBank,
  Units.UnitIFDate AS LoanDate,
  Units.UnitIFAmt AS LoanAmount,
  "Interim" AS LoanType,
  SchGrp.SGrpExpDt AS InterimExpiry,
  Units.UnitPayOff AS PayoffDate,
  Units.UnitID AS AssetID,
  Units.UnitVIN AS VIN,
  Units.UnitStatus,
  [unitfinalcost] + Nz([PartsCost], 0) AS Cost,
  [VendorCost] + Nz([PartsCost], 0) AS Vendor_wParts,
  Units.UnitIFDate,
  Units.UnitIFBank,
  Units.UnitIFAmt,
  Units.UnitPayOff
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
    INNER JOIN Units ON SchGrp.SGrpID = Units.GroupREF
  )
  LEFT JOIN ChildPart_Cost_OEC_summed ON Units.UnitID = ChildPart_Cost_OEC_summed.UnitId
WHERE
  (
    (
      (Units.UnitIFAmt)> 0.001
    )
    AND (
      (Clients.ClientGroupID)<> 18
    )
    AND (
      (
        ChildPart_Cost_OEC_summed.ExclFrGrpOEC
      )= False
    )
  )
ORDER BY
  Clients.ClientShNm,
  MstrLease.MLNo,
  Schedule.SchID,
  Schedule.SchNo,
  SchGrp.SGrpNO;
