SELECT
  ScheduleUnitDocs3.[Lease No],
  ScheduleUnitDocs3.Lessee,
  ScheduleUnitDocs3.Co_Lessee,
  ScheduleUnitDocs3.SchNo,
  ScheduleUnitDocs3.Group,
  ScheduleUnitDocs3.[Unit No],
  ScheduleUnitDocs3.Unitvin,
  ScheduleUnitDocs3.Desc,
  ScheduleUnitDocs3.UnitCost,
  ScheduleUnitDocs3.UnitInvoiceNum,
  ScheduleUnitDocs3.UnitInvoicePaid,
  Cost_ChildParts_AllTypes.[APU Make] AS Expr1,
  Cost_ChildParts_AllTypes.[APU Cost] AS Expr2,
  Cost_ChildParts_AllTypes.[APU Invoice] AS Expr3,
  Cost_ChildParts_AllTypes.[APU Paid] AS Expr4,
  Cost_ChildParts_AllTypes.[APU DrawNum] AS Expr5,
  Cost_ChildParts_AllTypes.[Deer Guards Make] AS Expr6,
  Cost_ChildParts_AllTypes.[Deer Guards Cost] AS Expr7,
  Cost_ChildParts_AllTypes.[Deer Guards Invoice] AS Expr8,
  Cost_ChildParts_AllTypes.[Deer Guards Paid] AS Expr9,
  Cost_ChildParts_AllTypes.[Deer Guards DrawNum] AS Expr10,
  Cost_ChildParts_AllTypes.[Headboards & Toolboxes Make] AS Expr11,
  Cost_ChildParts_AllTypes.[Headboards & Toolboxes Cost] AS Expr12,
  Cost_ChildParts_AllTypes.[Headboards & Toolboxes Invoice] AS Expr13,
  Cost_ChildParts_AllTypes.[Headboards & Toolboxes Paid] AS Expr14,
  Cost_ChildParts_AllTypes.[Headboards & Toolboxes DrawNum] AS Expr15,
  Cost_ChildParts_AllTypes.[Installation Make] AS Expr16,
  Cost_ChildParts_AllTypes.[Installation Cost] AS Expr17,
  Cost_ChildParts_AllTypes.[Installation Invoice] AS Expr18,
  Cost_ChildParts_AllTypes.[Installation Paid] AS Expr19,
  Cost_ChildParts_AllTypes.[Installation DrawNum] AS Expr20,
  Cost_ChildParts_AllTypes.[Inverters Make] AS Expr21,
  Cost_ChildParts_AllTypes.[Inverters Cost] AS Expr22,
  Cost_ChildParts_AllTypes.[Inverters Invoice] AS Expr23,
  Cost_ChildParts_AllTypes.[Inverters Paid] AS Expr24,
  Cost_ChildParts_AllTypes.[Inverters DrawNum] AS Expr25,
  Cost_ChildParts_AllTypes.ChildPartTotal,
  [UnitCost] + [ChildPartTotal] AS UnitTotalCost,
  ScheduleUnitDocs3.AssetId,
  ScheduleUnitDocs3.Status,
  ScheduleUnitDocs3.InService,
  ScheduleUnitDocs3.ActualDelDate,
  ScheduleUnitDocs3.[Title Num],
  ScheduleUnitDocs3.TitledState,
  ScheduleUnitDocs3.Location,
  ScheduleUnitDocs3.StreetAddress,
  ScheduleUnitDocs3.City,
  ScheduleUnitDocs3.State,
  ScheduleUnitDocs3.Zip,
  ScheduleUnitDocs3.InterimBank,
  ScheduleUnitDocs3.InterimDate,
  ScheduleUnitDocs3.InterimAmt
FROM
  ScheduleUnitDocs3
  INNER JOIN Cost_ChildParts_AllTypes ON ScheduleUnitDocs3.AssetId = Cost_ChildParts_AllTypes.AssetId
WHERE
  (
    (
      (ScheduleUnitDocs3.AssetId) Is Not Null
    )
  )
ORDER BY
  ScheduleUnitDocs3.Group,
  ScheduleUnitDocs3.[Unit No],
  ScheduleUnitDocs3.AssetId;
