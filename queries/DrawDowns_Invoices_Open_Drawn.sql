SELECT
  DDI.Invoice,
  DDI.Vendor,
  [Schedule] + ' | ' + [UnitGroup] AS SchGrp,
  Max(DDI.Open) AS [Open],
  Format(
    Sum(
      IIf(
        IsNull([DDI].[DrawDownId]),
        [VendorCost],
        0
      )
    ),
    '$#,###'
  ) AS [Open Amt],
  Max(DDI.Drawn) AS Drawn,
  Format(
    Sum(
      IIf(
        IsNull([DDI].[DrawDownId]),
        0,
        [IFAmount]
      )
    ),
    '$#,###'
  ) AS [Drawn Amt],
  DDI.DrawDownId AS DrawDownId,
  DDI.VendorID,
  DDI.ClientID,
  DDI.MLID,
  DDI.SchID,
  DDI.sGrpID,
  IsSplit.Split
FROM
  vw_DrawDowns_Invoices AS DDI
  LEFT JOIN (
    SELECT
      DISTINCTROW DrawCnt.UnitInvoiceNum,
      Count(*) AS Split
    FROM
      (
        SELECT
          DISTINCT Units.UnitInvoiceNum,
          Units.DrawDownId
        FROM
          Units
      ) AS DrawCnt
    WHERE
      (
        (
          (DrawCnt.DrawDownId)> 0
        )
      )
    GROUP BY
      DrawCnt.UnitInvoiceNum
    HAVING
      (
        (
          (
            Count(*)
          )> 1
        )
      )
  ) AS IsSplit ON DDI.Invoice = IsSplit.UnitInvoiceNum
GROUP BY
  DDI.Invoice,
  DDI.Vendor,
  [Schedule] + ' | ' + [UnitGroup],
  DDI.DrawDownId,
  DDI.VendorID,
  DDI.ClientID,
  DDI.MLID,
  DDI.SchID,
  DDI.sGrpID,
  IsSplit.Split;
