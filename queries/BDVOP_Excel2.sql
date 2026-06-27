SELECT
  (
    Nz(
      [BDVOP_Excel1].[SentYear], [BDVOP_Excel1].[DealYear]
    )
  ) AS SentYear,
  BDVOP_Excel1.Client,
  BDVOP_Excel1.FAID,
  BDVOP_Excel1.DealBEDEX,
  BDVOP_Excel1.DealYear,
  CInt(
    Nz([BDVOP_Excel1].[DealUnitQty], 0)
  ) AS DealUnitQty,
  BDVOP_Excel1.DealGrandTotal,
  CInt(
    Nz([UnitsInGroup], 0)
  ) AS Units,
  CCur(
    Nz([OrigEquipCost], 0)
  ) AS OEC,
  (
    CInt(
      Nz([UnitsInGroup], 0)
    )- CInt(
      Nz([DealUnitQty], 0)
    )
  ) AS UnitDiff,
  [OEC] - [DealGrandTotal] AS OECDiff,
  BDVOP_Excel1.OutcomeDate
FROM
  BDVOP_Excel1
  LEFT JOIN BDVOP_Excel0 ON BDVOP_Excel1.FAID = BDVOP_Excel0.LFAID
GROUP BY
  (
    Nz(
      [BDVOP_Excel1].[SentYear], [BDVOP_Excel1].[DealYear]
    )
  ),
  BDVOP_Excel1.Client,
  BDVOP_Excel1.FAID,
  BDVOP_Excel1.DealBEDEX,
  BDVOP_Excel1.DealYear,
  CInt(
    Nz([BDVOP_Excel1].[DealUnitQty], 0)
  ),
  BDVOP_Excel1.DealGrandTotal,
  CInt(
    Nz([UnitsInGroup], 0)
  ),
  CCur(
    Nz([OrigEquipCost], 0)
  ),
  (
    CInt(
      Nz([UnitsInGroup], 0)
    )- CInt(
      Nz([DealUnitQty], 0)
    )
  ),
  BDVOP_Excel1.OutcomeDate
ORDER BY
  (
    Nz(
      [BDVOP_Excel1].[SentYear], [BDVOP_Excel1].[DealYear]
    )
  ) DESC,
  BDVOP_Excel1.Client;
