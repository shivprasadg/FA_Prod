SELECT
  AcctRMKTInventoryRpt_Part1.*,
  Remarketing_AvgProjectedReserve.ProjectedSalePrice,
  IIf(
    Nz([ProjectedSalePrice], 0)> 1
    And Nz([ProjectedSalePrice], 0)< Nz([Final Cost], 0),
    CCur(
      Nz([ProjectedSalePrice], 0)- Nz([Final Cost], 0)
    ),
    0
  ) AS Reserve,
  IIf(
    Nz(
      [Remarketing_AvgSoldPrice_dated].[AvgSoldPrice],
      0
    )> Nz([Final Cost], 0),
    CCur(
      Nz(
        [Remarketing_AvgSoldPrice_dated].[AvgSoldPrice],
        0
      )- Nz([Final Cost], 0)
    ),
    0
  ) AS ProjectedProfit,
  Remarketing_AvgSoldPrice_dated.AvgSoldPrice AS DatedAvgSalePrice,
  Remarketing_AvgSoldPrice.AvgSoldPrice AS 90DaysAvgSoldPrice,
  IIf(
    Nz([ActualSoldPrice], 0)= 0,
    0,
    IIf(
      Nz([ActualSoldPrice], 0)> Nz([Final Cost], 0),
      CCur(
        Nz([ActualSoldPrice], 0)- Nz([Final Cost], 0)
      ),
      0
    )
  ) AS ActualProfit,
  IIf(
    Nz([ActualSoldPrice], 0)= 0,
    0,
    IIf(
      Nz([ActualSoldPrice], 0)< Nz([Final Cost], 0),
      CCur(
        Nz([ActualSoldPrice], 0)- Nz([Final Cost], 0)
      ),
      0
    )
  ) AS ActualLoss
FROM
  (
    (
      AcctRMKTInventoryRpt_Part1
      LEFT JOIN Remarketing_AvgSoldPrice ON AcctRMKTInventoryRpt_Part1.Description = Remarketing_AvgSoldPrice.Description
    )
    LEFT JOIN Remarketing_AvgSoldPrice_dated ON AcctRMKTInventoryRpt_Part1.Description = Remarketing_AvgSoldPrice_dated.Description
  )
  LEFT JOIN Remarketing_AvgProjectedReserve ON AcctRMKTInventoryRpt_Part1.Description = Remarketing_AvgProjectedReserve.Description;
