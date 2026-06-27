SELECT
  Format([AcceptedDate], "mmm") AS AcceptedMonth,
  Format([AcceptedDate], "yyyy") AS AcceptedYear,
  vw_SixKeys.*,
  CCur(
    Nz([Freight], 0)
  ) AS FreightCost,
  CCur(
    Nz([Decals], 0)
  ) AS DecalCost,
  CCur(
    Nz([Refrig], 0)
  ) AS RefrigCost,
  CCur(
    Nz([FeesTax], 0)
  ) AS FeesTaxesCost,
  Nz([Cost_OBC].[OBC_Make], "--") AS OBCMake,
  CCur(
    Nz([Cost_OBC].[OBC], 0)
  ) AS OBCCost,
  Nz(Cost_OBC.[OBCInvoice], "--") AS OBCInvoice,
  Nz(
    [Cost_OBC_Install].[OBCInstaller],
    "--"
  ) AS OBCInstaller,
  CCur(
    Nz(
      Cost_OBC_Install.[OBCInstall], 0
    )
  ) AS OBCInstall,
  Nz(
    Cost_OBC_Install.[OBCInsInvoice],
    "--"
  ) AS OBCInstall_Invoice,
  CCur(
    Nz([Warranty], 0)
  ) AS WarrantyCost,
  CCur(
    Nz([Other], 0)
  ) AS OtherCost,
  vw_SixKeys.UnitInvoiceDate AS InvoiceDate
FROM
  (
    (
      (
        (
          (
            (
              (
                vw_SixKeys
                LEFT JOIN Cost_Freight ON vw_SixKeys.AssetID = Cost_Freight.AssetId
              )
              LEFT JOIN Cost_Decals ON vw_SixKeys.AssetID = Cost_Decals.AssetId
            )
            LEFT JOIN Cost_Refrig ON vw_SixKeys.AssetID = Cost_Refrig.AssetId
          )
          LEFT JOIN Cost_FeesTax ON vw_SixKeys.AssetID = Cost_FeesTax.AssetId
        )
        LEFT JOIN Cost_OBC ON vw_SixKeys.AssetID = Cost_OBC.AssetId
      )
      LEFT JOIN Cost_Warranty ON vw_SixKeys.AssetID = Cost_Warranty.AssetId
    )
    LEFT JOIN Cost_Other ON vw_SixKeys.AssetID = Cost_Other.AssetId
  )
  LEFT JOIN Cost_OBC_Install ON vw_SixKeys.AssetID = Cost_OBC_Install.AssetId
ORDER BY
  vw_SixKeys.UnitInvoiceDate DESC;
