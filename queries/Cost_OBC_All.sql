SELECT
  Units.UnitID AS AssetId,
  Cost_OBC.OBC_Make,
  CCur(
    Nz([OBC], 0)
  ) AS OBC_Cost,
  Cost_OBC.OBCInvoice,
  Cost_OBC.OBCPaid,
  Cost_OBC_Adapt.OBC_Adapt,
  CCur(
    Nz([OBCAdapt], 0)
  ) AS OBCAdaptr,
  Cost_OBC_Adapt.OBCAdaptInv,
  Cost_OBC_Adapt.OBCAdaptPaid,
  Cost_OBC_Dome.OBC_Dome,
  CCur(
    Nz([Cost_OBC_Dome].[OBCDome], 0)
  ) AS OBCDome,
  Cost_OBC_Dome.OBCDomeInv,
  Cost_OBC_Dome.OBCDomePaid,
  Cost_OBC_Multi.OBC_Multi,
  CCur(
    Nz([Cost_OBC_Multi].[OBCMulti])
  ) AS OBCMulti,
  Cost_OBC_Multi.OBCMultiInv,
  Cost_OBC_Multi.OBCMultiPaid,
  Cost_OBC_Display.OBC_Display,
  CCur(
    Nz(
      [Cost_OBC_Display].[OBCDisplay],
      0
    )
  ) AS OBCDisplay,
  Cost_OBC_Display.OBCDisplayInv,
  Cost_OBC_Display.OBCDisplayPaid,
  Cost_OBC_Install.OBCInstaller,
  CCur(
    Nz(
      [Cost_OBC_Install].[OBCInstall],
      0
    )
  ) AS OBCInstall,
  Cost_OBC_Install.OBCInsInvoice,
  Cost_OBC_Install.OBCInstallPaid,
  Cost_OBC_Part.OBC_Part,
  CCur(
    Nz([Cost_OBC_Part].[OBCPart], 0)
  ) AS OBCPart,
  Cost_OBC_Part.OBCPartInv,
  Cost_OBC_Part.OBCPartPaid,
  CCur(
    Nz([OBC], 0)+ Nz([OBCAdapt], 0)+ Nz([OBCDome], 0)+ Nz([OBCMulti], 0)+ Nz([OBCInstall], 0)+ Nz(
      [Cost_OBC_Display].[OBCDisplay],
      0
    )+ [OBCPart]
  ) AS OBCtotal
FROM
  (
    (
      (
        (
          (
            (
              Units
              LEFT JOIN Cost_OBC ON Units.UnitID = Cost_OBC.AssetId
            )
            LEFT JOIN Cost_OBC_Adapt ON Units.UnitID = Cost_OBC_Adapt.AssetId
          )
          LEFT JOIN Cost_OBC_Dome ON Units.UnitID = Cost_OBC_Dome.AssetId
        )
        LEFT JOIN Cost_OBC_Multi ON Units.UnitID = Cost_OBC_Multi.AssetId
      )
      LEFT JOIN Cost_OBC_Install ON Units.UnitID = Cost_OBC_Install.AssetId
    )
    LEFT JOIN Cost_OBC_Display ON Units.UnitID = Cost_OBC_Display.AssetId
  )
  LEFT JOIN Cost_OBC_Part ON Units.UnitID = Cost_OBC_Part.AssetId;
