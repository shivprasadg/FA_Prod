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
  Cost_Refrig.RefrigMake AS RefrigMake,
  Cost_Refrig.RefrigSerialNum AS RefrigSerial,
  CCur(
    Nz([Cost_Refrig].[Refrig], 0)
  ) AS Refrig,
  Cost_Refrig.RefrigInvoice,
  Cost_Refrig.RefrigPaid,
  Cost_FeesTax.TaxName,
  CCur(
    Nz([Cost_FeesTax].[FeesTax], 0)
  ) AS FeesTax,
  Cost_FeesTax.TaxInvoice,
  Cost_FeesTax.TaxPaid,
  Cost_Freight.FreightName,
  CCur(
    Nz([Cost_Freight].[Freight], 0)
  ) AS Freight,
  Cost_Freight.FGTInvoice,
  Cost_Freight.FreightPaid,
  Cost_Warranty.WarrantyName,
  CCur(
    Nz([Cost_Warranty].[Warranty], 0)
  ) AS Warranty,
  Cost_Warranty.WarrantyInv,
  Cost_Warranty.WarrantyPaid,
  Cost_Inspection.InspCo,
  CCur(
    Nz(
      [Cost_Inspection].[Inspection],
      0
    )
  ) AS Inspection,
  Cost_Inspection.InspInvoice,
  Cost_Inspection.InspPaid,
  Cost_Other.OtherMake,
  CCur(
    Nz([Cost_Other].[other], 0)
  ) AS Other,
  Cost_Other.OtherInvoice,
  Cost_Other.OtherPaid,
  Cost_PTO.PTOMake,
  CCur(
    Nz([Cost_PTO].[PTO], 0)
  ) AS PTO,
  Cost_PTO.PTOInvoice,
  Cost_PTO.PTOPaid,
  Cost_DecalsCoName.Name AS CoNameDecal,
  CCur(
    Nz(
      [Cost_DecalsCoName].[CoNameCost],
      0
    )
  ) AS CoNameCost,
  Cost_DecalsCoName.CoNameInvoice,
  Cost_DecalsCoName.CoNamePaid,
  Cost_DecalsRegNumber.Name AS RegDecal,
  CCur(
    Nz(
      [Cost_DecalsRegNumber].[RegDecalCost],
      0
    )
  ) AS RegDecalCost,
  Cost_DecalsRegNumber.RegDecalInvoice,
  Cost_DecalsRegNumber.RegDecalPaid,
  Cost_DecalsMisc.Name AS MiscDecal,
  CCur(
    Nz([Cost_DecalsMisc].[MiscCost], 0)
  ) AS MiscCost,
  Cost_DecalsMisc.MiscInvoice,
  Cost_DecalsMisc.MiscPaid,
  Cost_DecalsMisc2.Name AS Misc2Decal,
  CCur(
    Nz(
      [Cost_DecalsMisc2].[MiscCost], 0
    )
  ) AS Misc2Cost,
  Cost_DecalsMisc2.MiscInvoice AS Misc2Invoice,
  Cost_DecalsMisc2.MiscPaid AS Misc2Paid,
  Cost_TankBody.TankMake,
  CCur(
    Nz([Cost_TankBody].[tankbody], 0)
  ) AS TankBody,
  Cost_TankBody.TankInvoice,
  Cost_TankBody.TankPaid,
  Cost_TankOther.TankOthMake,
  CCur(
    Nz([Cost_TankOther].[tankOther], 0)
  ) AS TankOther,
  Cost_TankOther.TankOthInvoice,
  Cost_TankOther.TankOthPaid,
  Cost_Oil.OilRigMake,
  CCur(
    Nz([Cost_Oil].[OilRig], 0)
  ) AS OilRig,
  Cost_Oil.OilRigInvoice,
  Cost_Oil.OilRigPaid,
  Cost_OBC_All.OBC_Make,
  Cost_OBC_All.OBC_Cost,
  Cost_OBC_All.OBCInvoice,
  Cost_OBC_All.OBCPaid,
  Cost_OBC_All.OBC_Display,
  Cost_OBC_All.OBCDisplay,
  Cost_OBC_All.OBCDisplayInv,
  Cost_OBC_All.OBCDisplayPaid,
  Cost_OBC_All.OBC_Adapt,
  Cost_OBC_All.OBCAdaptr,
  Cost_OBC_All.OBCAdaptInv,
  Cost_OBC_All.OBCAdaptPaid,
  Cost_OBC_All.OBC_Dome,
  Cost_OBC_All.OBCDome,
  Cost_OBC_All.OBCDomeInv,
  Cost_OBC_All.OBCDomePaid,
  Cost_OBC_All.OBC_Multi,
  Cost_OBC_All.OBCMulti,
  Cost_OBC_All.OBCMultiInv,
  Cost_OBC_All.OBCMultiPaid,
  Cost_OBC_All.OBCInstaller,
  Cost_OBC_All.OBCInstall,
  Cost_OBC_All.OBCInsInvoice,
  Cost_OBC_All.OBCInstallPaid,
  Cost_OBC_All.OBC_Part AS OBCPartMake,
  Cost_OBC_All.OBCPart,
  Cost_OBC_All.OBCPartInv,
  Cost_OBC_All.OBCPartPaid,
  CCur(
    [UnitCost] + [Refrig] + [OBCTotal] + [FeesTax] + [Freight] + [Warranty] + [Other] + [Inspection] + [PTO] + [CoNameCost] + [RegDecalCost] + [MiscCost] + [Misc2Cost] + [TankBody] + [TankOther] + [OilRig]
  ) AS UnitTotalCost,
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
  (
    (
      (
        (
          (
            (
              (
                (
                  (
                    (
                      (
                        (
                          (
                            (
                              ScheduleUnitDocs3
                              LEFT JOIN Cost_FeesTax ON ScheduleUnitDocs3.AssetId = Cost_FeesTax.AssetId
                            )
                            LEFT JOIN Cost_Freight ON ScheduleUnitDocs3.AssetId = Cost_Freight.AssetId
                          )
                          LEFT JOIN Cost_Refrig ON ScheduleUnitDocs3.AssetId = Cost_Refrig.AssetId
                        )
                        LEFT JOIN Cost_Warranty ON ScheduleUnitDocs3.AssetId = Cost_Warranty.AssetId
                      )
                      LEFT JOIN Cost_Other ON ScheduleUnitDocs3.AssetId = Cost_Other.AssetId
                    )
                    LEFT JOIN Cost_PTO ON ScheduleUnitDocs3.AssetId = Cost_PTO.AssetId
                  )
                  LEFT JOIN Cost_Inspection ON ScheduleUnitDocs3.AssetId = Cost_Inspection.AssetId
                )
                LEFT JOIN Cost_DecalsCoName ON ScheduleUnitDocs3.AssetId = Cost_DecalsCoName.AssetId
              )
              LEFT JOIN Cost_DecalsRegNumber ON ScheduleUnitDocs3.AssetId = Cost_DecalsRegNumber.AssetId
            )
            LEFT JOIN Cost_TankBody ON ScheduleUnitDocs3.AssetId = Cost_TankBody.AssetId
          )
          LEFT JOIN Cost_TankOther ON ScheduleUnitDocs3.AssetId = Cost_TankOther.AssetId
        )
        LEFT JOIN Cost_Oil ON ScheduleUnitDocs3.AssetId = Cost_Oil.AssetId
      )
      LEFT JOIN Cost_DecalsMisc ON ScheduleUnitDocs3.AssetId = Cost_DecalsMisc.AssetId
    )
    LEFT JOIN Cost_OBC_All ON ScheduleUnitDocs3.AssetId = Cost_OBC_All.AssetId
  )
  LEFT JOIN Cost_DecalsMisc2 ON ScheduleUnitDocs3.AssetId = Cost_DecalsMisc2.AssetId
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
