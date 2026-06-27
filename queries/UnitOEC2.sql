SELECT
  UnitOEC1.AssetId,
  UnitOEC1.UnitNo,
  UnitOEC1.Unitvin,
  UnitOEC1.Status,
  UnitOEC1.GroupREF,
  UnitOEC1.Desc,
  UnitOEC1.UnitCost,
  CCur(
    [UnitCost] + [Refrig] + [OBCTotal] + [FeesTax] + [Freight] + [Warranty] + [Other] + [Inspection] + [PTO] + [CoNameCost] + [RegDecalCost] + [MiscCost] + [TankBody] + [TankOther] + [OilRig]
  ) AS TotalCost,
  UnitOEC1.UnitInvoiceNum,
  UnitOEC1.UnitInvoicePaid,
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
  UnitOEC1.InService,
  UnitOEC1.[Title Num],
  UnitOEC1.TitledState,
  UnitOEC1.Location,
  UnitOEC1.StreetAddress,
  UnitOEC1.City,
  UnitOEC1.State,
  UnitOEC1.Zip,
  UnitOEC1.InterimBank,
  UnitOEC1.InterimDate,
  UnitOEC1.InterimAmt,
  UnitOEC1.DrawDownId,
  UnitOEC1.VendorREF,
  UnitOEC1.PerDiemRent,
  UnitOEC1.LORent,
  UnitOEC1.SYNDRent
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
                            UnitOEC1
                            LEFT JOIN Cost_FeesTax ON UnitOEC1.AssetId = Cost_FeesTax.AssetId
                          )
                          LEFT JOIN Cost_Freight ON UnitOEC1.AssetId = Cost_Freight.AssetId
                        )
                        LEFT JOIN Cost_Refrig ON UnitOEC1.AssetId = Cost_Refrig.AssetId
                      )
                      LEFT JOIN Cost_Warranty ON UnitOEC1.AssetId = Cost_Warranty.AssetId
                    )
                    LEFT JOIN Cost_Other ON UnitOEC1.AssetId = Cost_Other.AssetId
                  )
                  LEFT JOIN Cost_PTO ON UnitOEC1.AssetId = Cost_PTO.AssetId
                )
                LEFT JOIN Cost_Inspection ON UnitOEC1.AssetId = Cost_Inspection.AssetId
              )
              LEFT JOIN Cost_DecalsCoName ON UnitOEC1.AssetId = Cost_DecalsCoName.AssetId
            )
            LEFT JOIN Cost_DecalsRegNumber ON UnitOEC1.AssetId = Cost_DecalsRegNumber.AssetId
          )
          LEFT JOIN Cost_TankBody ON UnitOEC1.AssetId = Cost_TankBody.AssetId
        )
        LEFT JOIN Cost_TankOther ON UnitOEC1.AssetId = Cost_TankOther.AssetId
      )
      LEFT JOIN Cost_Oil ON UnitOEC1.AssetId = Cost_Oil.AssetId
    )
    LEFT JOIN Cost_DecalsMisc ON UnitOEC1.AssetId = Cost_DecalsMisc.AssetId
  )
  LEFT JOIN Cost_OBC_All ON UnitOEC1.AssetId = Cost_OBC_All.AssetId
WHERE
  (
    (
      (UnitOEC1.AssetId) Is Not Null
    )
  );
