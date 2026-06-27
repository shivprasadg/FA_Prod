SELECT
  LeaseReport1.[Lease No],
  LeaseReport1.Lessee,
  LeaseReport1.Co_Lessee,
  IIf(
    IsDate([SynDate]),
    DatePart("yyyy", [SynDate]),
    0
  ) AS SYDYear,
  IIf(
    [Status] = "B"
    And Not IsDate([UnitInvoicePaid]),
    DatePart(
      "yyyy",
      Nz(
        [UnitEstDelDate],
        Nz([PODate], [UnitCOABack])
      )
    ),
    0
  ) AS OrderedYear,
  IIf(
    [Status] = "B"
    And IsDate([UnitInvoicePaid]),
    DatePart("yyyy", [UnitInvoicePaid]),
    0
  ) AS OwnedYear,
  IIf(
    IsDate([SynDate]),
    DatePart("yyyy", [SynDate]),
    0
  ) AS LeaseYear,
  IIf(
    IsNull([unitcoaback]),
    [OrderedYear],
    CDbl(
      DatePart("yyyy", [unitcoaback])
    )
  ) AS InServYear,
  LeaseReport1.FAID,
  DatePart("m", [MstrLeaseDate]) AS LeaseMonth,
  LeaseReport1.SchNo,
  LeaseReport1.ModelYear,
  LeaseReport1.Make,
  LeaseReport1.UnitType,
  LeaseReport1.Group,
  LeaseReport1.[Unit No],
  LeaseReport1.Unitvin,
  LeaseReport1.Desc,
  LeaseReport1.UnitCost,
  LeaseReport1.UnitInvoiceNum,
  LeaseReport1.UnitInvoicePaid,
  IIf(
    IsNull([RefrigMake]),
    Null,
    [RefrigMake] & " - " & [RefrigSerialNum]
  ) AS RefrigMake_Serial,
  CCur(
    Nz([Cost_Refrig].[Refrig], 0)
  ) AS Refrig,
  Cost_Refrig.RefrigInvoice,
  Cost_Refrig.RefrigPaid,
  Cost_OBC_All.OBC_Make AS OBCMake,
  CCur(
    Nz([Cost_OBC_All].[OBC_Cost], 0)
  ) AS OBC,
  Cost_OBC_All.OBCInvoice,
  Cost_OBC_All.OBCPaid,
  Cost_OBC_All.OBC_Adapt,
  CCur(
    Nz([Cost_OBC_All].[OBCAdaptr], 0)
  ) AS OBCAdapt,
  Cost_OBC_All.OBCAdaptInv,
  Cost_OBC_All.OBCAdaptPaid,
  Cost_OBC_All.OBC_Dome,
  CCur(
    Nz([Cost_OBC_All].[OBCDome], 0)
  ) AS OBCDome,
  Cost_OBC_All.OBCDomeInv,
  Cost_OBC_All.OBCDomePaid,
  Cost_OBC_All.OBC_Multi,
  CCur(
    Nz([Cost_OBC_All].[OBCMulti], 0)
  ) AS OBCMulti,
  Cost_OBC_All.OBCMultiInv,
  Cost_OBC_All.OBCMultiPaid,
  Cost_OBC_All.OBCInstaller,
  CCur(
    Nz([Cost_OBC_All].[OBCInstall], 0)
  ) AS OBCInstall,
  Cost_OBC_All.OBCInsInvoice,
  Cost_OBC_All.OBCInstallPaid,
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
  CCur(
    Nz([Cost_TankBody].[tankbody], 0)
  ) AS TankBody,
  Cost_TankBody.TankMake,
  Cost_TankBody.TankInvoice,
  Cost_TankBody.TankPaid,
  CCur(
    Nz([Cost_TankOther].[tankOther], 0)
  ) AS TankOther,
  Cost_TankOther.TankOthMake,
  Cost_TankOther.TankOthInvoice,
  Cost_TankOther.TankOthPaid,
  CCur(
    Nz([Cost_Oil].[OilRig], 0)
  ) AS OilRig,
  Cost_Oil.OilRigMake,
  Cost_Oil.OilRigInvoice,
  Cost_Oil.OilRigPaid,
  CCur(
    [Refrig] + [OBCTotal] + [FeesTax] + [Freight] + [Warranty] + [Other] + [Inspection] + [PTO] + [CoNameCost] + [RegDecalCost] + [MiscCost] + [TankBody] + [TankOther] + [OilRig]
  ) AS TotalChildCost,
  CCur(
    [UnitCost] + [Refrig] + [OBCTotal] + [FeesTax] + [Freight] + [Warranty] + [Other] + [Inspection] + [PTO] + [CoNameCost] + [RegDecalCost] + [MiscCost] + [TankBody] + [TankOther] + [OilRig]
  ) AS TotalCost,
  LeaseReport1.AssetId,
  LeaseReport1.Status,
  LeaseReport1.InService,
  LeaseReport1.[Title Num],
  IIf(
    Not IsDate([UnitInvoicePaid])
    And Not IsDate([SynDate]),
    "Ordered",
    Null
  ) AS Ordered,
  IIf(
    IsNull([SynDate])
    And IsDate([UnitInvoicePaid]),
    "Owned",
    Null
  ) AS Owned,
  IIf(
    IsDate([SynDate]),
    "Syndicated",
    Null
  ) AS Syndicated
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
                            LeaseReport1
                            LEFT JOIN Cost_FeesTax ON LeaseReport1.AssetId = Cost_FeesTax.AssetId
                          )
                          LEFT JOIN Cost_Freight ON LeaseReport1.AssetId = Cost_Freight.AssetId
                        )
                        LEFT JOIN Cost_Refrig ON LeaseReport1.AssetId = Cost_Refrig.AssetId
                      )
                      LEFT JOIN Cost_Warranty ON LeaseReport1.AssetId = Cost_Warranty.AssetId
                    )
                    LEFT JOIN Cost_Other ON LeaseReport1.AssetId = Cost_Other.AssetId
                  )
                  LEFT JOIN Cost_PTO ON LeaseReport1.AssetId = Cost_PTO.AssetId
                )
                LEFT JOIN Cost_Inspection ON LeaseReport1.AssetId = Cost_Inspection.AssetId
              )
              LEFT JOIN Cost_DecalsCoName ON LeaseReport1.AssetId = Cost_DecalsCoName.AssetId
            )
            LEFT JOIN Cost_DecalsRegNumber ON LeaseReport1.AssetId = Cost_DecalsRegNumber.AssetId
          )
          LEFT JOIN Cost_DecalsMisc ON LeaseReport1.AssetId = Cost_DecalsMisc.DecalVendor
        )
        LEFT JOIN Cost_TankBody ON LeaseReport1.AssetId = Cost_TankBody.AssetId
      )
      LEFT JOIN Cost_TankOther ON LeaseReport1.AssetId = Cost_TankOther.AssetId
    )
    LEFT JOIN Cost_Oil ON LeaseReport1.AssetId = Cost_Oil.AssetId
  )
  LEFT JOIN Cost_OBC_All ON LeaseReport1.AssetId = Cost_OBC_All.AssetId
WHERE
  (
    (
      (LeaseReport1.UnitType) Not Like "*Warranty*"
    )
  );
