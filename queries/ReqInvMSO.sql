SELECT
  Clients.ClientShNm AS Regarding,
  Clients.clientCompanyName AS ClientName,
  Nz([sgrpfaid], [SchFAID]) AS [FAID/(PO)],
  Units.unitvin,
  Units.UnitID AS AssetID,
  Units.UnitYr AS MYr,
  Units.UnitMake AS Make,
  Units.unitmodel AS Model,
  Units.unittype AS Type,
  Units.unitSubType AS SubType,
  Units.UnitEstDelDate AS [Est Delivelry Dt],
  IIf(
    IsNull([unitmsocopy])
    And IsNull([UnitMSORecd]),
    "Needed",
    Null
  ) AS [MSO Copy],
  Nz([unitinvoicenum], "Need") AS [Invoice No],
  Locations.LCity,
  Locations.LCity AS [Delivery City],
  Locations.LState AS [Delivery St],
  Clients.ClientShNm,
  Schedule.SchNo,
  Inv_MSOContactq2.VendorContact,
  Inv_MSOContactq2.contactEmail,
  Inv_MSOContactq2.VendorID,
  IIf(
    IsNull([UnitMSOCopy]),
    "Need",
    IIf(
      IsNull([UnitMSORecd]),
      "Need",
      "Have"
    )
  ) AS GotMSO,
  DateDiff(
    "d",
    Date(),
    [unitestdeldate]
  ) AS Days2Delivery,
  Units.UnitMSORecd,
  SchGrp.PONum1
FROM
  (
    (
      (
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
          LEFT JOIN Units ON SchGrp.SGrpID = Units.GroupREF
        )
        LEFT JOIN lookup_groupTypes ON Units.UnitType = lookup_groupTypes.groupType
      )
      LEFT JOIN CLASS ON lookup_groupTypes.grpClass = CLASS.CLASS
    )
    LEFT JOIN Locations ON Units.UnitGarageLocationRef = Locations.LocationID
  )
  LEFT JOIN Inv_MSOContactq2 ON Units.VendorREF = Inv_MSOContactq2.VendorRef
WHERE
  (
    (
      (
        Nz([unitinvoicenum], "Need")
      )= "Need"
    )
    AND (
      (
        DateDiff(
          "d",
          Date(),
          [unitestdeldate]
        )
      )>-6
    )
    AND (
      (SchGrp.PONum1) Is Not Null
    )
    AND (
      (Clients.clienttype)<> "dead"
    )
    AND (
      (MstrLease.MLOrig)= "FA"
    )
    AND (
      (CLASS.TITLED)= Yes
    )
  )
ORDER BY
  Nz([sgrpfaid], [SchFAID]);
