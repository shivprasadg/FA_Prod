SELECT
  FALeased_Unitcount.ClientShortName,
  CDbl(
    Nz([UnitOrdered], 0)
  ) AS Ordered,
  CDbl(
    Nz([UnitTitled], 0)
  ) AS Titled,
  CDbl(
    Nz([UnitDelivered], 0)
  ) AS Delivered,
  CDbl(
    Nz([UnitAccepted], 0)
  ) AS Accepted,
  CDbl(
    Nz([UnitsInServiced], 0)
  ) AS InServiced,
  CDbl(
    Nz([UnitsOffLease], 0)
  ) AS OffLease,
  [ByYear_UnitsOrdered-PO].FirstPO,
  [ByYear_UnitsOrdered-PO].LastPO,
  ByYear_UnitsInServiced.LastInServiced,
  ByYear_UnitsOffLease.LastOL AS LastOffLease,
  FALeased_Unitcount.ClientGroupId,
  FALeased_Unitcount.clientID
FROM
  (
    (
      (
        (
          (
            FALeased_Unitcount
            LEFT JOIN [ByYear_UnitsOrdered-PO] ON FALeased_Unitcount.clientID = [ByYear_UnitsOrdered-PO].clientID
          )
          LEFT JOIN ByYear_UnitsDelivered ON FALeased_Unitcount.clientID = ByYear_UnitsDelivered.clientID
        )
        LEFT JOIN ByYear_UnitsTitled ON FALeased_Unitcount.clientID = ByYear_UnitsTitled.clientID
      )
      LEFT JOIN ByYear_UnitsAccepted ON FALeased_Unitcount.clientID = ByYear_UnitsAccepted.clientID
    )
    LEFT JOIN ByYear_UnitsOffLease ON FALeased_Unitcount.clientID = ByYear_UnitsOffLease.clientID
  )
  LEFT JOIN ByYear_UnitsInServiced ON FALeased_Unitcount.clientID = ByYear_UnitsInServiced.clientID
ORDER BY
  FALeased_Unitcount.ClientShortName;
