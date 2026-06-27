SELECT
  Vendors.VendorName,
  Vendors.VendorAddress,
  [VendorCity] & " " & [vendorstate] & " " & [vendorzip] AS VendorCst,
  Vendors.VendorPhone,
  Count(Units.UnitID) AS CountOfUnitID,
  [UnitYr] & " " & [UnitMake] & " " & [UnitModel] & " " & [UnitSubtype] & " " & [UnitType] & IIf([axle] = "TBD", Null, " " & [axle]) AS Unit,
  Clients.clientCompanyName,
  Min(Units.UnitActualDelDate) AS [Min Del Date],
  Max(Units.UnitActualDelDate) AS [Max Del Date],
  Schedule.schfaid AS FAID,
  Locations.LState,
  Locations.LCity,
  Locations.LAddrLine1,
  Locations.LocationStreet,
  Locations.LZip,
  Units.UnitFinalCost,
  Locations.LocationID,
  Min(Units.UnitEstDelDate) AS MinOfUnitEstDelDate,
  Max(Units.Unitestdeldate) AS MaxOfUnitestdeldate,
  Locations.LocSubLessee,
  1 AS rptsec,
  SchGrp.PONum1,
  Schedule.SchNo,
  SchGrp.RebateUnitProjected,
  SchGrp.POQuoteNums,
  [Attention To] AS AttentionTo,
  DLookUp(
    "ExtendedCost", "PO_ChildParts",
    "PoNumber='" & [PONum1] & "'"
  ) AS ChildCost
FROM
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
        INNER JOIN Units ON SchGrp.SGrpID = Units.GroupREF
      )
      LEFT JOIN Locations ON Units.UnitGarageLocationRef = Locations.LocationID
    )
    LEFT JOIN Vendors ON Units.VendorREF = Vendors.VendorID
  )
  LEFT JOIN lookup_groupTypes ON Units.UnitType = lookup_groupTypes.groupType
GROUP BY
  Vendors.VendorName,
  Vendors.VendorAddress,
  [VendorCity] & " " & [vendorstate] & " " & [vendorzip],
  Vendors.VendorPhone,
  [UnitYr] & " " & [UnitMake] & " " & [UnitModel] & " " & [UnitSubtype] & " " & [UnitType] & IIf([axle] = "TBD", Null, " " & [axle]),
  Clients.clientCompanyName,
  Schedule.schfaid,
  Locations.LState,
  Locations.LCity,
  Locations.LAddrLine1,
  Locations.LocationStreet,
  Locations.LZip,
  Units.UnitFinalCost,
  Locations.LocationID,
  Locations.LocSubLessee,
  1,
  SchGrp.PONum1,
  Schedule.SchNo,
  SchGrp.RebateUnitProjected,
  SchGrp.POQuoteNums,
  MstrLease.MLOrig,
  SchGrp.BLCD,
  SchGrp.sgrpfaid,
  Units.unitfaid
HAVING
  (
    (
      (SchGrp.PONum1)= Forms!GroupForm!PONum1
    )
    And (
      (MstrLease.MLOrig)= "fa"
    )
  )
ORDER BY
  Locations.LState,
  Locations.LCity,
  Locations.LocationStreet,
  Units.UnitFinalCost;
