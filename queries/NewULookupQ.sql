SELECT
  Units.*,
  Clients.clientCompanyName AS CompanyNm,
  Clients.ClientShNm,
  Clients.clientCity,
  Locations.LCity,
  Locations.LState,
  Clients.clientState,
  MstrLease.MLNo,
  Emps.EmpNm AS SalespersonName,
  Vendors.VendorName,
  Vendors.VendorSNm,
  Schedule.SchNo,
  SchGrp.SGrpNO,
  Clients.clientID,
  Locations.LocSubLessee AS DSUB,
  Locations.LocationStreet AS DST1,
  Locations.LAddrLine1 AS Dst2,
  Locations.LCity AS Dcity,
  Locations.LCounty AS DCounty,
  Locations.LState AS Dst,
  Locations.LZip AS Dzip,
  Locations_1.LocSubLessee AS RSub,
  Locations_1.LocationStreet AS Rs1,
  Locations_1.LAddrLine1 AS RS2,
  Locations_1.LCity AS RCity,
  Locations_1.LState AS RSt,
  Locations_1.LZip AS Rzip,
  Locations_2.LState AS Tst,
  Clients.clientID,
  MstrLease.Client,
  MstrLease.MLID,
  Schedule.SchMLRef,
  Schedule.SchID,
  SchGrp.SGrpSchID,
  SchGrp.SGrpID,
  Units.GroupREF AS GroupR,
  IIf(
    [UnitPlateNum] Is Null,
    "",
    Replace(
      Replace([UnitPlateNum], "-", ""),
      " ",
      ""
    )
  ) AS UnitPlateNumH,
  Emps.EmpInt AS SPInitials,
  Clients.ClientGroupID
FROM
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
    INNER JOIN (
      (
        (
          (
            Units
            LEFT JOIN Locations ON Units.UnitGarageLocationRef = Locations.LocationID
          )
          LEFT JOIN Locations AS Locations_1 ON Units.UnitRegLocationRef = Locations_1.LocationID
        )
        LEFT JOIN Locations AS Locations_2 ON Units.UnitTitleLocationRef = Locations_2.LocationID
      )
      LEFT JOIN Vendors ON Units.VendorREF = Vendors.VendorID
    ) ON SchGrp.SGrpID = Units.GroupREF
  )
  INNER JOIN Emps ON Clients.LOinitial = Emps.EmpInt;
