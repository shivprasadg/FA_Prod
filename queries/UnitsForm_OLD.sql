SELECT
  Units.*,
  Clients.ClientShNm,
  MstrLease.MLNo,
  MstrLease.MLOrig,
  MstrLease.MLLessor,
  Schedule.SchNo,
  SchGrp.SGrpNO,
  Clients.clientID,
  Schedule.SchID,
  SchGrp.SGrpID,
  Locations.LocSubLessee AS GSUB,
  Locations.LocationStreet AS GST1,
  Locations.LAddrLine1 AS Gst2,
  Locations.LCity AS Gcity,
  Locations.LCounty AS GCounty,
  Locations.LState AS Gst,
  Locations.LZip AS Gzip,
  Locations_1.LocSubLessee AS RSub,
  Locations_1.LocationStreet AS Rs1,
  Locations_1.LAddrLine1 AS RS2,
  Locations_1.LCity AS RCity,
  Locations_1.LState AS RSt,
  Locations_1.LZip AS Rzip,
  Locations_2.LState AS Tst,
  Locations.LTaxRate,
  lookup_States.StTaxType,
  SchGrp.BLCD,
  Locations_3.LocationStreet AS DsT1,
  Locations_3.LAddrLine1 AS DsT2,
  Locations_3.LAddrLine2 AS DsT3,
  Locations_3.LCity AS DCity,
  Locations_3.LState AS DSt,
  Locations_3.LZip AS DZip,
  Locations_3.LocSubLessee AS DSub,
  Locations.LTaxOption AS GTaxOpt,
  Schedule.SchID,
  DLookUp(
    "OffLeaseProj", "vw_sixkeys", "AssetID=" & [UnitID]
  ) AS ProjOffLease,
  DateAdd(
    "m", [SchGrp].[LeaseTermPrimary],
    [SchGrp].[BLCD]
  )-1 AS LeaseExp
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
            (
              Units
              LEFT JOIN Locations ON Units.UnitGarageLocationRef = Locations.LocationID
            )
            LEFT JOIN Locations AS Locations_1 ON Units.UnitRegLocationRef = Locations_1.LocationID
          )
          LEFT JOIN Locations AS Locations_2 ON Units.UnitTitleLocationRef = Locations_2.LocationID
        )
        LEFT JOIN lookup_States ON Locations.LState = lookup_States.StateAbbrev
      )
      LEFT JOIN Locations AS Locations_3 ON Units.UnitDeliveryLocationRef = Locations_3.LocationID
    ) ON SchGrp.SGrpID = Units.GroupREF
  )
  INNER JOIN UnitsFormFilter ON Units.UnitID = UnitsFormFilter.AssetIDx;
