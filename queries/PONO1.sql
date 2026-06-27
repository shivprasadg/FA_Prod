SELECT
  Count(Clients.ClientShNm) AS CountOfClientShNm,
  MstrLease.MLOrig,
  IIf(
    IsNull([unitfaid]),
    IIf(
      IsNull([sgrpfaid]),
      [schfaid],
      [sgrpfaid]
    ),
    [unitfaid]
  ) AS FAID,
  Units.VendorREF,
  [unityr] & " " & [unitmake] & " " & [unitmodel] & " " & [axle] AS [Main Unit],
  Locations.LocationName,
  Locations.LocationStreet,
  Locations.LAddrLine1,
  Locations.LAddrLine2,
  Locations.LAddrLine3,
  Locations.LCity,
  Locations.LState,
  Units.unitfinalcost,
  Locations.LZip
FROM
  Locations
  RIGHT JOIN (
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
  ) ON Locations.LocationID = Units.UnitDeliveryLocationRef
GROUP BY
  MstrLease.MLOrig,
  IIf(
    IsNull([unitfaid]),
    IIf(
      IsNull([sgrpfaid]),
      [schfaid],
      [sgrpfaid]
    ),
    [unitfaid]
  ),
  Units.VendorREF,
  [unityr] & " " & [unitmake] & " " & [unitmodel] & " " & [axle],
  Locations.LocationName,
  Locations.LocationStreet,
  Locations.LAddrLine1,
  Locations.LAddrLine2,
  Locations.LAddrLine3,
  Locations.LCity,
  Locations.LState,
  Units.unitfinalcost,
  Locations.LZip
HAVING
  (
    (
      (MstrLease.MLOrig)= "fa"
    )
    AND (
      (
        IIf(
          IsNull([unitfaid]),
          IIf(
            IsNull([sgrpfaid]),
            [schfaid],
            [sgrpfaid]
          ),
          [unitfaid]
        )
      )= "13-545"
    )
  );
