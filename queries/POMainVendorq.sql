SELECT
  Clients.ClientGroupID,
  Clients.ClientShNm,
  MstrLease.MLNo,
  Schedule.SchNo,
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
  Units.unityr,
  Units.unitmake,
  Units.unitmodel,
  Units.unittype,
  Units.Axle,
  Units.unitfinalcost,
  Childinv.ChildYear,
  Childinv.ChildMake,
  Childinv.ChildClass,
  Childinv.ChildType,
  Childinv.childfinalcost,
  Childinv.childvendorref
FROM
  Clients
  INNER JOIN (
    (
      (
        (
          MstrLease
          INNER JOIN Schedule ON MstrLease.MLID = Schedule.SchMLRef
        )
        INNER JOIN SchGrp ON Schedule.SchID = SchGrp.SGrpSchID
      )
      INNER JOIN Units ON SchGrp.SGrpID = Units.GroupREF
    )
    INNER JOIN Childinv ON Units.UnitID = Childinv.UnitID
  ) ON Clients.clientID = MstrLease.Client
WHERE
  (
    (
      (Childinv.childvendorref)= [vendorref]
    )
    AND (
      (MstrLease.MLOrig)= "fa"
    )
  );
