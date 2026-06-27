SELECT
  IIf(
    IsNull([sgrpfaid]),
    [schfaid],
    [sgrpfaid]
  ) AS LFAID,
  MstrLease.MLOrig,
  Sum(SchGrp.OrigEquipCost) AS OrigEquipCost,
  Sum(SchGrp.UnitsInGroup) AS UnitsInGroup
FROM
  (
    (
      Clients
      INNER JOIN MstrLease ON Clients.clientID = MstrLease.Client
    )
    INNER JOIN Schedule ON MstrLease.MLID = Schedule.SchMLRef
  )
  INNER JOIN SchGrp ON Schedule.SchID = SchGrp.SGrpSchID
GROUP BY
  IIf(
    IsNull([sgrpfaid]),
    [schfaid],
    [sgrpfaid]
  ),
  MstrLease.MLOrig
HAVING
  (
    (
      (
        IIf(
          IsNull([sgrpfaid]),
          [schfaid],
          [sgrpfaid]
        )
      ) Is Not Null
    )
    AND (
      (MstrLease.MLOrig)= "fa"
    )
  );
