SELECT
  Clients.ClientShNm,
  MstrLease.MLNo,
  Schedule.SchNo,
  SchGrp.SGrpNO,
  Schedule.Schsigned,
  SchGrp.BLCD,
  SchGrp.FundDate
FROM
  Clients
  INNER JOIN (
    (
      MstrLease
      INNER JOIN Schedule ON MstrLease.MLID = Schedule.SchMLRef
    )
    INNER JOIN SchGrp ON Schedule.SchID = SchGrp.SGrpSchID
  ) ON Clients.clientID = MstrLease.Client
WHERE
  (
    (
      (Clients.ClientShNm)= "costco"
      Or (Clients.ClientShNm)= "dyno"
      Or (Clients.ClientShNm)= "afstores"
      Or (Clients.ClientShNm)= "hannaford"
      Or (Clients.ClientShNm)= "conco"
    )
    AND (
      (MstrLease.MLNo) Not Like "*demo*"
    )
    AND (
      (MstrLease.MLOrig)= "fa"
    )
  )
ORDER BY
  Clients.ClientShNm,
  Schedule.SchID;
