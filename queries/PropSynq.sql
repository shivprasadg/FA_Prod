SELECT
  Clients.ClientShNm,
  MstrLease.MLNo,
  Schedule.SchNo,
  Schedule.SchShortDesc,
  Schedule.SchFAID,
  SchGrp.SGrpNO,
  Schedule.Schsigned,
  SchGrp.BLCD,
  SchGrp.Assignee_Bank,
  SchGrp.FundDate
FROM
  (
    (
      Clients
      INNER JOIN MstrLease ON Clients.clientID = MstrLease.Client
    )
    INNER JOIN Schedule ON MstrLease.MLID = Schedule.SchMLRef
  )
  INNER JOIN SchGrp ON Schedule.SchID = SchGrp.SGrpSchID
WHERE
  (
    (
      (Clients.ClientShNm)= "costco"
      Or (Clients.ClientShNm)= "dyno"
      Or (Clients.ClientShNm)= "afstores"
      Or (Clients.ClientShNm)= "hannaford"
      Or (Clients.ClientShNm)= "conco"
      Or (Clients.ClientShNm) Like "rein*"
    )
    AND (
      (MstrLease.MLNo) Not Like "*demo*"
    )
    AND (
      (MstrLease.MLOrig)= "fa"
    )
  )
  OR (
    (
      (MstrLease.MLNo)= "24030-WF-Rnhrt"
    )
    AND (
      (SchGrp.Assignee_Bank) Like "w*"
    )
  )
ORDER BY
  Clients.ClientShNm,
  Schedule.SchID,
  Schedule.SchNo,
  Schedule.SchFAID,
  SchGrp.SGrpNO,
  Schedule.SchID;
