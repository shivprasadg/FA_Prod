SELECT
  Clients.ClientShNm,
  Schedule.SchNo,
  Schedule.SchFSL,
  FSL.ID,
  Schedule.SchShortDesc,
  FSL.FSLYR1,
  FSL.FSLYR2,
  FSL.FSLYR3,
  FSL.FSLYR4,
  FSL.FSLYR5,
  FSL.FSLYR6,
  FSL.FSLYR7
FROM
  (
    Clients
    INNER JOIN MstrLease ON Clients.clientID = MstrLease.Client
  )
  INNER JOIN (
    Schedule
    LEFT JOIN FSL ON Schedule.FSLSCH = FSL.ID
  ) ON MstrLease.MLID = Schedule.SchMLRef
WHERE
  (
    (
      (Clients.ClientShNm) Like "costco*"
    )
  );
