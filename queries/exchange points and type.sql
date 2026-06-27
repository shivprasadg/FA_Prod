SELECT
  MstrLease.MLOrig,
  MstrLease.MLNo,
  Schedule.SchID,
  Schedule.SchNo,
  SchGrp.SGrpNO,
  SchGrp.LeaseTermPrimary,
  SchGrp.ExchangeITGroup,
  SchGrp.ExchangeITMonth,
  SchGrp.ExchangeITType
FROM
  (
    MstrLease
    INNER JOIN Schedule ON MstrLease.MLID = Schedule.SchMLRef
  )
  INNER JOIN SchGrp ON Schedule.SchID = SchGrp.SGrpSchID
WHERE
  (
    (
      (MstrLease.MLOrig)= "fa"
    )
    AND (
      (MstrLease.MLNo) Like "main*"
    )
  );
