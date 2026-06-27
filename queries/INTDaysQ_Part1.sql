SELECT
  MstrLease.MLNo,
  Schedule.SchNo,
  SchGrp.SGrpID,
  MstrLease.MLID,
  Schedule.schcrexp,
  SchGrp.sgrpnoslv
FROM
  (
    MstrLease
    INNER JOIN Schedule ON MstrLease.MLID = Schedule.SchMLRef
  )
  INNER JOIN SchGrp ON Schedule.SchID = SchGrp.SGrpSchID
WHERE
  (
    (
      (MstrLease.MLNo)= Forms!Billing!MLNo
    )
    And (
      (Schedule.SchNo)= Forms!Billing!SchNo
    )
  )
ORDER BY
  MstrLease.MLNo;
