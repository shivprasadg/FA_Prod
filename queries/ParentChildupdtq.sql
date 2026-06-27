SELECT
  MstrLease.MLNo,
  MstrLease.MLOrig,
  Schedule.SchID,
  Schedule.SchNo,
  ChildParts.ChildMake,
  ChildParts.ChildType,
  ChildParts.ChildInvoiceNum,
  ChildParts.childFinalCost,
  Units.UnitFinalCost,
  Units.UnitInvoiceNum
FROM
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
  INNER JOIN ChildParts ON Units.UnitID = ChildParts.UnitREF
WHERE
  (
    (
      (MstrLease.MLNo) Like "costco*"
    )
    AND (
      (Schedule.SchNo) Like "29*"
    )
    AND (
      (ChildParts.ChildType)= "freight"
    )
  );
