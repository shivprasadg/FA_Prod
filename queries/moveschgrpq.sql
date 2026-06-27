UPDATE
  SchGrp
SET
  SchGrp.SGrpSchID = [forms]![movegrp]![newsch]
WHERE
  (
    (
      (SchGrp.SGrpID)= [forms]![movegrp]![group]
    )
  );
