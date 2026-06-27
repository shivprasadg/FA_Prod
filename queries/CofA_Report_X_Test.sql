SELECT
  vw_CofABase.*
FROM
  vw_CofABase
WHERE
  (
    (
      (vw_CofABase.SGrpID)= [Forms]![BulkChange_Form]![GrpID]
    )
  );
