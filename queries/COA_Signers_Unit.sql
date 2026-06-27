SELECT
  "COA_Signer" AS SignerType,
  SchGrp.SGrpSchID AS SchID,
  SchGrp.SGrpID,
  Units.CofASigner AS CofASignerID,
  SchGrp.SGrpNO
FROM
  Units
  INNER JOIN SchGrp ON Units.GroupREF = SchGrp.SGrpID
GROUP BY
  SchGrp.SGrpSchID,
  SchGrp.SGrpID,
  Units.CofASigner,
  SchGrp.SGrpNO
HAVING
  (
    (
      (Units.CofASigner) Is Not Null
      And (Units.CofASigner)> 0
    )
  );
