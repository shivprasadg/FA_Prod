SELECT
  ATFRevT.schref,
  ATFRevT.RevNo,
  ATFRevT.RevDt,
  ATFRevT.ReviewDt
FROM
  ATFRevT
WHERE
  (
    (
      (ATFRevT.RevNo) Is Not Null
    )
    AND (
      (ATFRevT.ReviewDt) Is Null
    )
  );
