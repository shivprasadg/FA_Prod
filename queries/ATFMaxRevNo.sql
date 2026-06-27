SELECT
  ATFRevT.schref AS SchID,
  Max(ATFRevT.RevNo) AS RevNo,
  ATFRevT.RevDt
FROM
  ATFRevT
GROUP BY
  ATFRevT.schref,
  ATFRevT.RevDt;
