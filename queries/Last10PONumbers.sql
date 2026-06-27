SELECT
  TOP 10 SchGrp.PONum1,
  SchGrp.SGrpFAID,
  CInt(
    splitstring([PONum1], "-", 2)
  ) AS PoSeq
FROM
  SchGrp
GROUP BY
  SchGrp.PONum1,
  SchGrp.SGrpFAID
HAVING
  (
    (
      (SchGrp.PONum1) Is Not Null
    )
  )
ORDER BY
  SchGrp.SGrpFAID DESC;
