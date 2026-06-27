SELECT
  Leased_MLSchGrpUnits.MLID,
  Leased_MLSchGrpUnits.MLNo,
  Leased_MLSchGrpUnits.MLOrig,
  Leased_MLSchGrpUnits.Client
FROM
  LeasedClientsWUnits
  INNER JOIN Leased_MLSchGrpUnits ON LeasedClientsWUnits.clientID = Leased_MLSchGrpUnits.Client
WHERE
  (
    (
      (Leased_MLSchGrpUnits.MLOrig)= "FA"
    )
  )
ORDER BY
  Leased_MLSchGrpUnits.MLNo;
