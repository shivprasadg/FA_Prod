SELECT
  ReMarketing_MLSchGrpUnits.MLID,
  ReMarketing_MLSchGrpUnits.MLNo,
  ReMarketing_MLSchGrpUnits.MLOrig,
  ReMarketing_MLSchGrpUnits.Client
FROM
  RemarketingClientsWUnits
  INNER JOIN ReMarketing_MLSchGrpUnits ON RemarketingClientsWUnits.ClientID = ReMarketing_MLSchGrpUnits.Client
ORDER BY
  ReMarketing_MLSchGrpUnits.MLNo;
