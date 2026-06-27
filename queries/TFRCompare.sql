SELECT
  *
FROM
  TFR_Work
WHERE
  Assigned_Indicator = 'I'
  OR Assigned_Indicator <> ''
  AND FundDate >= #1/1/2024#
  AND FundDate <= #12/31/2024#
  Or FundDate Is Null;
