SELECT
  Max(MLAmends.ID) AS ID,
  Max(MLAmends.MLAmdDt) AS MLAmdDt,
  Max(MLAmends.MLAmtNO) AS MLAmtNO,
  MLAmends.MLRef
FROM
  MLAmends
GROUP BY
  MLAmends.MLRef;
