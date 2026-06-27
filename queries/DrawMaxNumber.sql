SELECT
  DrawDowns.DrawRequestNumber,
  Max(DrawDowns.DrawNumber) AS DrawNumber
FROM
  DrawDowns
GROUP BY
  DrawDowns.DrawRequestNumber;
