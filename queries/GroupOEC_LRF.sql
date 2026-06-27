SELECT
  UnitOEC2.GroupREF AS GroupId,
  Sum(UnitOEC2.TotalCost) AS OEC,
  Sum(UnitOEC2.LORent) AS GroupLoRent,
  Sum(UnitOEC2.SYNDRent) AS GroupSyndRent,
  GetLFR([GroupLORent], [OEC]) AS GroupLOLRF,
  GetLFR([GroupSyndRent], [OEC]) AS GroupSyndLRF
FROM
  UnitOEC2
GROUP BY
  UnitOEC2.GroupREF;
