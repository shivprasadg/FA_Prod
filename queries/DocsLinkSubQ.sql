SELECT
  "" & Chr(13)& Chr(10)& IIf(
    [dtype] = "DIR", "All", [G].[SGrpNO]
  ) AS sGrpNo,
  DocLinks.*
FROM
  DocLinks
  LEFT JOIN (
SELECT
  sGrpNo, SGrpId
FROM
  SchGrp) AS G ON DocLinks.GroupID = G.SGrpID;
