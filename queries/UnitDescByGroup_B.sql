SELECT
  UnitDescByGroup_A.GroupREF AS GroupID,
  Count(UnitDescByGroup_A.GroupREF) AS GroupCount,
  Max(UnitDescByGroup_A.Desc) AS ADesc,
  Min(UnitDescByGroup_A.Desc) AS BDesc,
  IIf(
    [GroupCount] > 1, "(" & [AUnits] & ") " & [ADesc] & " (" & [BUnits] & ") " & [BDesc],
    "(" & [AUnits] & ") " & [ADesc]
  ) AS GroupDesc,
  Max(UnitDescByGroup_A.Units) AS AUnits,
  Min(UnitDescByGroup_A.Units) AS BUnits INTO UnitDescByGroupTemp
FROM
  UnitDescByGroup_A
GROUP BY
  UnitDescByGroup_A.GroupREF;
