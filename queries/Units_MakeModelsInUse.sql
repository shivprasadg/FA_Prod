SELECT
  Units_MakeModelsInUse_A.Make AS UnitMake,
  Units_MakeModelsInUse_A.Model AS UnitModel,
  Max(Units_MakeModelsInUse_A.Leased) AS Leased,
  Max(
    Units_MakeModelsInUse_A.Tracked
  ) AS Tracked
FROM
  Units_MakeModelsInUse_A
GROUP BY
  Units_MakeModelsInUse_A.Make,
  Units_MakeModelsInUse_A.Model;
