SELECT
  Units_MakeModelsInUse.UnitMake,
  Units_MakeModelsInUse.UnitModel,
  Units_MakeModelsInUse.Leased,
  Units_MakeModelsInUse.Tracked
FROM
  Units_MakeModelsInUse
  LEFT JOIN LookUp_UnitModels ON Units_MakeModelsInUse.UnitModel = LookUp_UnitModels.UnitModel
WHERE
  (
    (
      (LookUp_UnitModels.UnitModel) Is Null
    )
  )
ORDER BY
  Units_MakeModelsInUse.UnitMake,
  Units_MakeModelsInUse.UnitModel;
