SELECT
  LookUp_ChildPartType.ChildPartTypeID AS MasterTypeID,
  LookUp_ChildPartType.ChildPartType,
  LookUp_ChildPartMake.ChildPartTypeID,
  LookUp_ChildPartMake.ChildPartMake,
  LookUp_ChildPartModel.ChildPartMakeID,
  LookUp_ChildPartModel.ChildPartModel,
  LookUp_ChildPartModel.ChildPartModelID
FROM
  (
    LookUp_ChildPartMake
    INNER JOIN LookUp_ChildPartType ON LookUp_ChildPartMake.ChildPartTypeID = LookUp_ChildPartType.ChildPartTypeID
  )
  LEFT JOIN LookUp_ChildPartModel ON LookUp_ChildPartMake.ChildPartMakeID = LookUp_ChildPartModel.ChildPartMakeID
ORDER BY
  LookUp_ChildPartType.ChildPartType,
  LookUp_ChildPartMake.ChildPartMake,
  LookUp_ChildPartModel.ChildPartModel;
