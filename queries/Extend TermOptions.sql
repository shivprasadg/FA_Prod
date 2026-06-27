SELECT
  LookUp_ExitOptionTypes.OptionSubType,
  LookUp_ExitOptionTypes.OptionSubTypeDesc
FROM
  LookUp_ExitOptionTypes
WHERE
  (
    (
      (
        LookUp_ExitOptionTypes.OptionTypeID
      )> 2
    )
  );
