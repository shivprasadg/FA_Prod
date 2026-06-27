SELECT
  ATF_ChangeReasons.INDX,
  ATF_ChangeReasons.FieldName,
  ATF_ChangeReasons.GroupID,
  ATF_ChangeReasons.SysUser,
  CDate(
    Format(
      Nz(
        [ATF_ChangeReasons].[ChangeDate],
        "1/1/1900"
      ),
      "mm/dd/yyyy"
    )
  ) AS ChangeDate,
  ATF_ChangeReasons.ChangeDate AS ChangeDateTime,
  ATF_ChangeReasons.Reason
FROM
  ATF_ChangeReasons
  INNER JOIN ATF_LastChangeIndx ON ATF_ChangeReasons.INDX = ATF_LastChangeIndx.INDX
GROUP BY
  ATF_ChangeReasons.INDX,
  ATF_ChangeReasons.FieldName,
  ATF_ChangeReasons.GroupID,
  ATF_ChangeReasons.SysUser,
  ATF_ChangeReasons.ChangeDate,
  ATF_ChangeReasons.Reason
HAVING
  (
    (
      (
        CDate(
          Format(
            Nz(
              [ATF_ChangeReasons].[ChangeDate],
              "1/1/1900"
            ),
            "mm/dd/yyyy"
          )
        )
      )= Date()
    )
  )
ORDER BY
  ATF_ChangeReasons.FieldName;
