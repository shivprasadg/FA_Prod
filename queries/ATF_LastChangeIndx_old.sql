SELECT
  Max(ATF_ChangeReasons.Indx) AS INDX,
  ATF_ChangeReasons.FieldName,
  ATF_ChangeReasons.GroupID,
  ATF_ChangeReasons.SysUser
FROM
  ATF_ChangeReasons
GROUP BY
  ATF_ChangeReasons.FieldName,
  ATF_ChangeReasons.GroupID,
  ATF_ChangeReasons.SysUser
HAVING
  (
    (
      (ATF_ChangeReasons.FieldName) Is Not Null
    )
  )
ORDER BY
  ATF_ChangeReasons.FieldName;
