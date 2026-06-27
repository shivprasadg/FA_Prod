SELECT
  *
FROM
  vw_FAID_RevRentGrouped AS X
  LEFT JOIN ATF_ChangeReasons ON (
    X.GroupID = ATF_ChangeReasons.GroupID
  )
  AND (
    X.FieldName = ATF_ChangeReasons.FieldName
  )
  AND (
    X.SysUser = ATF_ChangeReasons.SysUser
  )
WHERE
  (
    (
      (X.GroupID)= 2171
    )
    AND (
      (ATF_ChangeReasons.INDX) Is Null
    )
  );
