SELECT
  vw_SixKeys.ClientGroupId,
  vw_SixKeys.clientID,
  vw_SixKeys.CompanyName,
  vw_SixKeys.Schedule,
  vw_SixKeys.UnitGroup,
  vw_SixKeys.UnitID,
  vw_SixKeys.Unitnum,
  vw_SixKeys.Status,
  vw_SixKeys.MYear,
  vw_SixKeys.Make,
  vw_SixKeys.Model
FROM
  vw_SixKeys
WHERE
  (
    (
      (vw_SixKeys.Unitnum) Is Null
      Or (vw_SixKeys.Unitnum)= ""
    )
    AND (
      (vw_SixKeys.Status)<> "B"
    )
  );
