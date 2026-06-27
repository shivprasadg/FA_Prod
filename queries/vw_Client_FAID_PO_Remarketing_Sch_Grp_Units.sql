SELECT
  vw_SixKeys.*
FROM
  vw_SixKeys
WHERE
  (
    (
      (vw_SixKeys.Status) Like "*S*"
      Or (vw_SixKeys.Status) Like "*K*"
      Or (vw_SixKeys.Status) Like '*T*'
      Or (vw_SixKeys.Status) Like '*R*'
    )
  );
