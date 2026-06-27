SELECT
  [costco repairsNew].[Reefer Hours] AS Expr1,
  [costco repairsNew].[Close DAte] AS Expr2,
  [costco repairsNew].[Unit #] AS UnitNo,
  [costco repairsNew].Unitref AS Expr3
FROM
  [costco repairsNew]
WHERE
  (
    (
      (
        [costco repairsNew].[Reefer Hours]
      )> 0
    )
  );
