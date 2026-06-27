SELECT
  HrsOnlyq.Unitref AS Expr1,
  Max(HrsOnlyq.[Reefer Hours]) AS [MaxOfReefer Hours],
  Max(HrsOnlyq.[Close DAte]) AS [MaxOfClose DAte]
FROM
  HrsOnlyq
GROUP BY
  HrsOnlyq.Unitref;
