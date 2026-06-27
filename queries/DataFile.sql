SELECT
  CTT.[FA TitleLog Num] AS [FA-Asset Num],
  CTT.VIN AS Expr1,
  CTT.[File Type] AS Expr2,
  CTT.[Portfolio Certificate No] AS Expr3,
  CTT.[Lender or Holder Name] AS Expr4
FROM
  TList AS CTT
ORDER BY
  [FA TitleLog Num];
