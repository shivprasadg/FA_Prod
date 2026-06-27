SELECT
  Banks.BankShortName,
  [BankName] & Chr(13)& Chr(10)& [Address1] & Chr(13)& Chr(10)& IIf(
    IsNull([Address2]),
    Null,
    [Banks].[Address2] & Chr(13)& Chr(10)
  )& [Banks].[City] & " " & [Banks].[state] & " " & Left([Banks].[Zip], 5) AS BankAddress,
  Banks.BankId,
  Banks.Phone,
  Banks.PortfolioNumber,
  Banks.Fax
FROM
  Banks;
