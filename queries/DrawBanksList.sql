SELECT
  Banks.BankId,
  Banks.BankShortName AS Bank,
  IIf(
    Count([drawdowns].[DrawdownId])> 0,
    "Yes",
    "No"
  ) AS Draws
FROM
  Banks
  LEFT JOIN DrawDowns ON Banks.BankId = DrawDowns.DrawBankId
GROUP BY
  Banks.BankId,
  Banks.BankShortName
ORDER BY
  Banks.BankShortName;
