SELECT
  ATF.*,
  "FAID: " & [FAID] & " LPAF Date: " & [LPAFDate] AS ATFDesc,
  Clients.clientCompanyName AS Lessee,
  Clients_1.clientCompanyName AS CoLessee
FROM
  (
    ATF
    INNER JOIN Clients ON ATF.LesseeID = Clients.clientID
  )
  LEFT JOIN Clients AS Clients_1 ON ATF.CoLesseeID = Clients_1.clientID
WHERE
  (
    (
      (ATF.ATFID) Is Not Null
    )
  );
