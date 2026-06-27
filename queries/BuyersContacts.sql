SELECT
  Contacts.*,
  [contactFirst] & " " & [contactLast] AS Contact
FROM
  Buyers
  INNER JOIN Contacts ON Buyers.clientID = Contacts.clientRef;
