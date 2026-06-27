UPDATE
  Clients
  INNER JOIN Contacts ON Clients.clientID = Contacts.clientRef
SET
  Contacts.contactAddress1 = [clients].[clientAddress1],
  Contacts.contactAddress2 = [clients].[clientAddress2],
  Contacts.contactCity = [clients].[clientcity],
  Contacts.LState = [clients].[clientstate],
  Contacts.contactZip = [clients].[clientzip],
  Contacts.contactCountry = [clients].[clientcountry]
WHERE
  (
    (
      (Contacts.clientRef)= Forms!Contacts_view!clientRef
    )
    And (
      (Contacts.contactID)= Forms!Contacts_view!contactID
    )
  );
