SELECT
  ATF_ContactsActive_Base.ClientID,
  ATF_ContactsActive_Base.MLID,
  ATF_ContactsActive_Base.SignerType,
  ATF_ContactsActive_Base.SchID,
  ATF_ContactsActive_Base.ContactID,
  ClientContacts_Base.FullName,
  ClientContacts_Base.Title,
  ClientContacts_Base.Phone,
  ClientContacts_Base.Mobile,
  ClientContacts_Base.HasEmail,
  IIf(
    [SignerType] = "Master Lease",
    1,
    IIf(
      [SignerType] = "Guarantor",
      2,
      IIf(
        [SignerType] = "Schedule",
        3,
        IIf([SignerType] = "COA", 4, 10)
      )
    )
  ) AS Sorter
FROM
  ATF_ContactsActive_Base
  INNER JOIN ClientContacts_Base ON ATF_ContactsActive_Base.ContactID = ClientContacts_Base.contactID
GROUP BY
  ATF_ContactsActive_Base.ClientID,
  ATF_ContactsActive_Base.MLID,
  ATF_ContactsActive_Base.SignerType,
  ATF_ContactsActive_Base.SchID,
  ATF_ContactsActive_Base.ContactID,
  ClientContacts_Base.FullName,
  ClientContacts_Base.Title,
  ClientContacts_Base.Phone,
  ClientContacts_Base.Mobile,
  ClientContacts_Base.HasEmail;
