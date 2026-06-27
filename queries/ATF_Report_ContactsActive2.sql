SELECT
  ATF_ContactsActive.Sorter AS Sorter,
  ATF_ContactsActive.ClientID,
  ATF_ContactsActive.MLID,
  ATF_ContactsActive.SignerType AS SignerType,
  ATF_ContactsActive.SchID,
  ATF_ContactsActive.ContactID,
  ATF_ContactsActive.FullName,
  ATF_ContactsActive.Title,
  ATF_ContactsActive.Phone,
  ATF_ContactsActive.Mobile,
  ATF_ContactsActive.HasEmail
FROM
  ATF_Report_Active_ContactControl
  INNER JOIN ATF_ContactsActive ON (
    ATF_Report_Active_ContactControl.ContactID = ATF_ContactsActive.ContactID
  )
  AND (
    ATF_Report_Active_ContactControl.SchID = ATF_ContactsActive.SchID
  )
  AND (
    ATF_Report_Active_ContactControl.Sorter = ATF_ContactsActive.Sorter
  )
ORDER BY
  ATF_ContactsActive.Sorter;
