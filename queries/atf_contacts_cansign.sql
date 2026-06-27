TRANSFORM First(ATF_ContactsActive.SignerType) AS SignerType
SELECT
  ATF_ContactsActive.ContactID,
  ATF_ContactsActive.FullName,
  ATF_ContactsActive.Title
FROM
  ATF_ContactsActive
GROUP BY
  ATF_ContactsActive.ContactID,
  ATF_ContactsActive.FullName,
  ATF_ContactsActive.Title
ORDER BY
  ATF_ContactsActive.SignerType DESC PIVOT ATF_ContactsActive.SignerType;
