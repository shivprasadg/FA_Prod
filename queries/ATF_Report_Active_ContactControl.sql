SELECT
  ATF_Report_ContactsActive.SchID,
  Min(
    ATF_Report_ContactsActive.Sorter
  ) AS Sorter,
  ATF_Report_ContactsActive.ContactID AS ContactID
FROM
  ATF_Report_ContactsActive
GROUP BY
  ATF_Report_ContactsActive.SchID,
  ATF_Report_ContactsActive.ContactID;
