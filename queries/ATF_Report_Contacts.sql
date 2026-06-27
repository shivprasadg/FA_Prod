SELECT
  ATF_Contacts.ATFID,
  ATF_Contacts.ContactType,
  ATF_Contacts.FullName,
  ATF_Contacts.EmailYN,
  ATF_Contacts.contactPrimaryPhone AS Phone,
  ATF_Contacts.contactCell AS Mobile
FROM
  ATF_Contacts;
