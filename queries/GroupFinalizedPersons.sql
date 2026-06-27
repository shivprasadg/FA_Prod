SELECT
  GroupFinalizedTrackerFields.FieldName,
  GroupFinalizedTrackerFields.FieldLabel,
  DepartmentContacts.Department,
  TBL_Employees.[Employee Name] AS Person,
  [User ID] & "@fleetadvantage.com" AS email
FROM
  (
    GroupFinalizedTrackerFields
    INNER JOIN DepartmentContacts ON GroupFinalizedTrackerFields.DepartmentContactID = DepartmentContacts.INDX
  )
  INNER JOIN TBL_Employees ON DepartmentContacts.PrimaryContactID = TBL_Employees.ID;
