INSERT INTO Deals (
  clientRef, dealShortDesc, dealDate,
  FAID
)
SELECT
  [Forms]![ClientsView]![clientID] AS clientRef,
  "New Transaction Added " & Now() AS dealShortDesc,
  Now() AS dealDate,
  "FAID" AS FAID;
