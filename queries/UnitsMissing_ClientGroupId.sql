UPDATE
  (
    (
      (
        SchGrp
        INNER JOIN Schedule ON SchGrp.SGrpSchID = Schedule.SchID
      )
      INNER JOIN MstrLease ON Schedule.SchMLRef = MstrLease.MLID
    )
    INNER JOIN Clients ON MstrLease.Client = Clients.clientID
  )
  INNER JOIN Units ON SchGrp.SGrpID = Units.GroupREF
SET
  Units.ClientGroupId = [Clients].[ClientGroupId],
  Units.clientId = [Clients].[clientID]
WHERE
  (
    (
      (Units.ClientGroupId) Is Null
    )
  )
  OR (
    (
      (Units.clientId) Is Null
    )
  );
