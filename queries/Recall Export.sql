SELECT
  Clients.clientCompanyName AS Company,
  RecallCampaign.RecallNo AS [No],
  RecallCampaign.Type,
  RecallCampaign.RecallDesc AS Description,
  RecallCampaign.RecallDt AS [Notice Dt],
  Units.UnitUnitNum AS [Unit No],
  Units.unitvin AS VIIN,
  Locations.LocationName,
  Null AS [Action Needed?],
  Recalls.RCRepDt AS [Action Completed( Dt)],
  Units.UnitEngineSerial AS [Eng Serial No]
FROM
  (
    (
      (
        Clients
        INNER JOIN MstrLease ON Clients.clientID = MstrLease.Client
      )
      INNER JOIN Schedule ON MstrLease.MLID = Schedule.SchMLRef
    )
    INNER JOIN SchGrp ON Schedule.SchID = SchGrp.SGrpSchID
  )
  INNER JOIN (
    (
      RecallCampaign
      INNER JOIN (
        (
          Units
          INNER JOIN Locations ON Units.UnitGarageLocationRef = Locations.LocationID
        )
        INNER JOIN Recalls ON Units.UnitID = Recalls.RCUnitID
      ) ON RecallCampaign.ID = Recalls.recallref
    )
    LEFT JOIN Contacts ON Recalls.RCCont = Contacts.contactID
  ) ON SchGrp.SGrpID = Units.GroupREF
WHERE
  (
    (
      (RecallCampaign.ID)= forms!recallpform!recallno
    )
    And (
      (Clients.ClientShNm)= forms!recallpform!client
    )
    And (
      (Locations.LocationID) Like IIf(
        Nz(forms!recallpform!clientcity),
        forms!recallpform!clientcity,
        "*"
      )
    )
  )
ORDER BY
  Locations.LocationName,
  MstrLease.Client,
  Left([unitvin], 6);
