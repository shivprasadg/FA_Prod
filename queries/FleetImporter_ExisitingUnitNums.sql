SELECT
  Units.UnitID,
  Units.UnitUnitNum,
  Units.ClientGroupId,
  Units.ClientID
FROM
  Units
WHERE
  (
    (
      (Units.ClientID)= [Forms]![FleetStudy_Importer]![ClientID]
    )
  );
