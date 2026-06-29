SELECT
  IIf(
    Len(
      Trim([VIN])
    )< 17,
    "Short",
    IIf(
      Len(
        Trim([VIN])
      )> 17,
      "Long",
      "Good"
    )
  ) AS VINLenght,
  FS.UnitNumber,
  FS.OBCUnitNumber,
  Trim(FS.[VIN]) AS VIN,
  FS.UnitYear,
  FS.UnitMake,
  FS.UnitModel,
  FS.Type,
  FS.SubType,
  FS.InServiceDate,
  FS.Axle,
  FS.UnitRent,
  FS.LeaseStartDate,
  FS.LeaseEndDate,
  FS.Unitcost,
  FS.LeaseTerm,
  Nz([FS].[Odometer], 0) AS Odometer,
  FS.OdometerDate,
  FS.Ownership,
  FS.Location,
  FS.State,
  FS.Plate,
  FS.ReeferHours,
  FS.Division,
  FS.DivisionID,
  FS.Engine,
  FS.Trans
FROM
  FleetStudyImport_Work AS FS
  LEFT JOIN [SELECT Units].[UnitID, Units].[UnitVIN, Units].[ClientID, Units].[ClientGroupID FROM Units WHERE Units].[ClientGroupID=[Forms]]![FleetStudy_Importer]![ClientGroupId] AS U ON FS.VIN = U.UnitVIN
WHERE
  (
    (
      (
        Trim(FS.VIN)
      )<> ""
    )
    And (
      (U.UnitVIN) Is Null
    )
  );
