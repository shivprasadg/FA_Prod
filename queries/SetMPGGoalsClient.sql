UPDATE
  Units
SET
  Units.MpgGoal = [MPGOverAll],
  Units.DrivingMpgGoal = [MPGDriving]
WHERE
  (
    (
      (Units.ClientId)= [Which ClientID]
    )
  );
