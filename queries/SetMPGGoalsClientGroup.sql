UPDATE
  Units
SET
  Units.MpgGoal = [MPGOverAll],
  Units.DrivingMpgGoal = [MPGDriving]
WHERE
  (
    (
      (Units.ClientGroupId)= [Which ClientGroupID]
    )
  );
