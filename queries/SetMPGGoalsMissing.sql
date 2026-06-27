UPDATE
  Units
SET
  Units.MpgGoal = 6.5,
  Units.DrivingMpgGoal = 6.5
WHERE
  (
    (
      (Units.MpgGoal) Is Null
      Or (Units.MpgGoal)= 0
    )
    AND (
      (Units.DrivingMpgGoal) Is Null
      Or (Units.DrivingMpgGoal)= 0
    )
  );
