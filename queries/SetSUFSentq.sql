UPDATE
  Units
  LEFT JOIN Rmkt ON Units.UnitID = Rmkt.UnitRef
SET
  Rmkt.UnitRef = [unitid]
WHERE
  (
    (
      (Units.GroupREF)= forms!groupform!sgrpid
    )
    And (
      (Units.UnitStatus)= "tr"
    )
  );
