SELECT
  Units.UNITMAKE,
  Units.UNITVIN,
  Units.UNITMODEL,
  Units.UNITUNITNUM,
  Units.UNITTYPE,
  Rmkt.[SUF Sent],
  Units.GroupREF
FROM
  Units
  LEFT JOIN Rmkt ON Units.UnitID = Rmkt.UnitRef
WHERE
  (
    (
      (Units.GroupREF)= [forms]![groupform]![sgrpid]
    )
  );
