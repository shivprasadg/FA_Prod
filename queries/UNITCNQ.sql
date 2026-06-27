SELECT
  IIf(
    IsNull([unitacceptdate]),
    DatePart(
      "yyyy",
      Date()
    ),
    DatePart("yyyy", [UnitAcceptdate])
  ) AS yr,
  Units.Unitsubtype AS UType,
  Count(Units.UnitID) AS CountOfUnitID,
  Units.UnitType AS grpClass
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
    Units
    LEFT JOIN lookup_groupTypes ON Units.UnitType = lookup_groupTypes.groupType
  ) ON SchGrp.SGrpID = Units.GroupREF
WHERE
  (
    (
      (SchGrp.LeaseType)<> "loan"
    )
    AND (
      (Clients.ClientShNm) Not Like "demo*"
    )
    AND (
      (Units.unitstatus)= "a"
      Or (Units.unitstatus)= "c"
      Or (Units.unitstatus)= "b"
    )
  )
GROUP BY
  IIf(
    IsNull([unitacceptdate]),
    DatePart(
      "yyyy",
      Date()
    ),
    DatePart("yyyy", [UnitAcceptdate])
  ),
  Units.Unitsubtype,
  Units.UnitType
HAVING
  (
    (
      (Units.UnitType) Is Not Null
      And (Units.UnitType)<> "misc"
    )
  )
ORDER BY
  IIf(
    IsNull([unitacceptdate]),
    DatePart(
      "yyyy",
      Date()
    ),
    DatePart("yyyy", [UnitAcceptdate])
  ),
  Units.Unitsubtype,
  Count(Units.UnitID);
