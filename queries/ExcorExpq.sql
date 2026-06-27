PARAMETERS [enter start date] DateTime,
[enter end date] DateTime;
SELECT
  Clients.ClientShNm,
  MstrLease.MLNo,
  Schedule.SchNo,
  SchGrp.Assignee_Bank,
  SchGrp.BLCD,
  SchGrp.ExchangeITGroup,
  SchGrp.ExchangeITMonth,
  SchGrp.LeaseTermPrimary,
  IIf(
    IsNull([ExchangeITMonth])
    Or [ExchangeITMonth] = 0,
    "N/A",
    DateAdd("m", [ExchangeITMonth], [BLCD])
  ) AS Exchangeit,
  DateAdd("m", [ExchangeITMonth], [BLCD]) AS EXDT,
  DateAdd("m", [LeaseTermPrimary], [BLCD])-1 AS Exp,
  SchGrp.UnitsInGroup,
  Units.UnitUnitNum,
  Units.UnitVIN,
  Right([unitvin], 6) AS vin4,
  UCase(
    [unityr] & " " & [unitmake] & " " & [unitmodel] & " " & [unittype]
  ) AS [Desc],
  Units.unitstatus,
  Locations.LocationName,
  Locations.LCity,
  Locations.LState
FROM
  (
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
    INNER JOIN Units ON SchGrp.SGrpID = Units.GroupREF
  )
  INNER JOIN Locations ON Units.UnitGarageLocationRef = Locations.LocationID
WHERE
  (
    (
      (SchGrp.BLCD) Is Not Null
    )
    AND (
      (
        DateAdd("m", [ExchangeITMonth], [BLCD])
      )>= [enter start date]
      And (
        DateAdd("m", [ExchangeITMonth], [BLCD])
      )<= [enter end date]
    )
    AND (
      (Units.unitstatus)= "a"
      Or (Units.unitstatus)= "c"
    )
  )
  OR (
    (
      (SchGrp.BLCD) Is Not Null
    )
    AND (
      (
        DateAdd("m", [ExchangeITMonth], [BLCD])
      )>= [enter start date]
      And (
        DateAdd("m", [ExchangeITMonth], [BLCD])
      )<= [enter end date]
    )
    AND (
      (Units.unitstatus)= "a"
      Or (Units.unitstatus)= "c"
    )
  )
  OR (
    (
      (SchGrp.BLCD) Is Not Null
    )
    AND (
      (
        DateAdd("m", [LeaseTermPrimary], [BLCD])-1
      )>= [enter start date]
      And (
        DateAdd("m", [LeaseTermPrimary], [BLCD])-1
      )<= [enter end date]
    )
    AND (
      (Units.unitstatus)= "a"
      Or (Units.unitstatus)= "c"
    )
  )
  OR (
    (
      (SchGrp.BLCD) Is Not Null
    )
    AND (
      (
        DateAdd("m", [LeaseTermPrimary], [BLCD])-1
      )>= [enter start date]
      And (
        DateAdd("m", [LeaseTermPrimary], [BLCD])-1
      )<= [enter end date]
    )
  )
ORDER BY
  Units.UnitUnitNum,
  Right([unitvin], 6);
