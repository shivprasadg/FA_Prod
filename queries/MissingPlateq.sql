SELECT
  MstrLease.MLNo,
  Schedule.SchNo,
  IIf(
    [SchTitleToTrust] = True, "Yes", "No"
  ) AS [TITLETrust?],
  IIf([SchTitltoFA] = True, "Yes", "No") AS [TiteltoFA?],
  SchGrp.SGrpNO,
  Units.unitunitnum,
  Units.unitvin,
  Units.Unityr,
  Units.unitmake,
  Units.unittype,
  Locations.LState,
  Units.UnitAcceptDate,
  Units.UnitDocAccpt,
  Units.UnitInServ,
  Units.unitinservind,
  Clients.ClientShNm,
  Clients.ClientGroupID
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
    LEFT JOIN Locations ON Units.UnitTitleLocationRef = Locations.LocationID
  ) ON SchGrp.SGrpID = Units.GroupREF
WHERE
  (
    (
      (
        IIf(
          [SchTitleToTrust] = True, "Yes", "No"
        )
      )= "yes"
    )
    AND (
      (
        IIf([SchTitltoFA] = True, "Yes", "No")
      )= "no"
    )
    AND (
      (Units.UnitAcceptDate) Is Not Null
    )
    AND (
      (Units.UnitDocAccpt)= Yes
    )
    AND (
      (Units.UnitInServ) Is Not Null
      And (Units.UnitInServ)> #12/31/2010#
    )
    AND (
      (Units.unitinservind)= "a"
    )
    AND (
      (Clients.ClientGroupID)<> 18
    )
    AND (
      (Units.UnitPlateNum) Is Null
    )
    AND (
      (MstrLease.MLOrig)= "fa"
    )
    AND (
      (Units.unitstatus)= "a"
      Or (Units.unitstatus)= "tr"
      Or (Units.unitstatus)= "c"
    )
  )
  OR (
    (
      (
        IIf(
          [SchTitleToTrust] = True, "Yes", "No"
        )
      )= "no"
    )
    AND (
      (
        IIf([SchTitltoFA] = True, "Yes", "No")
      )= "yes"
    )
    AND (
      (Units.UnitAcceptDate) Is Not Null
    )
    AND (
      (Units.UnitDocAccpt)= Yes
    )
    AND (
      (Units.UnitInServ) Is Not Null
      And (Units.UnitInServ)> #12/31/2010#
    )
    AND (
      (Units.unitinservind)= "a"
    )
    AND (
      (Clients.ClientGroupID)<> 18
    )
    AND (
      (Units.UnitPlateNum) Is Null
    )
    AND (
      (MstrLease.MLOrig)= "fa"
    )
    AND (
      (Units.unitstatus)= "a"
      Or (Units.unitstatus)= "tr"
      Or (Units.unitstatus)= "c"
    )
  )
  OR (
    (
      (
        IIf(
          [SchTitleToTrust] = True, "Yes", "No"
        )
      )= "yes"
    )
    AND (
      (
        IIf([SchTitltoFA] = True, "Yes", "No")
      )= "no"
    )
    AND (
      (Units.UnitAcceptDate) Is Not Null
      And (Units.UnitAcceptDate)> #12/31/2010#
    )
    AND (
      (Units.UnitDocAccpt)= No
    )
    AND (
      (Clients.ClientGroupID)<> 18
    )
    AND (
      (Units.UnitPlateNum) Is Null
    )
    AND (
      (MstrLease.MLOrig)= "fa"
    )
    AND (
      (Units.unitstatus)= "a"
      Or (Units.unitstatus)= "tr"
      Or (Units.unitstatus)= "c"
    )
  )
  OR (
    (
      (
        IIf(
          [SchTitleToTrust] = True, "Yes", "No"
        )
      )= "no"
    )
    AND (
      (
        IIf([SchTitltoFA] = True, "Yes", "No")
      )= "yes"
    )
    AND (
      (Units.UnitAcceptDate) Is Not Null
      And (Units.UnitAcceptDate)> #12/31/2010#
    )
    AND (
      (Units.UnitDocAccpt)= No
    )
    AND (
      (Clients.ClientGroupID)<> 18
    )
    AND (
      (Units.UnitPlateNum) Is Null
    )
    AND (
      (MstrLease.MLOrig)= "fa"
    )
    AND (
      (Units.unitstatus)= "a"
      Or (Units.unitstatus)= "tr"
      Or (Units.unitstatus)= "c"
    )
  )
ORDER BY
  MstrLease.MLNo,
  Schedule.SchID,
  SchGrp.SGrpNO;
