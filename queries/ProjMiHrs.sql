SELECT
  Units.GroupREF,
  Units.UnitID,
  Units.Unitinserv,
  MaxMiQ.MaxOfHubMI,
  MaxMiQ.MaxOfMiClDt,
  IIf(
    IsNull([MaxOfReefer Hours]),
    0,
    [maxofreefer hours]
  ) AS MaxHrs,
  MaxHrsQ.[MaxOfClose DAte] AS MaxHrsDt,
  DateAdd("m", [LeaseTermPrimary], [BLCD])-1 AS Expiry,
  DateDiff("d", [unitinserv], [MaxofMIclDt])+ 1 AS MiDays,
  IIf(
    IsNull([maxofclose date]),
    0,
    DateDiff(
      "d", [unitinserv], [maxofclose date]
    )+ 1
  ) AS HrDays,
  [MaxOfHubMI] /(
    DateDiff("d", [unitinserv], [MaxofMIclDt])+ 1
  ) AS Dailymi,
  IIf(
    IsNull([maxofreefer hours]),
    0,
    [maxofreefer hours] /(
      DateDiff(
        "d", [unitinserv], [maxofclose date]
      )+ 1
    )
  ) AS DailyHr,
  DateDiff(
    "d",
    [unitinserv],
    (
      DateAdd("m", [LeaseTermPrimary], [BLCD])-1
    )
  ) AS Leasedays,
  [sgprallow1] * [LeaseTermPrimary] / 12 AS MiAllowance,
  (
    DateDiff(
      "d",
      [unitinserv],
      (
        DateAdd("m", [LeaseTermPrimary], [BLCD])-1
      )
    )
  )*(
    [MaxOfHubMI] /(
      DateDiff("d", [unitinserv], [MaxofMIclDt])+ 1
    )
  ) AS ProjLTMI,
  [sgprallow4] * [LeaseTermPrimary] / 12 AS HrAllowance,
  (
    DateDiff(
      "d",
      [unitinserv],
      (
        DateAdd("m", [LeaseTermPrimary], [BLCD])-1
      )
    )
  )*(
    IIf(
      IsNull([maxofreefer hours]),
      0,
      [maxofreefer hours] /(
        DateDiff(
          "d", [unitinserv], [maxofclose date]
        )+ 1
      )
    )
  ) AS ProjLTDHr,
  Units.UnitUnitNum,
  (
    (
      DateDiff(
        "d",
        [unitinserv],
        (
          DateAdd("m", [LeaseTermPrimary], [BLCD])-1
        )
      )
    )*(
      [MaxOfHubMI] /(
        DateDiff("d", [unitinserv], [MaxofMIclDt])+ 1
      )
    )
  )/ [LeaseTermPrimary] AS ProjMoMi,
  (
    (
      DateDiff(
        "d",
        [unitinserv],
        (
          DateAdd("m", [LeaseTermPrimary], [BLCD])-1
        )
      )
    )*(
      IIf(
        IsNull([maxofreefer hours]),
        0,
        [maxofreefer hours] /(
          DateDiff(
            "d", [unitinserv], [maxofclose date]
          )+ 1
        )
      )
    )
  )/ [LeaseTermPrimary] AS ProjMoHr,
  [SGPRALLOW1] / 12 AS LeaseMoMI,
  [sgPRallow4] / 12 AS LeaseMoHr
FROM
  (
    (
      MstrLease
      INNER JOIN Schedule ON MstrLease.MLID = Schedule.SchMLRef
    )
    INNER JOIN SchGrp ON Schedule.SchID = SchGrp.SGrpSchID
  )
  INNER JOIN (
    (
      Units
      LEFT JOIN MaxMiQ ON Units.UnitID = MaxMiQ.Unitref
    )
    LEFT JOIN MaxHrsQ ON Units.UnitID = MaxHrsQ.Unitref
  ) ON SchGrp.SGrpID = Units.GroupREF
WHERE
  (
    (
      (Schedule.SchNo) Like "*toll*"
    )
    AND (
      (MstrLease.MLNo) Like "Costco*"
    )
  )
ORDER BY
  Schedule.SchNo,
  Units.GroupREF,
  Units.UnitID;
