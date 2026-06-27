SELECT
  Schedule.SchID,
  Schedule.SchNo,
  SchGrp.SGrpNO AS GrpNo,
  ProjBLCDq.ProjBLCD,
  [schedule].[schno] & "_" & [Sgrpno] AS Sch,
  Schedule.SchShortDesc AS ShortDesc,
  SchGrp.LeaseTermPrimary AS PTErm,
  SchGrp.UnitsInGroup AS NoofUnits,
  SchGrp.sgrpfslpmt AS FSLPmt,
  FSL.ID AS FSLSch,
  FSL.FSLYR1,
  FSL.FSLYR2,
  FSL.FSLYR3,
  FSL.FSLYR4,
  FSL.FSLYR5,
  FSL.FSLYR6,
  FSL.FSLYR7,
  Clients.ClientShNm
FROM
  (
    (
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
        INNER JOIN FSL ON Schedule.FSLSCH = FSL.ID
      )
      INNER JOIN Units ON SchGrp.SGrpID = Units.GroupREF
    )
    INNER JOIN Locations ON Units.UnitGarageLocationRef = Locations.LocationID
  )
  LEFT JOIN ProjBLCDq ON Schedule.SchID = ProjBLCDq.SchID
GROUP BY
  Schedule.SchID,
  Schedule.SchNo,
  SchGrp.SGrpNO,
  ProjBLCDq.ProjBLCD,
  [schedule].[schno] & "_" & [Sgrpno],
  Schedule.SchShortDesc,
  SchGrp.LeaseTermPrimary,
  SchGrp.UnitsInGroup,
  SchGrp.sgrpfslpmt,
  FSL.ID,
  FSL.FSLYR1,
  FSL.FSLYR2,
  FSL.FSLYR3,
  FSL.FSLYR4,
  FSL.FSLYR5,
  FSL.FSLYR6,
  FSL.FSLYR7,
  Clients.ClientShNm,
  SchGrp.FundDate
HAVING
  (
    (
      (Schedule.SchID)> 24
    )
    AND (
      (Clients.ClientShNm)= "costco"
    )
  )
ORDER BY
  Schedule.SchID,
  Schedule.SchNo,
  SchGrp.SGrpNO;
