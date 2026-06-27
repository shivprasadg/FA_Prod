SELECT
  SchGrp.SGrpSchID AS SchId,
  SchGrp.SGrpID AS GroupId,
  SchGrp.BLCD AS BLCD,
  SchGrp.ScheduleGroupPaymentTerm AS FeeTerm,
  SchGrp.SGrpFSLPmt AS FSLFeeX,
  DateAdd(
    "m",
    Nz([ScheduleGroupPaymentTerm], 0),
    [BLCD] - 1
  ) AS DateFeesEnd
FROM
  SchGrp
GROUP BY
  SchGrp.SGrpSchID,
  SchGrp.SGrpID,
  SchGrp.BLCD,
  SchGrp.ScheduleGroupPaymentTerm,
  SchGrp.SGrpFSLPmt,
  DateAdd(
    "m",
    Nz([ScheduleGroupPaymentTerm], 0),
    [BLCD] - 1
  );
