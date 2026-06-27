SELECT
  MstrLease.MLNo,
  MstrLease.MLOrig,
  Schedule.SchID,
  Schedule.SchNo,
  SchGrp.SGrpNO,
  SchGrp.SGrpID,
  Max(
    IIf(
      IsNull([unitpdstartdate]),
      IIf(
        IsNull([unitacceptdate]),
        IIf(
          IsNull([unitinserv]),
          IIf(
            IsNull([unitestdeldate]),
            [unitbuilddate],
            [unitestdeldate]
          ),
          [unitinserv]
        ),
        [unitacceptdate]
      ),
      [unitpdstartdate]
    )
  ) AS ProjIns,
  Avg(
    IIf(
      IsNull([unitpdstartdate]),
      IIf(
        IsNull([unitacceptdate]),
        IIf(
          IsNull([unitinserv]),
          IIf(
            IsNull([unitestdeldate]),
            [unitbuilddate],
            [unitestdeldate]
          ),
          [unitinserv]
        ),
        [unitacceptdate]
      ),
      [unitpdstartdate]
    )
  ) AS AVGIns
FROM
  (
    (
      MstrLease
      INNER JOIN Schedule ON MstrLease.MLID = Schedule.SchMLRef
    )
    INNER JOIN SchGrp ON Schedule.SchID = SchGrp.SGrpSchID
  )
  INNER JOIN Units ON SchGrp.SGrpID = Units.GroupREF
GROUP BY
  MstrLease.MLNo,
  MstrLease.MLOrig,
  Schedule.SchID,
  Schedule.SchNo,
  SchGrp.SGrpNO,
  SchGrp.SGrpID
HAVING
  (
    (
      (MstrLease.MLOrig)= "fa"
      Or (MstrLease.MLOrig)= "other"
    )
    AND (
      (
        Max(
          IIf(
            IsNull([unitpdstartdate]),
            IIf(
              IsNull([unitacceptdate]),
              IIf(
                IsNull([unitinserv]),
                IIf(
                  IsNull([unitestdeldate]),
                  [unitbuilddate],
                  [unitestdeldate]
                ),
                [unitinserv]
              ),
              [unitacceptdate]
            ),
            [unitpdstartdate]
          )
        )
      ) Is Not Null
    )
    AND (
      (
        Avg(
          IIf(
            IsNull([unitpdstartdate]),
            IIf(
              IsNull([unitacceptdate]),
              IIf(
                IsNull([unitinserv]),
                IIf(
                  IsNull([unitestdeldate]),
                  [unitbuilddate],
                  [unitestdeldate]
                ),
                [unitinserv]
              ),
              [unitacceptdate]
            ),
            [unitpdstartdate]
          )
        )
      ) Is Not Null
    )
  )
ORDER BY
  MstrLease.MLNo,
  Schedule.SchID;
