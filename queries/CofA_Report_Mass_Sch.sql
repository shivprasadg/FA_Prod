SELECT
  Switch(
    [SGRPPLB] = False,
    Null,
    Nz([PLBMileage], 0)> 0,
    Format([PLBMileage], "#,##0"),
    Nz([PLBHours], 0),
    Format([PLBHours], "#,##0"),
    Nz([PLBMileage], 0)> 0
    And Nz([PLBHours], 0)> 0,
    Format([PLBMileage], "#,##0")& " / " & Format([PLBHours], "#,##0")
  ) AS PLBMilesHours,
  Switch(
    [SGRPPLB] = False,
    Null,
    Nz([PLBMileage], 0)> 0,
    "Miles as of " & Format(
      [Forms]![BulkChange_Form]![PLBAsOfDate],
      "mm/dd/yy"
    ),
    Nz([PLBHours], 0)> 0,
    "Hours",
    Nz([PLBMileage], 0)= 0
    And Nz([PLBHours], 0)= 0,
    "Miles / Hours"
  ) AS PLBLabel,
  *
FROM
  vw_CofABase
WHERE
  (
    (
      (vw_CofABase.SchID)= [Forms]![BulkChange_Form]![ScheduleList]
    )
  );
