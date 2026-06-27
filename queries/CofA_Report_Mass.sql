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
  vw_CofABase.*
FROM
  vw_CofABase
  INNER JOIN Bulk_Update_WorkTable ON vw_CofABase.UnitID = Bulk_Update_WorkTable.UnitID
WHERE
  (
    (
      (vw_CofABase.SGrpID)= [Forms]![BulkChange_Form]![GrpID]
    )
  );
