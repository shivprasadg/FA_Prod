SELECT
  UnitReportStartEndDates.UnitID,
  UnitReportStartEndDates.ReportStartDate,
  UnitReportStartEndDates.ReportEndDate,
  UnitReportStartEndDates.StartDate,
  UnitReportStartEndDates.EndDate,
  UnitReportStartEndDates.UnitStatus
FROM
  UnitReportStartEndDates
WHERE
  (
    (
      (
        UnitReportStartEndDates.ReportStartDate
      ) Is Null
      Or (
        UnitReportStartEndDates.ReportStartDate
      )<> [UnitReportStartEndDates].[StartDate]
    )
  );
