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
        UnitReportStartEndDates.ReportEndDate
      ) Is Null
      Or (
        UnitReportStartEndDates.ReportEndDate
      )<> [UnitReportStartEndDates].[EndDate]
    )
  );
