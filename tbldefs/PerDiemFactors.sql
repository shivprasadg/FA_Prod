CREATE TABLE [PerDiemFactors] (
  [ID] AUTOINCREMENT CONSTRAINT [PrimaryKey] PRIMARY KEY UNIQUE NOT NULL,
  [ClientID] LONG,
  [ScheduleID] LONG,
  [GroupID] LONG,
  [PerdiemFactor] DOUBLE
)
