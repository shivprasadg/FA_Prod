CREATE TABLE [ScheduleHealthCheckWorkTable] (
  [ID] AUTOINCREMENT CONSTRAINT [PrimaryKey] PRIMARY KEY UNIQUE NOT NULL,
  [ClientID] LONG,
  [SchID] LONG,
  [GroupID] LONG,
  [SchNo] VARCHAR (255),
  [Client] VARCHAR (255),
  [SyndRent] VARCHAR (255),
  [PDRent] VARCHAR (255),
  [FundDate] VARCHAR (255),
  [UnitCost] VARCHAR (255),
  [OrigEquipCost] VARCHAR (255),
  [SchOut] VARCHAR (255),
  [SchIn] VARCHAR (255)
)
