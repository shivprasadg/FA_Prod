CREATE TABLE [CloneSchGrp_WorkTable] (
  [ID] AUTOINCREMENT CONSTRAINT [PrimaryKey] PRIMARY KEY UNIQUE NOT NULL,
  [SourceSchID] LONG,
  [SourceGroupID] LONG,
  [SourceSchNo] VARCHAR (255),
  [SourceGroupNo] VARCHAR (255),
  [CloneMe] BIT,
  [NewSchID] LONG,
  [NewSchNo] VARCHAR (255),
  [NewSchDesc] VARCHAR (255),
  [NewGroupNo] VARCHAR (255),
  [NewGroupDesc] VARCHAR (255)
)
