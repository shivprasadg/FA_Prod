CREATE TABLE [_LOG_FindAndReplace] (
  [SeqID] AUTOINCREMENT CONSTRAINT [PrimaryKey] PRIMARY KEY UNIQUE NOT NULL,
  [SessionTimeStamp] DATETIME CONSTRAINT [SessionIndex] NOT NULL,
  [EntryType] VARCHAR (8),
  [Container] VARCHAR (24),
  [DocumentName] VARCHAR (150),
  [ElementName] VARCHAR (150),
  [PropertyName] VARCHAR (64),
  [PropertyType] LONG,
  [OtherInfo] VARCHAR (64),
  [ReplacePerformed] VARCHAR (1),
  [OldValue] LONGTEXT,
  [NewValue] LONGTEXT,
  [lngValue1] LONG,
  [VersB4] SHORT
)
