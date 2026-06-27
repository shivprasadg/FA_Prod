CREATE TABLE [_CACHE_FindAndReplace] (
  [Container] VARCHAR (8),
  [DocumentName] VARCHAR (64),
  [EntryNumber] AUTOINCREMENT,
  [LastUpdated] DATETIME,
  [CachedText] LONGTEXT,
  [CustomProps] LONGTEXT,
  [Vers9a] SHORT,
   CONSTRAINT [PrimaryKey] PRIMARY KEY ([Container], [DocumentName], [EntryNumber])
)
