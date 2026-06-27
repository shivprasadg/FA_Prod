CREATE TABLE [_LOG_FindAndReplace_xref] (
  [FoundID] LONG,
  [FindSpecID] LONG CONSTRAINT [FindSpecIndex] NOT NULL,
   CONSTRAINT [PrimaryKey] PRIMARY KEY ([FoundID], [FindSpecID])
)
