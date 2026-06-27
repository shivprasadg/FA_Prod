CREATE TABLE [CustomFields] (
  [RefID] LONG,
  [RefTable] VARCHAR (50),
  [RefFieldName] VARCHAR (50),
  [FieldName] VARCHAR (55),
  [DisplayName] VARCHAR (50),
  [FieldType] VARCHAR (20),
  [Format] VARCHAR (20),
  [CustomID] AUTOINCREMENT CONSTRAINT [PrimaryKey] PRIMARY KEY UNIQUE NOT NULL,
  [Value] DOUBLE,
  [Text] VARCHAR (75),
  [HL] LONGTEXT
)
