CREATE TABLE [CustomFields_Values] (
  [RefID] LONG,
  [CustomID] LONG CONSTRAINT [CustomFieldsCustomFields_Values] REFERENCES [CustomFields] ([CustomID]) ON UPDATE CASCADE  ON DELETE CASCADE ,
  [FieldValue] DOUBLE,
  [FieldText] VARCHAR (255),
  [ID] AUTOINCREMENT CONSTRAINT [PrimaryKey] PRIMARY KEY UNIQUE NOT NULL
)
