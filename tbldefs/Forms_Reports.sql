CREATE TABLE [Forms_Reports] (
  [index] AUTOINCREMENT CONSTRAINT [PrimaryKey] PRIMARY KEY UNIQUE NOT NULL,
  [Type] VARCHAR (30),
  [ItemName] VARCHAR (100),
  [ItemMenuName] VARCHAR (100),
  [ItemCommand] LONGTEXT,
  [Group] VARCHAR (25),
  [SubGroup] VARCHAR (255),
  [SecurityLevel] VARCHAR (25),
  [ParentMenu] VARCHAR (30),
  [MenuLocation] VARCHAR (255),
  [Sequence] SHORT,
  [Active] BIT
)
