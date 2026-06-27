CREATE TABLE [TaxBaseTemp] (
  [ID] AUTOINCREMENT CONSTRAINT [PK_TaxCountiesBase] PRIMARY KEY UNIQUE NOT NULL,
  [TaxTypeId] LONG,
  [BaseTaxRate] VARCHAR,
  [StartsOn] DATETIME,
  [EndsOn] DATETIME,
  [TaxLevelLookupID] LONG,
  [TaxLevel] VARCHAR (255),
  [EditedTaxID] LONG,
  [TaxScript] LONGTEXT,
  [TaxNote] LONGTEXT
)
