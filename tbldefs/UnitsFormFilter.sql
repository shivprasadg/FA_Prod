CREATE TABLE [UnitsFormFilter] (
  [AssetIDx] LONG,
  [ClientID] VARCHAR (255),
  [BLCD] VARCHAR (255),
  [ClientShNm] VARCHAR (255),
  [CurrRent] CURRENCY,
  [CurrRentType] VARCHAR (255),
  [MLNo] VARCHAR (255),
  [MLOrig] VARCHAR (255),
  [LeaseExp] DATETIME,
  [ProjOffLease] DATETIME,
  [SchID] LONG,
  [SchNo] VARCHAR (255),
  [sGrpNo] VARCHAR (255),
  [UnitOEC] CURRENCY,
  [FundedDate] DATETIME,
  [index] AUTOINCREMENT CONSTRAINT [PK_UnitWorkTable] PRIMARY KEY UNIQUE NOT NULL
)
