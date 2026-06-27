CREATE TABLE [OpenVendorDataTemp] (
  [AssetID] LONG CONSTRAINT [PK_tblCustomers] PRIMARY KEY UNIQUE NOT NULL,
  [UnitVIN] VARCHAR (50),
  [UnitBuildDate] DATETIME,
  [UnitEstDelDate] DATETIME,
  [UnitActualDelDate] DATETIME,
  [UnitEngineSerial] VARCHAR (50),
  [UnitTransSerial] VARCHAR (50),
  [UnitFASerial] VARCHAR (50),
  [UnitRAFSerial] VARCHAR (50),
  [UnitRARSerial] VARCHAR (50),
  [Ref_Serial] VARCHAR (50),
  [Ref_Engine] VARCHAR (50)
)
