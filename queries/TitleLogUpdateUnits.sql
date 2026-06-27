UPDATE
  TitleLogTemp
  INNER JOIN Units ON TitleLogTemp.UnitID = Units.UnitID
SET
  Units.UnitTitleNumber = [TitleLogTemp].[Titlenumber],
  Units.UnitTitleRecd = [TitleLogTemp].[DateLog];
