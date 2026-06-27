INSERT INTO RALTemp1 (
  clientID, SchID, RemitAddressID, Letterdate,
  Salutation, ScheduleDate, Schedule,
  MLNo, MLDate, Lessee, LesseeAddress,
  BankRemitAddress, Paragraph1, Paragraph2,
  Paragraph3, Paragraph4, Paragraph5
)
SELECT
  RALSource.clientID,
  RALSource.SchID,
  RALSource.RemitAddressID,
  Format(
    Date(),
    "mmmm dd"", ""yyyy"
  )& "<br />" AS Letterdate,
  "Gentlemen:<br />" AS Salutation,
  RALSource.ScheduleDate,
  RALSource.Schedule,
  RALSource.MLNo,
  RALSource.MLDate,
  RALSource.Lessee,
  RALSource.LesseeAddress,
  RALSource.BankRemitAddress,
  RALSource.Paragraph1,
  [Paragraph2a] & [Paragraph2b] AS Paragraph2,
  [Paragraph3a] & [Paragraph3b] & [Paragraph3c] AS Paragraph3,
  RALSource.Paragraph4,
  RALSource.Paragraph5
FROM
  RALSource;
