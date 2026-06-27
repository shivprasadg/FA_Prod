SELECT
  RALSource.clientID,
  RALSource.SchID,
  RALSource.RemitAddressID,
  Date() AS Letterdate,
  "Gentlemen:" AS Salutation,
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
