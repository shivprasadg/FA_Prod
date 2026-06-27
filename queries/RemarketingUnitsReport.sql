SELECT
  Rmkt.InvoicePaidDate,
  vw_SixKeys.AssetID,
  vw_SixKeys.Status,
  vw_SixKeys.clientID,
  vw_SixKeys.MLID,
  vw_SixKeys.SchID,
  vw_SixKeys.SGrpID
FROM
  Rmkt
  INNER JOIN vw_SixKeys ON Rmkt.UnitRef = vw_SixKeys.AssetID
WHERE
  (
    (
      (Rmkt.InvoicePaidDate) Is Null
      Or (Rmkt.InvoicePaidDate) Between [forms]![rmkt Menu]![FromDate]
      And [forms]![rmkt Menu]![ToDate]
    )
  );
