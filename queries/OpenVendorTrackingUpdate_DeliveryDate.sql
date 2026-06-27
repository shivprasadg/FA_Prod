SELECT
  Units.UnitID,
  OpenVendorDataTemp.UnitActualDelDate AS DeliveryDate,
  Units.UnitActualDelDate AS PrevDeliveryDate
FROM
  OpenVendorDataTemp
  INNER JOIN Units ON OpenVendorDataTemp.AssetID = Units.UnitID
WHERE
  (
    (
      (Units.UnitActualDelDate)<> [OpenVendorDataTemp].[UnitActualDelDate]
    )
  );
