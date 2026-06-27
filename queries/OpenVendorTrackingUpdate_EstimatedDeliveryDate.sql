SELECT
  Units.UnitID,
  OpenVendorDataTemp.UnitEstDelDate AS EstimatedDeliveryDate,
  Units.UnitEstDelDate AS PrevEstimatedDate
FROM
  OpenVendorDataTemp
  INNER JOIN Units ON OpenVendorDataTemp.AssetID = Units.UnitID
WHERE
  (
    (
      (Units.UnitEstDelDate)<> [OpenVendorDataTemp].[UnitEstDelDate]
    )
  );
