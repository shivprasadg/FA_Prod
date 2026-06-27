INSERT INTO TrackingSheetChangedDatesLog (
  UnitID, FAID, SchNo, GrpNo, UnitNum,
  PONum, SubmissionDate, BuildDate,
  EstimatedDeliveryDate, DeliveryDate,
  PrevBuildDate, PrevEstimatedDate,
  PrevDeliveryDate
)
SELECT
  Units.UnitID,
  vw_SixKeys.FAID,
  vw_SixKeys.Schedule AS SchNo,
  vw_SixKeys.UnitGroup AS GrpNo,
  Units.UnitUnitNum AS UnitNum,
  vw_SixKeys.PurchaseOrder AS PONum,
  Date() AS SubmissionDate,
  OpenVendorTrackingUpdate_BuildDate.BuildDate,
  OpenVendorTrackingUpdate_EstimatedDeliveryDate.EstimatedDeliveryDate,
  OpenVendorTrackingUpdate_DeliveryDate.DeliveryDate,
  IIf(
    [PrevBuildDate] = "1/1/1900", Null,
    [PrevBuildDate]
  ) AS PrevBuild,
  OpenVendorTrackingUpdate_EstimatedDeliveryDate.PrevEstimatedDate AS PrevEstimatedDate,
  OpenVendorTrackingUpdate_DeliveryDate.PrevDeliveryDate
FROM
  (
    (
      (
        (
          OpenVendorDataTemp
          INNER JOIN Units ON OpenVendorDataTemp.AssetID = Units.UnitID
        )
        INNER JOIN vw_SixKeys ON OpenVendorDataTemp.AssetID = vw_SixKeys.AssetID
      )
      LEFT JOIN OpenVendorTrackingUpdate_BuildDate ON Units.UnitID = OpenVendorTrackingUpdate_BuildDate.UnitID
    )
    LEFT JOIN OpenVendorTrackingUpdate_EstimatedDeliveryDate ON Units.UnitID = OpenVendorTrackingUpdate_EstimatedDeliveryDate.UnitID
  )
  LEFT JOIN OpenVendorTrackingUpdate_DeliveryDate ON Units.UnitID = OpenVendorTrackingUpdate_DeliveryDate.UnitID
WHERE
  (
    (
      (
        [BuildDate] & [EstimatedDeliveryDate] & [BuildDate]
      ) Is Not Null
    )
  );
