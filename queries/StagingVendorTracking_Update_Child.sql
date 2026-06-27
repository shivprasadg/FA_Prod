SELECT
  StagingVendorTracking.ClientId,
  ChildParts.REFENG,
  ChildParts.COMPRESSS,
  ChildParts.MICROPROC,
  ChildParts.ELECDISP
FROM
  StagingVendorTracking
  INNER JOIN ChildParts ON StagingVendorTracking.ClientId = ChildParts.ChildPartID;
