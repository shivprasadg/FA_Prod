SELECT
  Clients.ClientShNm AS Client,
  Schedule.SchFAID AS FAID,
  Schedule.SchNo AS Schedule,
  SchGrp.SGrpNO AS [Group],
  SchGrp.PONum1 AS POnumber,
  SchGrp.POSentDt AS POSent,
  Vendors.VendorName,
  CDate(
    DMin(
      "UnitEstDelDate", "Units", "GroupREF=" & [SchGrp].[SGrpID] & ""
    )
  ) AS EarlyEstimated,
  CDate(
    DMin(
      "UnitBuildDate", "Units", "GroupREF=" & [SchGrp].[SGrpID] & ""
    )
  ) AS EarlyBuild,
  CDate(
    DMax(
      "UnitBuildDate", "Units", "GroupREF=" & [SchGrp].[SGrpID] & ""
    )
  ) AS LateBuild,
  CDate(
    DMin(
      "UnitActualDelDate", "Units", "GroupREF=" & [SchGrp].[SGrpID] & ""
    )
  ) AS EarlyActualDelivery,
  CDate(
    DMax(
      "UnitActualDelDate", "Units", "GroupREF=" & [SchGrp].[SGrpID] & ""
    )
  ) AS LateActualDelivery
FROM
  (
    (
      MstrLease
      INNER JOIN (
        Schedule
        INNER JOIN SchGrp ON Schedule.SchID = SchGrp.SGrpSchID
      ) ON MstrLease.MLID = Schedule.SchMLRef
    )
    LEFT JOIN Vendors ON SchGrp.POVendor = Vendors.VendorID
  )
  INNER JOIN Clients ON MstrLease.Client = Clients.clientID
WHERE
  (
    (
      (Clients.ClientShNm) Not Like "%Dem%"
    )
    AND (
      (Schedule.SchFAID) Like "14%"
      Or (Schedule.SchFAID) Like "13%"
      Or (Schedule.SchFAID) Like "12%"
      Or (Schedule.SchFAID) Like "11%"
      Or (Schedule.SchFAID) Like "10%"
    )
    AND (
      (MstrLease.MLOrig)= "FA"
    )
    AND (
      (Schedule.SchID) Is Not Null
    )
    AND (
      (SchGrp.SGrpID) Is Not Null
    )
  )
ORDER BY
  Schedule.SchFAID,
  Schedule.SchNo,
  SchGrp.SGrpNO;
