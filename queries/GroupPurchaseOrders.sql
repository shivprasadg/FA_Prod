SELECT
  SchGrp.POVendor,
  SchGrp.POSentDt As POSentDate,
  SchGrp.PONum1 As PONumber,
  'UnitOEMPO' As POType,
  sGrpId As GroupId
FROM
  SchGrp
Union ALL
SELECT
  PO2Vendor,
  PO2SentDt,
  PONum2,
  'UpFitPO1' As POType,
  sGrpId As GroupId
FROM
  SchGrp
Union ALL
SELECT
  PO3Vendor,
  PO3SentDt,
  PONum3,
  'UpFitPO2' As POType,
  sGrpId As GroupId
FROM
  SchGrp
UNION ALL
SELECT
  PO4Vendor,
  PO4SentDt,
  PONum4,
  'UpFitPO3' As POType,
  sGrpId As GroupId
FROM
  SchGrp;
