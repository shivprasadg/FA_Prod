SELECT
  DISTINCT vw_ATF_Avail_Signed_FAID.clientID,
  vw_ATF_Avail_Signed_FAID.ClientGroupId,
  ATFScheduleSpec.ATFSchSpecID,
  ATFGroupSpec.ATFGroupSpecID,
  ATF.FAID,
  ATF.ATFID,
  vw_ATF_Avail_Signed_FAID.[Spec_Name] AS ATFUnitSpecDesc,
  vw_ATF_Avail_Signed_FAID.Spec_UnitQty AS NumUnitsInSpec,
  CInt([Units_UnitYr]) AS UnitModelYear,
  vw_ATF_Avail_Signed_FAID.Units_UnitMake AS UnitMake,
  vw_ATF_Avail_Signed_FAID.Units_UnitModel AS UnitModel,
  vw_ATF_Avail_Signed_FAID.Units_UnitType AS UnitType,
  Nz([Units_UnitSubtype], "TBD") AS UnitSubType,
  vw_ATF_Avail_Signed_FAID.Units_UnitConfiguration AS UnitConfiguration,
  vw_ATF_Avail_Signed_FAID.Units_Axle AS Axle,
  'Diesel' AS PowerType,
  vw_ATF_Avail_Signed_FAID.Units_UnitEngineType AS EngineType,
  vw_ATF_Avail_Signed_FAID.Units_UnitTransmissionType AS TransmissionType,
  CInt(
    Nz([Units_Weight], 0)
  ) AS GVW,
  vw_ATF_Avail_Signed_FAID.VendorID,
  CCur(
    Nz([Total_Unit_Price__c], 0)
  ) AS UnitCost,
  CCur(
    Nz([Payment__c], 0)
  ) AS SchRent,
  CCur(
    Nz([Payment__c], 0)* 0.8
  ) AS FirstExtendedRent,
  CCur(
    Nz([Units_UnitSalesResid], 0)
  ) AS SalesResidual,
  CCur(
    Nz([UTA_Amount_per_Unit__c], 0)
  ) AS UTARebate,
  CDate(
    Nz(
      [Units_UnitEstDelDate], "1/1/1900"
    )
  ) AS EstimatedDeliveryDate,
  Nz([Total_UTAs_per_ERA__c], 0) AS NumUTAs,
  CCur(
    Nz([Fee_Dollar__c], 0)
  ) AS LOFee,
  Nz(
    vw_ATF_Avail_Signed_FAID.Fee__c,
    0
  ) AS LOFeePct,
  CCur(
    Nz([Fee_Dollar__c], 0)* [Spec_UnitQty]
  ) AS LOFeeGroup,
  0 AS DeliveryLocationID,
  0 AS GarageLocationID,
  0 AS RegTitleLocationID
FROM
  (
    ATFGroupSpec
    INNER JOIN (
      ATF
      INNER JOIN ATFScheduleSpec ON ATF.ATFID = ATFScheduleSpec.ATFID
    ) ON ATFGroupSpec.ATFSchSpecID = ATFScheduleSpec.ATFSchSpecID
  )
  INNER JOIN vw_ATF_Avail_Signed_FAID ON (
    ATFGroupSpec.GrpFAID = vw_ATF_Avail_Signed_FAID.FAID
  )
  AND (
    ATFScheduleSpec.ATFSchName = vw_ATF_Avail_Signed_FAID.SchName
  )
  AND (
    ATFGroupSpec.ATFGroupSpecDesc = vw_ATF_Avail_Signed_FAID.GroupName
  );
