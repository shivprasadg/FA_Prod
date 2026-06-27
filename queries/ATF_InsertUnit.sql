PARAMETERS GroupIDx Long,
UnitSpecX Long; INSERT INTO Units (
  ATFGroupSpecID, ATFUnitSpecID, Axle,
  UnitAxleProperty, ClientID, ClientGroupID,
  UnitDeliveryLocationRef, UnitBuildDate,
  UnitEstDelDate, FirstExtRent, UnitGVW,
  UnitGarageLocationRef, LegalDescription,
  MTMRent, UNITPDRENT, UnitRegLocationRef,
  UnitSalesResid, UNITSRENT, GroupREF,
  UnitAddDt, UNITRENT, UnitSynResid,
  UnitTransmissionType, UnitConfiguration,
  UnitFinalCost, UnitDescriptor, UnitMake,
  UnitModel, UnitYr, UnitStatus, UnitSubType,
  UnitType, UnitUOM, VendorREF
)
SELECT
  DISTINCT ATFUnitSpec.ATFGroupSpecID,
  ATFUnitSpec.ATFUnitSpecID,
  ATFUnitSpec.Axle,
  ATFUnitSpec.AxleProperty,
  ATFUnitSpec.ClientID,
  ATFUnitSpec.ClientGroupID,
  ATFUnitSpec.DeliveryLocationID,
  ATFUnitSpec.EstimatedBuildDate,
  ATFUnitSpec.EstimatedDeliveryDate,
  ATFUnitSpec.FirstExtendedRent,
  ATFUnitSpec.GVW,
  ATFUnitSpec.GarageLocationID,
  ATFUnitSpec.LegalDescription,
  ATFUnitSpec.MTMRent,
  ATFUnitSpec.PerDiemRent,
  ATFUnitSpec.RegTitleLocationID,
  ATFUnitSpec.SalesResidual,
  ATFUnitSpec.SchRent,
  SchGrp.SGrpID,
  ATFUnitSpec.SpecInserted,
  ATFUnitSpec.SyndRent,
  ATFUnitSpec.SyndResidual,
  ATFUnitSpec.TransmissionType,
  ATFUnitSpec.UnitConfiguration,
  ATFUnitSpec.UnitCost,
  ATFUnitSpec.UnitDescriptor,
  ATFUnitSpec.UnitMake,
  ATFUnitSpec.UnitModel,
  ATFUnitSpec.UnitModelYear,
  'B' AS UnitStatus,
  ATFUnitSpec.UnitSubType,
  ATFUnitSpec.UnitType,
  ATFUnitSpec.UOM,
  ATFUnitSpec.VendorID
FROM
  ATFUnitSpec
  INNER JOIN SchGrp ON ATFUnitSpec.ATFGroupSpecID = SchGrp.ATFGroupSpecID
WHERE
  (
    (
      (ATFUnitSpec.ATFUnitSpecID)= [UnitSpecX]
    )
    AND (
      (SchGrp.SGrpID)= [GroupIDx]
    )
  );
