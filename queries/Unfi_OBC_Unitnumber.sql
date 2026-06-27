SELECT
  Nz([OBCDataStatus], 0) AS ReportingStatus,
  Units.ClientGroupId,
  Units.UnitID,
  Units.unityr AS ModelYear,
  Units.UnitInServ AS InService,
  Units.UnitStatus,
  Units.UnitUnitNum AS ATLAAS_UnitNum,
  Units.[ObcUnitNumber] AS ATLAAS_PNOBC,
  UNFI_Units_OBC.[UnitNumber] AS UNFIUnitNumber,
  UNFI_Units_OBC.PN AS PNOBC,
  Units.UnitVIN AS AtlaasVIN,
  UNFI_Units_OBC.VIN AS UNFI_Vin
FROM
  (
    Units
    LEFT JOIN UNFI_Units_OBC ON Units.UnitUnitNum = UNFI_Units_OBC.UnitNumber
  )
  LEFT JOIN UMP_Latest ON Units.UnitID = UMP_Latest.UnitId
WHERE
  (
    (
      (Units.ClientGroupId)= 60
    )
  )
ORDER BY
  Nz([OBCDataStatus], 0),
  Units.UnitID;
