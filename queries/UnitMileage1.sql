SELECT
  vw_UMP_Mileage.UnitID AS AssetId,
  vw_UMP_Mileage.UnitNumber AS UnitNo,
  Max(
    IIf(
      [OBCDataStatus] = 6, 0, [OdometerEnd]
    )
  ) AS Odometer,
  Round(
    IIf(
      Nz(
        Max([MilesLTD]),
        0
      )> 0
      And Nz(
        Max([FuelLTD]),
        0
      )> 0,
      (
        Max([MilesLTD])/ Max([FuelLTD])
      ),
      0
    ),
    2
  ) AS MPGOverAll,
  Round(
    IIf(
      Nz(
        Max([MilesLTD]),
        0
      )> 0
      And Nz(
        Max([FuelmovingLTD]),
        0
      )> 0,
      (
        Max([MilesLTD])/ Max([FuelMovingLTD])
      ),
      0
    ),
    2
  ) AS MPGMoving,
  Max(vw_UMP_Mileage.LogActivityDate) AS ODOReadDate,
  Count(vw_UMP_Mileage.UnitId) AS MonthsReported,
  Max(vw_UMP_Mileage.OBCDataStatus) AS OBCStatus,
  0 AS OdoError,
  Max(
    IIf(
      [OBCDataStatus] = 7, [vw_UMP_Mileage].[MPGMoving],
      0
    )
  ) AS MPGError,
  MileageOdoReadingAdj.FirstReport,
  MileageOdoReadingAdj.LastReport,
  MileageOdoReadingAdj.ODOMiles,
  MileageOdoReadingAdj.MthsRptd,
  MileageOdoReadingAdj.MPM
FROM
  vw_UMP_Mileage
  INNER JOIN MileageOdoReadingAdj ON vw_UMP_Mileage.UnitID = MileageOdoReadingAdj.UnitID
GROUP BY
  vw_UMP_Mileage.UnitID,
  vw_UMP_Mileage.UnitNumber,
  MileageOdoReadingAdj.FirstReport,
  MileageOdoReadingAdj.LastReport,
  MileageOdoReadingAdj.ODOMiles,
  MileageOdoReadingAdj.MthsRptd,
  MileageOdoReadingAdj.MPM
ORDER BY
  vw_UMP_Mileage.UnitID DESC;
