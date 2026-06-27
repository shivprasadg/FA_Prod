SELECT
  Left([VIN], 11) AS VinPreFix_Decoder,
  UnitVinDecoder.DRIVE_LINE_TYPE AS Expr1,
  IIf(
    [DRIVE_LINE_TYPE] = "6x4",
    "T/A",
    IIf(
      [DRIVE_LINE_TYPE] = "4x2", "S/A", [DRIVE_LINE_TYPE]
    )
  ) AS NewAxel,
  UnitVinDecoder.MAKE,
  UnitVinDecoder.MODEL
FROM
  UnitVinDecoder
GROUP BY
  Left([VIN], 11),
  UnitVinDecoder.DRIVE_LINE_TYPE,
  IIf(
    [DRIVE_LINE_TYPE] = "6x4",
    "T/A",
    IIf(
      [DRIVE_LINE_TYPE] = "4x2", "S/A", [DRIVE_LINE_TYPE]
    )
  ),
  UnitVinDecoder.MAKE,
  UnitVinDecoder.MODEL
HAVING
  (
    (
      (
        [UnitVinDecoder].[DRIVE_LINE_TYPE]
      )<> ""
    )
  );
