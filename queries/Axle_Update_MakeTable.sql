SELECT
  Axle_Update_PartA.UnitID,
  Axle_Update_PartA.UnitVIN,
  Axle_Update_PartA.Axle,
  Axle_Update_PartA.UnitMake,
  Axle_Update_PartA.UnitModel,
  Axle_Update_PartA.UnitYr,
  Axle_Update_PartB.NewAxel INTO Axle_Update_Temp
FROM
  Axle_Update_PartA
  LEFT JOIN Axle_Update_PartB ON Axle_Update_PartA.VinPrefix = Axle_Update_PartB.VinPreFix_Decoder
WHERE
  (
    (
      (Axle_Update_PartA.Axle)= "TBD"
    )
    AND (
      (Axle_Update_PartB.NewAxel) Like "*/A"
    )
  );
