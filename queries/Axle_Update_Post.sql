UPDATE
  Units
  INNER JOIN Axle_Update_Temp ON Units.UnitID = Axle_Update_Temp.UnitID
SET
  Units.Axle = [Axle_Update_Temp].[newaxel];
