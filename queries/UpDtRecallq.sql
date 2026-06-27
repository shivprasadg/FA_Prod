UPDATE
  Recalls
  INNER JOIN updtrecall ON Recalls.RCUnitID = updtrecall.RCUnitID
SET
  Recalls.RCInspDt = [updtrecall]![RCInspDt],
  Recalls.RCRepairReq = [updtrecall]![RCRepairReq];
