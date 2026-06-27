SELECT
  Awarded_Deals_Client.FAID,
  Awarded_Deals_Client.Client,
  ML_Sch_Grp_FAID_Count_Cost_Rents.UnitCount AS Units,
  ML_Sch_Grp_FAID_Count_Cost_Rents.TotalUnitCost AS OEC,
  Awarded_Deals_Client.LPAFQtr,
  Awarded_Deals_Client.LPAFYear,
  Awarded_Deals_Client.DealQtr,
  Awarded_Deals_Client.DealYear,
  Awarded_Deals_Client.DealUnitCount,
  Awarded_Deals_Client.DealGrandTotal,
  ML_Sch_Grp_FAID_Count_Cost_Rents.OPSQtr,
  ML_Sch_Grp_FAID_Count_Cost_Rents.OPSYear,
  ML_Sch_Grp_FAID_Count_Cost_Rents.UnitCount AS ActualUnitCount,
  ML_Sch_Grp_FAID_Count_Cost_Rents.TotalLORent,
  ML_Sch_Grp_FAID_Count_Cost_Rents.TotalSyndRent,
  ML_Sch_Grp_FAID_Count_Cost_Rents.UnitCost,
  ML_Sch_Grp_FAID_Count_Cost_Rents.ChildPartCost
FROM
  Awarded_Deals_Client
  LEFT JOIN ML_Sch_Grp_FAID_Count_Cost_Rents ON Awarded_Deals_Client.FAID = ML_Sch_Grp_FAID_Count_Cost_Rents.LFAID;
