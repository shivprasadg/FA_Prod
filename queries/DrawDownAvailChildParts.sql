SELECT
  DCP.ChildPartID,
  DCP.AssetID,
  DCP.VIN,
  DCP.Unitnum,
  DCP.ChildInvoiceNum AS Invoice,
  DCP.ChildType,
  DCP.ChildCost,
  DCP.SGrpID AS GroupID,
  DCP.ChildMake,
  DCP.ChildVendorID,
  DCP.Schedule,
  DCP.FAID,
  DCP.SchID,
  DCP.ChildDrawDownId
FROM
  vw_Sch_Grp_Units_ChildParts AS DCP;
