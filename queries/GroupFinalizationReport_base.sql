SELECT
  vw_SixKeys.ClientShortName AS Client,
  vw_SixKeys.Schedule,
  vw_SixKeys.UnitGroup,
  Count(vw_SixKeys.UnitId) AS Units,
  SchGrp.AssignmentDate,
  SchGrp.Assigned_Indicator,
  SchGrp.SynFee,
  SchGrp.Fee_Projected,
  SchGrp.SGrpSalesIndex,
  SchGrp.SGrpIndex,
  SchGrp.SgrpYield,
  SchGrp.SgrpSalesYield,
  SchGrp.SGrpRALB,
  Schedule.SchIn,
  IIf(
    [SchGrp].[SGRPNoPD] =-1
    And Sum(
      Nz([UnitPDRent], 0)
    )= 0,
    DLookUp(
      "Person", "GroupFinalizedPersons",
      "Fieldname='PDRent'"
    )
  ) AS PDiemRent,
  IIf(
    Nz([SchGrp].[BillToAddressID], "")= "",
    DLookUp(
      "Person", "GroupFinalizedPersons",
      "Fieldname='BillToAddressID'"
    )
  ) AS BillToAddress,
  IIf(
    Nz([SchGrp].[SgrpCommInd], "")= "",
    DLookUp(
      "Person", "GroupFinalizedPersons",
      "Fieldname='SgrpCommInd'"
    )
  ) AS BLCDBasis,
  IIf(
    Sum(
      Nz([UnitRent], 0)
    )= 0,
    DLookUp(
      "Person", "GroupFinalizedPersons",
      "Fieldname='UnitRent'"
    )
  ) AS UnitSyndRent,
  IIf(
    Sum(
      Nz([UnitSynResid], 0)
    )= 0,
    DLookUp(
      "Person", "GroupFinalizedPersons",
      "Fieldname='UnitSynResid'"
    )
  ) AS UnitResidual,
  Schedule.SchID,
  SchGrp.SGrpID
FROM
  (
    (
      SchGrp
      INNER JOIN vw_SixKeys ON SchGrp.SGrpID = vw_SixKeys.SGrpID
    )
    INNER JOIN Schedule ON vw_SixKeys.SchID = Schedule.SchID
  )
  INNER JOIN Units ON vw_SixKeys.UnitID = Units.UnitID
GROUP BY
  vw_SixKeys.ClientShortName,
  vw_SixKeys.Schedule,
  vw_SixKeys.UnitGroup,
  SchGrp.AssignmentDate,
  SchGrp.Assigned_Indicator,
  SchGrp.SynFee,
  SchGrp.Fee_Projected,
  SchGrp.SGrpSalesIndex,
  SchGrp.SGrpIndex,
  SchGrp.SgrpYield,
  SchGrp.SgrpSalesYield,
  SchGrp.SGrpRALB,
  Schedule.SchIn,
  IIf(
    Nz([SchGrp].[BillToAddressID], "")= "",
    DLookUp(
      "Person", "GroupFinalizedPersons",
      "Fieldname='BillToAddressID'"
    )
  ),
  IIf(
    Nz([SchGrp].[SgrpCommInd], "")= "",
    DLookUp(
      "Person", "GroupFinalizedPersons",
      "Fieldname='SgrpCommInd'"
    )
  ),
  vw_SixKeys.MLOrig,
  SchGrp.SGRPNoPD,
  Schedule.SchID,
  SchGrp.SGrpID
HAVING
  (
    (
      (SchGrp.Assigned_Indicator)= "F"
      Or (SchGrp.Assigned_Indicator)= "P"
    )
    AND (
      (SchGrp.SynFee)= 0
    )
    AND (
      (vw_SixKeys.MLOrig)= "FA"
    )
  )
  OR (
    (
      (SchGrp.Assigned_Indicator)= "F"
      Or (SchGrp.Assigned_Indicator)= "P"
    )
    AND (
      (SchGrp.Fee_Projected)= 0
    )
    AND (
      (vw_SixKeys.MLOrig)= "FA"
    )
  )
  OR (
    (
      (SchGrp.Assigned_Indicator)= "F"
      Or (SchGrp.Assigned_Indicator)= "P"
    )
    AND (
      (SchGrp.SGrpSalesIndex) Is Null
    )
    AND (
      (vw_SixKeys.MLOrig)= "FA"
    )
  )
  OR (
    (
      (SchGrp.Assigned_Indicator)= "F"
      Or (SchGrp.Assigned_Indicator)= "P"
    )
    AND (
      (SchGrp.SGrpIndex) Is Null
    )
    AND (
      (vw_SixKeys.MLOrig)= "FA"
    )
  )
  OR (
    (
      (SchGrp.Assigned_Indicator)= "F"
      Or (SchGrp.Assigned_Indicator)= "P"
    )
    AND (
      (SchGrp.SgrpYield) Is Null
    )
    AND (
      (vw_SixKeys.MLOrig)= "FA"
    )
  )
  OR (
    (
      (SchGrp.Assigned_Indicator)= "F"
      Or (SchGrp.Assigned_Indicator)= "P"
    )
    AND (
      (SchGrp.SgrpSalesYield) Is Null
    )
    AND (
      (vw_SixKeys.MLOrig)= "FA"
    )
  )
  OR (
    (
      (SchGrp.Assigned_Indicator)= "F"
      Or (SchGrp.Assigned_Indicator)= "P"
    )
    AND (
      (SchGrp.SGrpRALB) Is Null
    )
    AND (
      (vw_SixKeys.MLOrig)= "FA"
    )
  )
  OR (
    (
      (SchGrp.Assigned_Indicator)= "F"
      Or (SchGrp.Assigned_Indicator)= "P"
    )
    AND (
      (Schedule.SchIn) Is Null
    )
    AND (
      (vw_SixKeys.MLOrig)= "FA"
    )
  )
  OR (
    (
      (SchGrp.Assigned_Indicator)= "F"
      Or (SchGrp.Assigned_Indicator)= "P"
    )
    AND (
      (
        IIf(
          [SchGrp].[SGRPNoPD] =-1
          And Sum(
            Nz([UnitPDRent], 0)
          )= 0,
          DLookUp(
            "Person", "GroupFinalizedPersons",
            "Fieldname='PDRent'"
          )
        )
      ) Is Not Null
    )
    AND (
      (vw_SixKeys.MLOrig)= "FA"
    )
  )
  OR (
    (
      (SchGrp.Assigned_Indicator)= "F"
      Or (SchGrp.Assigned_Indicator)= "P"
    )
    AND (
      (
        IIf(
          Nz([SchGrp].[BillToAddressID], "")= "",
          DLookUp(
            "Person", "GroupFinalizedPersons",
            "Fieldname='BillToAddressID'"
          )
        )
      ) Is Not Null
    )
    AND (
      (vw_SixKeys.MLOrig)= "FA"
    )
  )
  OR (
    (
      (SchGrp.Assigned_Indicator)= "F"
      Or (SchGrp.Assigned_Indicator)= "P"
    )
    AND (
      (
        IIf(
          Nz([SchGrp].[SgrpCommInd], "")= "",
          DLookUp(
            "Person", "GroupFinalizedPersons",
            "Fieldname='SgrpCommInd'"
          )
        )
      ) Is Not Null
    )
    AND (
      (vw_SixKeys.MLOrig)= "FA"
    )
  )
  OR (
    (
      (SchGrp.Assigned_Indicator)= "F"
      Or (SchGrp.Assigned_Indicator)= "P"
    )
    AND (
      (
        IIf(
          Sum(
            Nz([UnitRent], 0)
          )= 0,
          DLookUp(
            "Person", "GroupFinalizedPersons",
            "Fieldname='UnitRent'"
          )
        )
      ) Is Not Null
    )
    AND (
      (vw_SixKeys.MLOrig)= "FA"
    )
  )
  OR (
    (
      (SchGrp.Assigned_Indicator)= "F"
      Or (SchGrp.Assigned_Indicator)= "P"
    )
    AND (
      (
        IIf(
          Sum(
            Nz([UnitSynResid], 0)
          )= 0,
          DLookUp(
            "Person", "GroupFinalizedPersons",
            "Fieldname='UnitSynResid'"
          )
        )
      ) Is Not Null
    )
    AND (
      (vw_SixKeys.MLOrig)= "FA"
    )
  )
ORDER BY
  vw_SixKeys.ClientShortName,
  vw_SixKeys.Schedule,
  vw_SixKeys.UnitGroup;
