UPDATE
  SchGrp
  INNER JOIN UnitDescByGroupTemp ON SchGrp.SGrpID = UnitDescByGroupTemp.GroupID
SET
  SchGrp.SgrpDesc = [UnitDescByGroupTemp].[GroupDesc]
WHERE
  (
    (
      (SchGrp.SgrpDesc) Is Null
      Or (SchGrp.SgrpDesc)= ""
    )
  );
