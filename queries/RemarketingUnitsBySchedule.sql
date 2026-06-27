SELECT
  DLookUp(
    "ClientShNm", "Clients", "ClientshNm='" & [schGrp].[Assignee_Bank] & "'"
  ) AS AssigneeBank,
  DLookUp(
    "ClientID", "Clients", "ClientshNm='" & [schGrp].[Assignee_Bank] & "'"
  ) AS AssigneeBankID,
  Units.UnitID,
  Units.Unitunitnum,
  Units.Unitvin,
  Schedule.SchID,
  Schedule.SchNo,
  Units.UnitStatus,
  Rmkt.*,
  DLookUp(
    "ClientShNm",
    "Clients",
    "ClientID=" & Nz([Rmkt].[IFBankID], 0)
  ) AS IFBank,
  CCur(
    IIf(
      [BuyOutCost] > 0,
      [BuyOutCost] - Nz([IFAmount], 0),
      0
    )
  ) AS FAAmount
FROM
  Schedule
  INNER JOIN (
    (
      SchGrp
      INNER JOIN Units ON SchGrp.SGrpID = Units.GroupREF
    )
    INNER JOIN Rmkt ON Units.UnitID = Rmkt.UnitRef
  ) ON Schedule.SchID = SchGrp.SGrpSchID
WHERE
  (
    (
      (Units.UnitStatus)<> "A"
      And (Units.UnitStatus)<> "B"
      And (Units.UnitStatus)<> "T"
    )
  );
