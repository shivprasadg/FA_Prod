SELECT
  SchGrp.SGrpSchID AS SchID,
  [SGrpNO] & " - " & Format([BLCD], "mm/dd/yy") AS GBLCD,
  [SGrpNO] & " - " & Format([FundDate], "mm/dd/yy") AS GFundDate,
  [SGrpNO] & " - " & Format([AssignmentDate], "mm/dd/yy") AS GAssignmentDate,
  [SGrpNO] & " - " & [LeaseTermPrimary] AS GTerm,
  [SGrpNO] & " - " & [SGprAllow1] & " | " & [SGprAllow2] AS GMiles,
  [SGrpNO] & " - " & IIf(
    IsDate([sgrpralo]),
    Format([SGrpRALo], "mm/dd/yy"),
    ""
  ) AS GRALOut,
  [SGrpNO] & " - " & IIf(
    IsDate([sgrpralb]),
    Format([SGrpRALB], "mm/dd/yy"),
    ""
  ) AS GRALBack,
  SchGrp.Assignee_Bank AS Assignee,
  TMTSignedDates.TMTSigned
FROM
  SchGrp
  LEFT JOIN TMTSignedDates ON SchGrp.SGrpID = TMTSignedDates.GroupID
GROUP BY
  SchGrp.SGrpSchID,
  [SGrpNO] & " - " & Format([BLCD], "mm/dd/yy"),
  [SGrpNO] & " - " & Format([FundDate], "mm/dd/yy"),
  [SGrpNO] & " - " & Format([AssignmentDate], "mm/dd/yy"),
  [SGrpNO] & " - " & [LeaseTermPrimary],
  [SGrpNO] & " - " & [SGprAllow1] & " | " & [SGprAllow2],
  [SGrpNO] & " - " & IIf(
    IsDate([sgrpralo]),
    Format([SGrpRALo], "mm/dd/yy"),
    ""
  ),
  [SGrpNO] & " - " & IIf(
    IsDate([sgrpralb]),
    Format([SGrpRALB], "mm/dd/yy"),
    ""
  ),
  SchGrp.Assignee_Bank,
  TMTSignedDates.TMTSigned;
