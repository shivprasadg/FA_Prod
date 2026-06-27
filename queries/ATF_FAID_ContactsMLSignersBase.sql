SELECT
  DISTINCTROW Clients.clientID,
  Clients.ClientShNm,
  MstrLease.MLNo,
  MstrLease.MLOrig,
  Schedule.SchID,
  Schedule.SchNo,
  COA_Signers_Unit.SGrpID,
  COA_Signers_Unit.SGrpNO,
  MstrLease.MLSignerID,
  MstrLease.MLSigner2ID,
  MstrLease.MLSigner3ID,
  MstrLease.MLGuarantor,
  MstrLease.MLGuarSignerID,
  IIf(
    Nz([MLSignerID], 0)+ Nz([MLSigner2ID], 0)+ Nz([MLSigner3ID], 0)= 0,
    Null,
    "Master Lease"
  ) AS SignerGroup1,
  IIf(
    IsNull([MLGuarantor]),
    Null,
    "Guarantor"
  ) AS SignerGroup2,
  IIf(
    Nz([SchSignerId], 0)+ Nz([SchSigner2ID], 0)+ Nz([SchSigner3ID], 0)= 0,
    Null,
    "Schedule"
  ) AS SignerGroup3,
  IIf(
    IsNull([CofASignerID]),
    Null,
    "COA"
  ) AS SignerGroup4,
  Schedule.SchSignerID,
  Schedule.SchSigner2ID,
  Schedule.SchSigner3ID,
  COA_Signers_Unit.CofASignerID,
  MstrLease.MLID
FROM
  (
    (
      Clients
      INNER JOIN MstrLease ON Clients.clientID = MstrLease.Client
    )
    INNER JOIN Schedule ON MstrLease.MLID = Schedule.SchMLRef
  )
  LEFT JOIN COA_Signers_Unit ON Schedule.SchID = COA_Signers_Unit.SchID
GROUP BY
  Clients.clientID,
  Clients.ClientShNm,
  MstrLease.MLNo,
  MstrLease.MLOrig,
  Schedule.SchID,
  Schedule.SchNo,
  COA_Signers_Unit.SGrpID,
  COA_Signers_Unit.SGrpNO,
  MstrLease.MLSignerID,
  MstrLease.MLSigner2ID,
  MstrLease.MLSigner3ID,
  MstrLease.MLGuarantor,
  MstrLease.MLGuarSignerID,
  IIf(
    Nz([MLSignerID], 0)+ Nz([MLSigner2ID], 0)+ Nz([MLSigner3ID], 0)= 0,
    Null,
    "Master Lease"
  ),
  IIf(
    IsNull([MLGuarantor]),
    Null,
    "Guarantor"
  ),
  IIf(
    Nz([SchSignerId], 0)+ Nz([SchSigner2ID], 0)+ Nz([SchSigner3ID], 0)= 0,
    Null,
    "Schedule"
  ),
  IIf(
    IsNull([CofASignerID]),
    Null,
    "COA"
  ),
  Schedule.SchSignerID,
  Schedule.SchSigner2ID,
  Schedule.SchSigner3ID,
  COA_Signers_Unit.CofASignerID,
  MstrLease.MLID
HAVING
  (
    (
      (MstrLease.MLOrig)= "FA"
    )
  )
ORDER BY
  Clients.ClientShNm,
  MstrLease.MLNo,
  Schedule.SchID,
  COA_Signers_Unit.SGrpID;
