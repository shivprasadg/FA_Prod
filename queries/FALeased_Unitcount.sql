SELECT
  VW_Clients_ML_FAID_SCH_Groups_count.ClientGroupId,
  VW_Clients_ML_FAID_SCH_Groups_count.clientID,
  VW_Clients_ML_FAID_SCH_Groups_count.ClientShortName,
  VW_Clients_ML_FAID_SCH_Groups_count.CompanyName,
  Sum(
    VW_Clients_ML_FAID_SCH_Groups_count.Units
  ) AS UnitCount,
  VW_Clients_ML_FAID_SCH_Groups_count.MLOrig
FROM
  VW_Clients_ML_FAID_SCH_Groups_count
GROUP BY
  VW_Clients_ML_FAID_SCH_Groups_count.ClientGroupId,
  VW_Clients_ML_FAID_SCH_Groups_count.clientID,
  VW_Clients_ML_FAID_SCH_Groups_count.ClientShortName,
  VW_Clients_ML_FAID_SCH_Groups_count.CompanyName,
  VW_Clients_ML_FAID_SCH_Groups_count.MLOrig
HAVING
  (
    (
      (
        VW_Clients_ML_FAID_SCH_Groups_count.MLOrig
      )= "FA"
    )
  );
