UPDATE
  (
    Clients
    INNER JOIN Deals ON Clients.clientID = Deals.clientRef
  )
  INNER JOIN Documentation ON Deals.dealID = Documentation.dealRef
SET
  Clients.docMLAContactREF = [Documentation].[docMLAContactREF],
  Clients.docMLASigRef = [documentation].[docMLASigRef],
  Clients.docMLASigTitle = [documentation].[docMLASigTitle],
  Clients.docGuarSigREF = [documentation].[docGuarSigREF],
  Clients.docGuarSigrTitle = [documentation].[docGuarSigrTitle],
  Clients.docLeaseID = [documentation].[docLeaseID],
  Clients.docNameVer = [documentation].[docNameVer],
  Clients.docCGS = [documentation].[docCGS],
  Clients.docTaxID = [documentation].[docTaxID]
WHERE
  (
    (
      Not (Clients.docMLAContactREF) Is Null
    )
  )
  Or (
    (
      Not (Clients.docMLASigRef) Is Null
    )
  )
  Or (
    (
      Not (Clients.docMLASigTitle) Is Null
    )
  )
  Or (
    (
      Not (Clients.docGuarSigREF) Is Null
    )
  )
  Or (
    (
      Not (Clients.docGuarSigrTitle) Is Null
    )
  )
  Or (
    (
      Not (Clients.docLeaseID) Is Null
    )
  )
  Or (
    (
      Not (Clients.docNameVer) Is Null
    )
  )
  Or (
    (
      Not (Clients.docCGS) Is Null
    )
  )
  Or (
    (
      Not (Clients.docTaxID) Is Null
    )
  );
