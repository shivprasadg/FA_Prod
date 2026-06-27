SELECT
  DISTINCT Clients.clientCompanyName,
  MstrLease.MLOrig,
  Schedule.schno,
  SchGrp.SGrpRALO,
  SchGrp.SGrpRALb
FROM
  (
    (
      Clients
      INNER JOIN MstrLease ON Clients.clientID = MstrLease.Client
    )
    INNER JOIN Schedule ON MstrLease.MLID = Schedule.SchMLRef
  )
  INNER JOIN SchGrp ON Schedule.SchID = SchGrp.SGrpSchID
WHERE
  (
    (
      (MstrLease.MLOrig)= "fa"
    )
    AND (
      (SchGrp.SGrpRALb)>= [enter start dt]
      And (SchGrp.SGrpRALb)<= [enterend date]
    )
  )
  OR (
    (
      (SchGrp.SGrpRALO)>= [enter start dt]
      And (SchGrp.SGrpRALO)<= [enterend date]
    )
  )
  OR (
    (
      (Clients.clientCompanyName)= "Southern Wine & Spirits of America, Inc."
    )
  )
  OR (
    (
      (Clients.clientCompanyName) Like "rein*"
    )
  );
