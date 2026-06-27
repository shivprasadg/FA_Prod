SELECT
  TempSC.clientID,
  CFOVISTQ.CFOVisti,
  Treasvisits.TreasVisit,
  FinAnal.AnalVisit,
  Trans.TransMgrVisit,
  FleetMgr.FMgrvisit,
  ShopMgrq.ShopMgrVisit
FROM
  (
    (
      (
        (
          (
            CFOVISTQ
            RIGHT JOIN TempSC ON CFOVISTQ.clientID = TempSC.clientID
          )
          LEFT JOIN Treasvisits ON TempSC.clientID = Treasvisits.clientID
        )
        LEFT JOIN Trans ON TempSC.clientID = Trans.clientID
      )
      LEFT JOIN FleetMgr ON TempSC.clientID = FleetMgr.clientID
    )
    LEFT JOIN FinAnal ON TempSC.clientID = FinAnal.clientID
  )
  LEFT JOIN ShopMgrq ON TempSC.clientID = ShopMgrq.clientID;
