SELECT
  RemarketingUnitToBuyerScoreMatched.BuyerScore AS SCR,
  RemarketingUnitToBuyerScoreMatched.TypePurchased AS LastBuy,
  RemarketingUnitToBuyerScoreMatched.DaysAgo AS Days,
  RemarketingUnitToBuyerScoreMatched.MilesFromBuyer AS Miles,
  RemarketingUnitToBuyerScoreMatched.UnitID,
  RemarketingUnitToBuyerScoreMatched.UnitStatus AS Sts,
  RemarketingUnitToBuyerScoreMatched.UnitType,
  RemarketingUnitToBuyerScoreMatched.UnitYr AS [Year],
  UnitMake & ' ' & UnitModel & ' ' & Axle AS UnitDesc,
  Format(
    Nz(TargetSale, 0),
    '$#,###'
  ) AS Price,
  RemarketingUnitToBuyerScoreMatched.ParkDateProjected AS [Proj'd Park],
  RemarketingUnitToBuyerScoreMatched.LCity AS [Parked City],
  RemarketingUnitToBuyerScoreMatched.LState AS St,
  RemarketingUnitToBuyerScoreMatched.UnitVin,
  RemarketingUnitToBuyerScoreMatched.ODORead,
  Format(ODOReadDate, 'mm/dd/yy') AS ODODate,
  RemarketingUnitToBuyerScoreMatched.ODOSource,
  RemarketingUnitToBuyerScoreMatched.Latitude,
  RemarketingUnitToBuyerScoreMatched.Longitude,
  RemarketingUnitToBuyerScoreMatched.HomeLat,
  RemarketingUnitToBuyerScoreMatched.HomeLog,
  RemarketingUnitToBuyerScoreMatched.clientID
FROM
  RemarketingUnitToBuyerScoreMatched;
