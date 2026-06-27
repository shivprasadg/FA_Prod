SELECT
  UnitTotalOEC.GroupID,
  UnitTotalOEC.SyndRent,
  Count(UnitTotalOEC.AssetID) AS Units,
  Format(
    Nz([SyndRent], 0),
    "$#,###.##"
  )& " per unit (x " & Count([assetID])& " units)" AS RentLine
FROM
  UnitTotalOEC
GROUP BY
  UnitTotalOEC.GroupID,
  UnitTotalOEC.SyndRent;
