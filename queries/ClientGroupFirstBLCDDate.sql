SELECT
  Min(Clients.ClientBecamedate) AS AwardedDate,
  Min(vw_SixKeys.BLCD) AS MinOfBLCD,
  vw_SixKeys.ClientShortName AS ClientGroupName,
  Count(vw_SixKeys.AssetID) AS Units,
  IIf(
    AwardedDate < #2/1/2009#, [CLientGroupName],
    Null
  ) AS 2009,
  IIf(
    AwardedDate < #2/1/2009#,
    Count([AssetID]),
    Null
  ) AS [2009 Units],
  IIf(
    AwardedDate Between #2/1/2009#
    And #2/1/2011#,
    [CLientGroupName], Null
  ) AS 2010,
  IIf(
    AwardedDate Between #2/1/2009#
    And #2/1/2011#,
    Count([AssetId]),
    Null
  ) AS [2010 Units],
  IIf(
    AwardedDate Between #2/1/2010#
    And #2/1/2012#,
    [CLientGroupName], Null
  ) AS 2011,
  IIf(
    AwardedDate Between #2/1/2011#
    And #2/1/2012#,
    Count([AssetId]),
    Null
  ) AS [2011 Units],
  IIf(
    AwardedDate Between #2/1/2012#
    And #2/1/2013#,
    [CLientGroupName], Null
  ) AS 2012,
  IIf(
    AwardedDate Between #2/1/2012#
    And #2/1/2013#,
    Count([AssetId]),
    Null
  ) AS [2012 Units],
  IIf(
    AwardedDate Between #2/1/2013#
    And #2/1/2014#,
    [CLientGroupName], Null
  ) AS 2013,
  IIf(
    AwardedDate Between #2/1/2013#
    And #2/1/2014#,
    Count([AssetId]),
    Null
  ) AS [2013 Units],
  IIf(
    AwardedDate Between #2/1/2014#
    And #2/1/2015#,
    [CLientGroupName], Null
  ) AS 2014,
  IIf(
    AwardedDate Between #2/1/2014#
    And #2/1/2015#,
    Count([AssetId]),
    Null
  ) AS [2014 Units],
  IIf(
    AwardedDate Between #2/1/2015#
    And #2/1/2016#,
    [CLientGroupName], Null
  ) AS 2015,
  IIf(
    AwardedDate Between #2/1/2015#
    And #2/1/2016#,
    Count([AssetId]),
    Null
  ) AS [2015 Units],
  IIf(
    AwardedDate Between #2/1/2016#
    And #2/1/2017#,
    [CLientGroupName], Null
  ) AS 2016,
  IIf(
    AwardedDate Between #2/1/2016#
    And #2/1/2017#,
    Count([AssetId]),
    Null
  ) AS [2016 Units]
FROM
  (
    vw_SixKeys
    INNER JOIN ProposalsSummaryQuery ON vw_SixKeys.clientID = ProposalsSummaryQuery.clientRef
  )
  INNER JOIN Clients ON vw_SixKeys.clientID = Clients.clientID
WHERE
  (
    (
      (
        ProposalsSummaryQuery.dealOutcome
      )= "Awarded"
    )
  )
GROUP BY
  vw_SixKeys.ClientShortName,
  vw_SixKeys.MLOrig
HAVING
  (
    (
      (vw_SixKeys.MLOrig)= "FA"
    )
  )
ORDER BY
  Min(Clients.ClientBecamedate) DESC;
