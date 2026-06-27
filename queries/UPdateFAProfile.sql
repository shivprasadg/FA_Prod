UPDATE
  clients
SET
  clients.fleetSize = [updateFAProfile].[fleetDaycabs] + [updateFAProfile].[fleetSleepers],
  clients.creditFAScore = [updateFAProfile].[ScoreCredit] + [updateFAProfile].[ScoreCredit2] + [updateFAProfile].[ScoreFleetSize] + [updateFAProfile].[ScorePercentDC] + [updateFAProfile].[ScoreDCMiles] + [updateFAProfile].[ScoreSLMiles] + [updateFAProfile].[ScoreMaintenance] + [updateFAProfile].[ScoreLifecycle] + [updateFAProfile].[ScoreAcquisition],
  clients.creditFAProfile = IIf(
    [updateFAProfile].[TOTALSCORE] >= 80,
    "Qualifies",
    IIf(
      [updateFAProfile].[TOTALSCORE] >= 70,
      "Review Required", "Does Not Qualify"
    )
  );
