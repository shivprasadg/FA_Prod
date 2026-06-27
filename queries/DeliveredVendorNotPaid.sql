SELECT
  vw_CustDelvReportBase.ClientShortName,
  vw_CustDelvReportBase.Schedule,
  vw_CustDelvReportBase.UnitGroup,
  vw_CustDelvReportBase.AssetID,
  vw_CustDelvReportBase.UnitDescAbbr,
  vw_CustDelvReportBase.VendorSNm,
  vw_CustDelvReportBase.FundDate,
  vw_CustDelvReportBase.ActualDeliveryDate,
  vw_CustDelvReportBase.[Interim Bank],
  vw_CustDelvReportBase.VendorPaidDate,
  DateDiff(
    "d",
    [ActualDeliveryDate],
    Date()
  ) AS DaysDelivered,
  vw_CustDelvReportBase.BankShortName
FROM
  vw_CustDelvReportBase
WHERE
  (
    (
      (vw_CustDelvReportBase.FundDate) Is Null
    )
    AND (
      Not (
        vw_CustDelvReportBase.ActualDeliveryDate
      ) Is Null
    )
    AND (
      (
        vw_CustDelvReportBase.VendorPaidDate
      ) Is Null
    )
    AND (
      (
        vw_CustDelvReportBase.BankShortName
      )<> "FA LLC"
    )
  )
ORDER BY
  vw_CustDelvReportBase.ClientShortName,
  vw_CustDelvReportBase.Schedule;
