SELECT
  Sum(
    LeaseInOutExtendedBase.OnLease2015
  ) AS UnitsOnLease2015,
  Sum(
    LeaseInOutExtendedBase.Added2015
  ) AS UnitsAdded2015,
  Sum(
    LeaseInOutExtendedBase.Terminated2015
  ) AS UnitsTerminated2015,
  Sum(
    LeaseInOutExtendedBase.Extended2015
  ) AS UnitsExtended2015,
  (
    [UnitsOnLease2015] + [UnitsAdded2015]
  )- [UnitsTerminated2015] AS OnLease2016,
  Sum(
    LeaseInOutExtendedBase.Added2016
  ) AS Added2016,
  Sum(
    LeaseInOutExtendedBase.Terminated2016
  ) AS Terminated2016
FROM
  LeaseInOutExtendedBase;
