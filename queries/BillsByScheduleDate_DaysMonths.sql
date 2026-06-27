SELECT
  BillsByScheduleDate_Base.BilledSchID,
  IIf(
    [billType] Like '*Perdiem*'
    Or [billType] Like '*PD*',
    DateDiff('d', [BilledFrom], [Billedto])+ 1,
    0
  ) AS PerdiemDays,
  IIf(
    (
      [billType] = 'Perdiem'
      Or [billType] Like '*PD*'
    )
    And DatePart('yyyy', [billedFrom])= 2015,
    DateDiff('d', [BilledFrom], [Billedto] + 1),
    0
  ) AS PDDays2015,
  IIf(
    (
      [billType] = 'Perdiem'
      Or [billType] Like '*PD*'
    )
    And DatePart('yyyy', [billedFrom])= 2014,
    DateDiff('d', [BilledFrom], [Billedto] + 1),
    0
  ) AS PDDays2014,
  IIf(
    (
      [billType] = 'Perdiem'
      Or [billType] Like '*PD*'
    )
    And DatePart('yyyy', [billedFrom])= 2013,
    DateDiff('d', [BilledFrom], [Billedto] + 1),
    0
  ) AS PDDays2013,
  IIf(
    [billType] Like '*Remit*'
    Or [billType] Like '*Warr*'
    Or [BILLTYPE] = 'AnnRent',
    DateDiff('m', [BilledFrom], [Billedto]),
    0
  ) AS RemitMths,
  IIf(
    (
      [billType] Like '*remit*'
      Or [billType] Like '*Warr*'
      Or [billtype] = 'AnnRent'
    )
    And DatePart('yyyy', [billedFrom])= 2015,
    DateDiff('m', [BilledFrom], [Billedto]),
    0
  ) AS Remit2015,
  IIf(
    (
      [billType] Like '*remit*'
      Or [billType] Like '*Warr*'
      Or [billtype] = 'AnnRent'
    )
    And DatePart('yyyy', [billedFrom])= 2014,
    DateDiff('m', [BilledFrom], [Billedto]),
    0
  ) AS Remit2014,
  IIf(
    (
      [billType] Like '*remit*'
      Or [billType] Like '*Warr*'
      Or [billtype] = 'AnnRent'
    )
    And DatePart('yyyy', [billedFrom])= 2013,
    DateDiff('m', [BilledFrom], [Billedto]),
    0
  ) AS Remit2013,
  IIf(
    ([billType] Like '*Fee*')
    And DatePart('yyyy', [billedFrom])= 2015,
    DateDiff('d', [BilledFrom], [Billedto]),
    0
  ) AS FeeOnly2015,
  IIf(
    ([billType] Like '*Fee*')
    And DatePart('yyyy', [billedFrom])= 2014,
    DateDiff('d', [BilledFrom], [Billedto]),
    0
  ) AS FeeOnly2014,
  IIf(
    ([billType] Like '*Fee*')
    And DatePart('yyyy', [billedFrom])= 2013,
    DateDiff('d', [BilledFrom], [Billedto]),
    0
  ) AS FeeOnly2013
FROM
  BillsByScheduleDate_Base
WHERE
  (
    (
      (
        BillsByScheduleDate_Base.BillType
      )<> ''
      Or (
        BillsByScheduleDate_Base.BillType
      ) Is Not Null
    )
  );
