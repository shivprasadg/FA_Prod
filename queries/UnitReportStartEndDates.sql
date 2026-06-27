SELECT
  Units.UnitID,
  Units.UnitStatus,
  Units.UnitInServ,
  IIf(
    [UnitActualDelDate] > Date(),
    Null,
    [UnitActualDelDate]
  ) AS ActualDelDate,
  Units.UnitAcceptDate,
  CDate(
    Nz(
      [UnitInServ],
      Nz(
        [UnitAcceptDate],
        Nz([ActualDelDate], #1/1/1900#)
      )
    )
  ) AS StartDate,
  Rmkt.BuyOutDate,
  IIf(
    [unitstatus] = "O",
    Nz(
      [Inspected2],
      Nz(
        [Inspected2],
        Nz(
          [ReadDate],
          Date()
        )
      )
    ),
    Null
  ) AS LastReadDate,
  Units.UnitOLInd,
  IIf(
    [unitoffleasedt] > Date()
    Or [UnitOLInd] = "P",
    Null,
    [unitoffleasedt]
  ) AS OffLeaseDate,
  Rmkt.ParkDT,
  Rmkt.InvoicePaidDate,
  CDate(
    Nz(
      [OffLeaseDate],
      Nz(
        [BuyOutDate],
        Nz(
          [ParkDT],
          Nz(
            [InvoicePaidDate],
            Nz([LastReadDate], #12/31/2030#)
          )
        )
      )
    )
  ) AS EndDate,
  Units.ReportStartDate,
  Units.ReportEndDate
FROM
  (
    Units
    LEFT JOIN Rmkt ON Units.UnitID = Rmkt.UnitRef
  )
  LEFT JOIN UnitLastMileageRead ON Units.UnitID = UnitLastMileageRead.UnitID;
