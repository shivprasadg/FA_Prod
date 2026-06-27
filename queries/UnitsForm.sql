SELECT
  Units.*,
  Delv.DelvAddr,
  Gar.GarAddr,
  Title.TitleAddr,
  Reg.RegAddr,
  Title.LState AS TitleState,
  LookUp_States.StTaxType,
  Gar.LTaxOption,
  Gar.LTaxRate,
  UnitsFormFilter.*,
  Rmkt.BuyDateEffective
FROM
  (
    (
      (
        (
          (
            UnitsFormFilter
            INNER JOIN (
              Units
              LEFT JOIN (
                SELECT
                  LocationId,
                  LTaxRate,
                  "<B>" & LocationStreet & "</b><Br />" & LAddrLine1 & "<Br />" & IIf(
                    IsNull(LocSubLessee),
                    Null,
                    LocSubLessee & "<Br />"
                  )& [LCity] & ", " & [LState] & " " & [Lzip] & " |<b> County</b> " & LCounty AS GarAddr,
                  LState,
                  LTaxOption
                FROM
                  Locations
              ) AS Gar ON Units.UnitGarageLocationREF = Gar.LocationID
            ) ON UnitsFormFilter.AssetIDx = Units.UnitID
          )
          LEFT JOIN (
            SELECT
              LocationId,
              LTaxRate,
              "<B>" & LocationStreet & "</b><Br />" & LAddrLine1 & "<Br />" & IIf(
                IsNull(LocSubLessee),
                Null,
                LocSubLessee & "<Br />"
              )& [LCity] & ", " & [LState] & " " & [Lzip] & " | <b> County</b> " & LCounty AS DelvAddr
            FROM
              Locations
          ) AS Delv ON Units.UnitDeliveryLocationRef = Delv.LocationID
        )
        LEFT JOIN (
          SELECT
            LocationId,
            LTaxRate,
            "<B>" & LocationStreet & "</b><Br />" & LAddrLine1 & "<Br />" & IIf(
              IsNull(LocSubLessee),
              Null,
              LocSubLessee & "<Br />"
            )& [LCity] & ", " & [LState] & " " & [Lzip] & " | <b> County</b> " & LCounty AS TitleAddr,
            LState
          FROM
            Locations
        ) AS Title ON Units.UnitTitleLocationRef = Title.LocationID
      )
      LEFT JOIN (
        SELECT
          LocationId,
          LTaxRate,
          "<B>" & LocationStreet & "</b><Br />" & LAddrLine1 & "<Br />" & IIf(
            IsNull(LocSubLessee),
            Null,
            LocSubLessee & "<Br />"
          )& [LCity] & ", " & [LState] & " " & [Lzip] & " |<b> County</b> " & LCounty AS RegAddr
        FROM
          Locations
      ) AS Reg ON Units.UnitRegLocationRef = Reg.LocationID
    )
    LEFT JOIN LookUp_States ON Gar.LState = LookUp_States.StateAbbrev
  )
  LEFT JOIN Rmkt ON Units.UnitID = Rmkt.UnitRef;
