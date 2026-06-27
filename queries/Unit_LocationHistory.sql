SELECT
  LocChangeT.ID,
  LocChangeT.LocChDt AS FirstCHdt1,
  LocChangeT.LocChDt AS LastCHdt2,
  LocChangeT.OldUnitloc,
  LocChangeT.OLocationStreet AS FirstOldStreet,
  LocChangeT.OCity AS FirstOldCity,
  LocChangeT.OCounty AS FirstOldCounty,
  LocChangeT.OSt AS FirstOldState,
  LocChangeT.Ozip AS FirstOldZip,
  LocChangeT.OldTxstatus AS FirstOldTaxStatus,
  LocChangeT.OTaxNotes AS FirstoldTaxNote,
  LocChangeT.NLocationStreet AS LastNewStreet,
  LocChangeT.Newtaxstatus AS LastNewTaxStatus,
  LocChangeT.NewTaxnotes AS LastNewTaxNote,
  LocChangeT.NCity AS NewLocation,
  LocChangeT.NCounty AS LastNewCounty,
  LocChangeT.NSt AS LastNewState,
  LocChangeT.Nzip AS LastNewZip,
  LocChangeT.LocRptDt AS MinRptChDate,
  LocChangeT.LocRptDt AS MaxRptChDate,
  LocChangeT.unitref
FROM
  LocChangeT
ORDER BY
  LocChangeT.ID;
