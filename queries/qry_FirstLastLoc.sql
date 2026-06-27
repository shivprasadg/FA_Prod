SELECT
  LocChangeT.ID,
  First(LocChangeT.LocChDt) AS FirstCHdt1,
  First(LocChangeT.LocChDt) AS LastCHdt2,
  First(LocChangeT.OldUnitloc) AS FirstOfOldUnitloc,
  First(LocChangeT.OLocationStreet) AS FirstOldStreet,
  First(LocChangeT.OCity) AS FirstOldCity,
  First(LocChangeT.OCounty) AS FirstOldCounty,
  First(LocChangeT.OSt) AS FirstOldState,
  First(LocChangeT.Ozip) AS FirstOldZip,
  First(LocChangeT.OldTxstatus) AS FirstOldTaxStatus,
  First(LocChangeT.OTaxNotes) AS FirstoldTaxNote,
  Last(LocChangeT.UnitsFormloc) AS LastOfUnitsFormloc,
  Last(LocChangeT.NLocationStreet) AS LastNewStreet,
  Last(LocChangeT.Newtaxstatus) AS LastNewTaxStatus,
  Last(LocChangeT.NewTaxnotes) AS LastNewTaxNote,
  Last(LocChangeT.NCity) AS NewLocation,
  Last(LocChangeT.NCounty) AS LastNewCounty,
  Last(LocChangeT.NSt) AS LastNewState,
  Last(LocChangeT.Nzip) AS LastNewZip,
  First(LocChangeT.LocRptDt) AS MinRptChDate,
  Last(LocChangeT.LocRptDt) AS MaxRptChDate,
  LocChangeT.unitref
FROM
  LocChangeT
GROUP BY
  LocChangeT.ID,
  LocChangeT.unitref
HAVING
  (
    (
      (LocChangeT.unitref)= [forms]![UnitsForm]![unitid]
    )
  );
