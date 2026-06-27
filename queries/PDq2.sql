SELECT
  (
    [PDDAYS] * [PDRENT] + IIf(
      IsNull([SLVRENT]),
      0,
      [slvrent]
    )
  )*(
    1 + IIf(
      [SCHFSL] = True
      Or [UNITTAXEXEMPT] = True
      Or [STTAXTYPE] <> "RR",
      0, [LOCATIONS].[LTaxRate]
    )
  ) AS LineTot,
  PDQ1a.*,
  PDQ1a.LCity
FROM
  PDQ1a;
