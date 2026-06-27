SELECT
  Lines.Linebank,
  Lines.Lineid,
  Lines.LineDt,
  Lines.LineType,
  Lines.LineMinLoan,
  [lineamt] + [lineinc1] + [lineinc2] + [lineinc3] + [lineinc4] + [lineinc5] AS linetot,
  IIf(
    IsNull([SumOfPNAmt]),
    0,
    [SumOfPNAmt]
  ) AS Use
FROM
  Lines
  LEFT JOIN PntSumq ON Lines.Lineid = PntSumq.Lineid
WHERE
  (
    (
      (Lines.Linebank)= [forms]![bank_view]![clientid]
    )
  );
