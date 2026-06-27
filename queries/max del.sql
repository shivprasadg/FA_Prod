SELECT
  Units.GroupREF,
  Max(
    IIf(
      IsNull([unitinserv]),
      IIf(
        IsNull([unitestdeldate]),
        [unitbuilddate]
      ),
      [unitinserv]
    )
  ) AS [Maxin-serv],
  Max(
    DateAdd(
      "m",
      1,
      IIf(
        IsNull([unitinserv]),
        IIf(
          IsNull([unitestdeldate]),
          [unitbuilddate]
        ),
        [unitinserv]
      )
    )-1
  ) AS ProjSyn,
  Max(
    IIf(
      IsNull([unitactualdeldate]),
      [unitestdeldate],
      [unitactualdeldate]
    )
  ) AS Deldt,
  Max(Units.UnitAcceptDate) AS MaxAccpt
FROM
  Units
GROUP BY
  Units.GroupREF;
