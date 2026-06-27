SELECT
  Emps.EmpInt,
  Emps.EmpNm
FROM
  Emps
WHERE
  (
    (
      (Emps.[LO Exec])= True
    )
  )
ORDER BY
  Emps.EmpInt;
