SELECT
  Emps.empint AS SalespersonID,
  Emps.empnm AS SPInitials,
  Emps.[LO Exec],
  Emps.FS
FROM
  Emps
WHERE
  (
    (
      (Emps.[LO Exec])= True
    )
  )
  OR (
    (
      (Emps.FS)= True
    )
  )
ORDER BY
  Emps.empnm;
