SELECT
  *
FROM
  vw_PayProceeds_Union
WHERE
  (
    (
      ([PayInterim])> 0
    )
    AND (
      ([SchID]) In (2512)
    )
  );
