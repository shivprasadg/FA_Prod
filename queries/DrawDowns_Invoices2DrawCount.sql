SELECT
  DrawDowns_Invoices_Drawn.Invoice,
  Count(
    DrawDowns_Invoices_Drawn.DrawDownId
  ) AS NumOfDraws
FROM
  DrawDowns_Invoices_Drawn
  INNER JOIN DrawDowns ON DrawDowns_Invoices_Drawn.DrawDownId = DrawDowns.DrawDownId
GROUP BY
  DrawDowns_Invoices_Drawn.Invoice
ORDER BY
  DrawDowns_Invoices_Drawn.Invoice;
