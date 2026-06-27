SELECT
  ATFChildPartSpec.ATFUnitSpecID,
  Sum(ATFChildPartSpec.ChildPartCost) AS ChildPartCost,
  Count(
    ATFChildPartSpec.ATFChildSpecID
  ) AS Parts,
  ATFChildPartSpec.ATFID
FROM
  ATFChildPartSpec
GROUP BY
  ATFChildPartSpec.ATFUnitSpecID,
  ATFChildPartSpec.ATFID;
