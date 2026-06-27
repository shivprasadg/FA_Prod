SELECT
  Clients_Sch_Group_OEC_Part1.CompanyName,
  Clients_Sch_Group_OEC_Part1.MLNo,
  Clients_Sch_Group_OEC_Part1.Schedule,
  Clients_Sch_Group_OEC_Part1.UnitGroup,
  ScheduleGroupPOTypes.POType,
  ScheduleGroupPOTypes.Ordered,
  ScheduleGroupPOTypes.OrderedDate,
  Clients_Sch_Group_OEC_Part1.Units,
  Sum(UnitDeliveredCost.UnitCost) AS UnitCost,
  Sum(UnitDeliveredCost.ChildCost) AS ChildCost,
  Sum(
    UnitDeliveredCost.DeliveredCost
  ) AS DeliveredCost,
  Sum(UnitDeliveredCost.SyndRent) AS SyndRent,
  IIf(
    [SyndRent] = 0,
    0,
    Round([SyndRent] / [DeliveredCost], 6)
  ) AS LRF,
  UnitDeliveredCost.GroupId
FROM
  (
    (
      Clients_Sch_Group_OEC_Part1
      INNER JOIN UnitDeliveredCost ON Clients_Sch_Group_OEC_Part1.SGrpID = UnitDeliveredCost.GroupId
    )
    INNER JOIN SchGrp ON Clients_Sch_Group_OEC_Part1.SGrpID = SchGrp.SGrpID
  )
  LEFT JOIN ScheduleGroupPOTypes ON Clients_Sch_Group_OEC_Part1.SGrpID = ScheduleGroupPOTypes.GroupID
GROUP BY
  Clients_Sch_Group_OEC_Part1.CompanyName,
  Clients_Sch_Group_OEC_Part1.MLNo,
  Clients_Sch_Group_OEC_Part1.Schedule,
  Clients_Sch_Group_OEC_Part1.UnitGroup,
  ScheduleGroupPOTypes.POType,
  ScheduleGroupPOTypes.Ordered,
  ScheduleGroupPOTypes.OrderedDate,
  Clients_Sch_Group_OEC_Part1.Units,
  UnitDeliveredCost.GroupId,
  Clients_Sch_Group_OEC_Part1.SchID
HAVING
  (
    (
      (UnitDeliveredCost.GroupId) Is Not Null
    )
  )
ORDER BY
  Clients_Sch_Group_OEC_Part1.CompanyName,
  Clients_Sch_Group_OEC_Part1.SchID;
