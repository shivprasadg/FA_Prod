SELECT
  RALTemp1.Letterdate,
  vw_RALAttachment.*,
  GroupsNoinSchedule([vw_RALAttachment].[SchID]) AS GroupsIn
FROM
  RALTemp1
  INNER JOIN vw_RALAttachment ON RALTemp1.SchID = vw_RALAttachment.SchID;
