PARAMETERS ATFX Long,
MLIDx Long,
FAIDx Text (255); INSERT INTO ATFScheduleSpec (
  ATFSchName, SchFAID, ATFID, MLID, FSLType,
  ATFSchSpecDesc, SchDescription,
  NumGroupsInSch, SchSpecInserted
)
SELECT
  ATF_SF_ScheduleProfile.SchName AS ATFSchName,
  ATF_SF_ScheduleProfile.FAID AS SchFAID,
  [ATFX] AS ATFID,
  [MLIDx] AS MLID,
  "Yes" AS FSLType,
  ATF_SF_ScheduleProfile.SchName AS ATFSchSpecDesc,
  "(" & DSum(
    "UnitsInGroup", "ATF_SF_GroupProfile",
    "SchName='" & [SchName] & "' AND GrpFAID='" & [FAIDx] & "'"
  )& ") " & SplitString([SchShortDesc], ") ", 1) AS SchDescription,
  DLookUp(
    "GroupsIn", "ATF_SF_GroupProfile",
    "SchName='" & [SchName] & "' AND GrpFAID='15-8168 '"
  ) AS NumGroupsInSch,
  Now() AS SchSpecInserted
FROM
  ATF_SF_ScheduleProfile
WHERE
  (
    (
      (ATF_SF_ScheduleProfile.FAID)= [FAIDx]
    )
  );
