SELECT
  Notes.Noteuser,
  Clients.clientID,
  Clients.clientCompanyName,
  Count(Notes.noteid) AS CountOfnoteid,
  Sum(
    IIf(
      IIf(
        IsNull([NoteEctDt]),
        [notedatetime],
        [noteectdt]
      )< Date()
      And IIf(
        IsNull([NoteEctDt]),
        [notedatetime],
        [noteectdt]
      )>= DateAdd(
        "d",
        Date(),
        -15
      ),
      1,
      0
    )
  ) AS Last15E,
  Sum(
    IIf(
      IIf(
        IsNull([NoteEctDt]),
        [notedatetime],
        [noteectdt]
      )< DateAdd(
        "d",
        Date(),
        -15
      )
      And IIf(
        IsNull([NoteEctDt]),
        [notedatetime],
        [noteectdt]
      )>= DateAdd(
        "d",
        Date(),
        -30
      ),
      1,
      0
    )
  ) AS Last30E,
  Sum(
    IIf(
      IIf(
        IsNull([NoteEctDt]),
        [notedatetime],
        [noteectdt]
      )< DateAdd(
        "d",
        Date(),
        -30
      )
      And IIf(
        IsNull([NoteEctDt]),
        [notedatetime],
        [noteectdt]
      )>= DateAdd(
        "d",
        Date(),
        -45
      ),
      1,
      0
    )
  ) AS Last45E,
  Sum(
    IIf(
      IIf(
        IsNull([NoteEctDt]),
        [notedatetime],
        [noteectdt]
      )< DateAdd(
        "d",
        Date(),
        -45
      )
      And IIf(
        IsNull([NoteEctDt]),
        [notedatetime],
        [noteectdt]
      )>= DateAdd(
        "d",
        Date(),
        -60
      ),
      1,
      0
    )
  ) AS Last60E
FROM
  (
    Notes
    INNER JOIN Clients ON Notes.clientRef = Clients.clientID
  )
  INNER JOIN Salesperson ON Clients.LOinitial = Salesperson.SPInitials
WHERE
  (
    (
      (
        IIf(
          IsNull([NoteEctDt]),
          [notedatetime],
          [noteectdt]
        )
      )>= DateAdd(
        "d",
        Date(),
        -60
      )
    )
    AND (
      (Notes.CallType)= "e-mail"
    )
  )
GROUP BY
  Notes.Noteuser,
  Clients.clientID,
  Clients.clientCompanyName
ORDER BY
  Notes.Noteuser,
  Clients.clientCompanyName;
