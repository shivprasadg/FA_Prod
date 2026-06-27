SELECT
  Notes.noteUser,
  Clients.clientID,
  Clients.clientcompanyname,
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
  ) AS Last15p,
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
  ) AS Last30p,
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
  ) AS Last45p,
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
  ) AS Last60p
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
      (Notes.CallType)= "phone call"
    )
  )
GROUP BY
  Notes.noteUser,
  Clients.clientID,
  Clients.clientcompanyname
ORDER BY
  Notes.noteUser,
  Clients.clientcompanyname;
