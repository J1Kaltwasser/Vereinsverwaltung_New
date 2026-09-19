SELECT
  m.MitgliedID,
  m.Mitgliedsnummer,
  m.Vorname,
  m.Nachname_Name,
  m.Titel,
  IIf(
    m.Vorname Is Not Null
    And m.Vorname <> "",
    m.Vorname & " " & m.Nachname_Name,
    m.Nachname_Name
  ) AS Vollname,
  m.Mitgliedstyp,
  m.Mitgliedsgruppe,
  m.Eintrittsdatum,
  m.Beitrag,
  m.Email,
  m.Telefon,
  m.Status,
  m.Created_Date
FROM
  tbl_Mitglieder AS m
WHERE
  (
    (
      (m.Status) = "Aktiv"
    )
    AND (
      (m.IsDeleted) = False
    )
  )
ORDER BY
  m.Nachname_Name,
  m.Vorname;
