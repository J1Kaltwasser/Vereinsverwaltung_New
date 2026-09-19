CREATE TABLE [tbl_Mitglied_Beiträge] (
  [BeitragID] AUTOINCREMENT CONSTRAINT [PrimaryKey] PRIMARY KEY UNIQUE NOT NULL,
  [MitgliedID] LONG,
  [Jahr] LONG,
  [Betrag] CURRENCY,
  [Eingangsdatum] DATETIME,
  [Eingangsbetrag] CURRENCY,
  [Zahlungsdatum] DATETIME,
  [Status] VARCHAR (20),
  [Anmerkungen] LONGTEXT,
  [Created_Date] DATETIME,
  [Modified_Date] DATETIME
)
