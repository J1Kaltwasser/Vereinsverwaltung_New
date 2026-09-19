CREATE TABLE [tbl_Zuwendungsbestätigungen] (
  [ZuwendungsID] AUTOINCREMENT CONSTRAINT [PrimaryKey] PRIMARY KEY UNIQUE NOT NULL,
  [MitgliedID] LONG,
  [BeitragID] LONG,
  [Jahr] SHORT,
  [Gesamtbetrag] CURRENCY,
  [QuitungNummer] VARCHAR (20) CONSTRAINT [QuitungNummer] UNIQUE,
  [Erstellungsdatum] DATETIME,
  [Status] VARCHAR (20),
  [Versanddatum] DATETIME,
  [Anmerkungen] LONGTEXT,
  [Created_Date] DATETIME
)
