CREATE TABLE [Mitglied_Beiträge] (
  [BeitragID] AUTOINCREMENT CONSTRAINT [PrimaryKey] PRIMARY KEY UNIQUE NOT NULL,
  [Mitgliedsnummer] VARCHAR (10) CONSTRAINT [MitgliederMitglied_Beiträge] REFERENCES [Mitglieder] ([Mitgliedsnummer]) ON UPDATE CASCADE ,
  [Beitragsjahr] LONG,
  [Betrag] CURRENCY,
  [GültigAb] DATETIME,
  [GültigBis] DATETIME,
  [ÄnderungsDatum] DATETIME,
  [ÄnderungsBenutzer] VARCHAR (50),
  [Bemerkung] LONGTEXT
)
