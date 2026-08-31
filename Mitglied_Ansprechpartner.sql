CREATE TABLE [Mitglied_Ansprechpartner] (
  [AnsprechpartnerID] AUTOINCREMENT CONSTRAINT [PrimaryKey] PRIMARY KEY UNIQUE NOT NULL,
  [Mitgliedsnummer] VARCHAR (10) CONSTRAINT [MitgliederMitglied_Ansprechpartner] REFERENCES [Mitglieder] ([Mitgliedsnummer]) ON UPDATE CASCADE ,
  [GeschlechtID] LONG CONSTRAINT [GeschlechtMitglied_Ansprechpartner] REFERENCES [Geschlecht] ([GeschlechtID]) ON UPDATE CASCADE ,
  [Titel] VARCHAR (20),
  [Vorname] VARCHAR (50),
  [Nachname] VARCHAR (50),
  [StellungsID] LONG CONSTRAINT [Hierarchische StellungenMitglied_Ansprechpartner] REFERENCES [Hierarchische Stellungen] ([StellungsID]) ON UPDATE CASCADE ,
  [Telefon] VARCHAR (30),
  [Email] VARCHAR (100),
  [IstHauptansprechpartner] BIT,
  [GültigAb] DATETIME,
  [GültigBis] DATETIME,
  [ÄnderungsDatum] DATETIME,
  [ÄnderungsBenutzer] VARCHAR (50),
  [Bemerkung] LONGTEXT
)
