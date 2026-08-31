CREATE TABLE [Mitglied_Adressen] (
  [AdressID] AUTOINCREMENT CONSTRAINT [PrimaryKey] PRIMARY KEY UNIQUE NOT NULL,
  [Mitgliedsnummer] VARCHAR (10) CONSTRAINT [MitgliederMitglied_Adressen] REFERENCES [Mitglieder] ([Mitgliedsnummer]) ON UPDATE CASCADE ,
  [AdressTypID] LONG CONSTRAINT [AdresstypenMitglied_Adressen] REFERENCES [Adresstypen] ([AdressTypID]) ON UPDATE CASCADE ,
  [VerwendungsID] LONG CONSTRAINT [AdressenverwendungenMitglied_Adressen] REFERENCES [Adressenverwendungen] ([VerwendungsID]) ON UPDATE CASCADE ,
  [LandID] LONG CONSTRAINT [LänderMitglied_Adressen] REFERENCES [Länder] ([LandID]) ON UPDATE CASCADE ,
  [Straße] VARCHAR (50),
  [Hausnummer] VARCHAR (10),
  [Postfachnummer] VARCHAR (20),
  [Postleitzahl] VARCHAR (5),
  [Ortsname] VARCHAR (50),
  [Bundesland_Region] VARCHAR (50),
  [GültigAb] DATETIME,
  [GültigBis] DATETIME,
  [ÄnderungsDatum] DATETIME,
  [ÄnderungsBenutzer] VARCHAR (50),
  [Bemerkung] LONGTEXT
)
