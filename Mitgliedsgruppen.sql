CREATE TABLE [Mitgliedsgruppen] (
  [GruppenID] LONG CONSTRAINT [PrimaryKey] PRIMARY KEY UNIQUE NOT NULL,
  [Gruppennummer_Präfix] VARCHAR (2),
  [Gruppenname] VARCHAR (50),
  [Beschreibung] LONGTEXT
)
