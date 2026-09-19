CREATE TABLE [tbl_Mitgliedsgruppen] (
  [Mitgliedsgruppe] VARCHAR (100) CONSTRAINT [PrimaryKey] PRIMARY KEY UNIQUE NOT NULL,
  [Beschreibung] LONGTEXT,
  [Beitrag] CURRENCY,
  [IsAktiv] BIT,
  [SortOrder] SHORT
)
