CREATE TABLE [tbl_Mitgliedstypen] (
  [Mitgliedstyp] VARCHAR (50) CONSTRAINT [PrimaryKey] PRIMARY KEY UNIQUE NOT NULL,
  [Beschreibung] VARCHAR (200),
  [RequiresVorname] BIT,
  [IsAktiv] BIT
)
