CREATE TABLE [tbl_Geschlecht] (
  [Geschlecht] VARCHAR (20) CONSTRAINT [PrimaryKey] PRIMARY KEY UNIQUE NOT NULL,
  [Anrede] VARCHAR (20),
  [IsAktiv] BIT
)
