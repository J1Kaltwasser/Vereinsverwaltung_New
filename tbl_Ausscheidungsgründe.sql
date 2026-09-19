CREATE TABLE [tbl_Ausscheidungsgründe] (
  [Ausscheidungsgrund] VARCHAR (100) CONSTRAINT [PrimaryKey] PRIMARY KEY UNIQUE NOT NULL,
  [Kategorie] VARCHAR (50),
  [Beschreibung] LONGTEXT,
  [IsAktiv] BIT
)
