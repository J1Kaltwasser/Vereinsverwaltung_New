CREATE TABLE [tbl_ErrorLog] (
  [ErrorLogID] AUTOINCREMENT CONSTRAINT [PrimaryKey] PRIMARY KEY UNIQUE NOT NULL,
  [ErrorNumber] LONG,
  [ErrorCode] LONG,
  [ErrorCategory] BYTE,
  [ErrorSeverity] BYTE,
  [ErrorDescription] LONGTEXT,
  [ErrorSource] VARCHAR (200),
  [LineNumber] LONG,
  [StackTrace] LONGTEXT,
  [TimeStamp] DATETIME,
  [UserName] VARCHAR (50),
  [DatabasePath] VARCHAR (255),
  [Context] LONGTEXT,
  [AffectedRecords] LONGTEXT,
  [AktionTaken] LONGTEXT,
  [IsResolved] BIT,
  [ResolutionNotes] LONGTEXT
)
