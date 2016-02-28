CREATE TABLE [ErrorLog] (
  [ErrorCode] LONG ,
  [ErrorText] LONGTEXT ,
  [UserName] VARCHAR (50),
  [CurrentForm] VARCHAR (255),
  [CurrentControl] VARCHAR (255),
  [ActiveForms] SHORT ,
  [Date] DATETIME ,
  [CallingProcedure] VARCHAR (255)
)
