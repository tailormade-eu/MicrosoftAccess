CREATE TABLE [omEmptyTables] (
  [Id] AUTOINCREMENT CONSTRAINT [PrimaryKey] PRIMARY KEY UNIQUE NOT NULL,
  [Code] VARCHAR (255),
  [Name] VARCHAR (255),
  [Description] VARCHAR (255),
  [Memo] LONGTEXT,
  [Sort] DOUBLE,
  [Active] BIT,
  [CreateDate] DATETIME,
  [CreateUserId] LONG,
  [CreateUserName] VARCHAR (255),
  [ModifyDate] DATETIME,
  [ModifyUserId] LONG,
  [ModifyUserName] VARCHAR (255)
)
