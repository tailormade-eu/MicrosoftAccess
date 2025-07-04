CREATE TABLE [UserRoles] (
  [Id] AUTOINCREMENT,
  [Code] VARCHAR (255),
  [Name] VARCHAR (255),
  [Description] VARCHAR (255),
  [Active] BIT,
  [Sort] DOUBLE,
  [Memo] LONGTEXT,
  [CreateDate] DATETIME,
  [CreateUserId] LONG,
  [ModifyDate] DATETIME,
  [ModifyUserId] LONG
)
