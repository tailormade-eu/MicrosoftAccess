CREATE TABLE [Users] (
  [Id] AUTOINCREMENT,
  [UserRoleId] LONG,
  [Code] VARCHAR (255),
  [Name] VARCHAR (255),
  [Description] VARCHAR (255),
  [Login] VARCHAR (255),
  [Password] VARCHAR (255),
  [Email] VARCHAR (255),
  [WorkHoursPerDay] DOUBLE,
  [Active] BIT,
  [Sort] DOUBLE,
  [Memo] LONGTEXT,
  [CreateDate] DATETIME,
  [CreateUserId] LONG,
  [ModifyDate] DATETIME,
  [ModifyUserId] LONG
)
