SELECT
  Users.Id,
  UserRoles.Name AS UserRoleName,
  Users.Code,
  Users.Name,
  Users.Description,
  Users.Sort,
  Users.Active,
  Users.Memo
FROM
  Users
  LEFT JOIN UserRoles ON Users.UserRoleId = UserRoles.Id
WHERE
  (
    (
      (Users.Code) Like "*" & Forms!User_List!txtSearch & "*"
    )
  )
  Or (
    (
      (Users.Name) Like "*" & Forms!User_List!txtSearch & "*"
    )
  )
  Or (
    (
      (Users.Description) Like "*" & Forms!User_List!txtSearch & "*"
    )
  )
  Or (
    (
      (Users.Memo) Like "*" & Forms!User_List!txtSearch & "*"
    )
  )
ORDER BY
  Users.Name;
