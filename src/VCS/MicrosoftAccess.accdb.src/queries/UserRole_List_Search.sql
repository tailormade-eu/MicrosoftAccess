SELECT
  UserRoles.Id,
  UserRoles.Code,
  UserRoles.Name,
  UserRoles.Description,
  UserRoles.Sort,
  UserRoles.Active,
  UserRoles.Memo
FROM
  UserRoles
WHERE
  (
    (
      (UserRoles.Code) Like "*" & Forms!UserRole_List!txtSearch & "*"
    )
  )
  Or (
    (
      (UserRoles.Name) Like "*" & Forms!UserRole_List!txtSearch & "*"
    )
  )
  Or (
    (
      (UserRoles.Description) Like "*" & Forms!UserRole_List!txtSearch & "*"
    )
  )
  Or (
    (
      (UserRoles.Memo) Like "*" & Forms!UserRole_List!txtSearch & "*"
    )
  )
ORDER BY
  UserRoles.Name;
