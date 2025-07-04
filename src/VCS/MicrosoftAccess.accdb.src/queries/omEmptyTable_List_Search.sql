SELECT
  omEmptyTables.Id,
  omEmptyTables.Code,
  omEmptyTables.Name,
  omEmptyTables.Description,
  omEmptyTables.Sort,
  omEmptyTables.Active,
  omEmptyTables.Memo
FROM
  omEmptyTables
WHERE
  (
    (
      (omEmptyTables.Code) Like "*" & Forms!omEmptyTable_List!txtSearch & "*"
    )
  )
  Or (
    (
      (omEmptyTables.Name) Like "*" & Forms!omEmptyTable_List!txtSearch & "*"
    )
  )
  Or (
    (
      (omEmptyTables.Description) Like "*" & Forms!omEmptyTable_List!txtSearch & "*"
    )
  )
  Or (
    (
      (omEmptyTables.Memo) Like "*" & Forms!omEmptyTable_List!txtSearch & "*"
    )
  )
ORDER BY
  omEmptyTables.Name;
