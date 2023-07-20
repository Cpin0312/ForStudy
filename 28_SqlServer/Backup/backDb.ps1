Invoke-Sqlcmd -Query "exec sp_databases" | ?{ -not (@("master","tempdb","msdb","model") -contains $_.Item(0)) } | %{
    $datebaseName = $_.Item(0)
    
    Write-Host ("■ {0}のフルバックアップを実行します" -f $datebaseName)
    
    $backupSql=@"
BACKUP DATABASE {0}
   TO DISK = 'C:\Users\hisol\Desktop\sqlsvBK\Backup\{0}.Bak'
   WITH FORMAT,
   MEDIANAME = 'SQLServerBackups',
   NAME = 'Full Backup of {0}'
go
exit
"@ -f $datebaseName

    sqlcmd -q $backupSql
    
    Write-Host ""
} 
