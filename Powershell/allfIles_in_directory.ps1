$DirPath = '\\quintiles.net\Enterprise\Organization\Biostatistics\Business Solutions\Solution Providers\General\GENIE\_ErrorReporCollation'
$Outpath = '\\quintiles.net\Enterprise\Organization\Biostatistics\Business Solutions\Solution Providers\General\GENIE\_ErrorReporCollation\Services.txt'

Get-ChildItem -Path $DirPath -Recurse | Out-File -FilePath $Outpath

#Get-ChildItem -Path '\\quintiles.net\Enterprise\Organization\Biostatistics\Business Solutions\Solution Providers\General\GENIE\_ErrorReporCollation' -Recurse | Out-File -FilePath '\\quintiles.net\Enterprise\Organization\Biostatistics\Business Solutions\Solution Providers\General\GENIE\_ErrorReporCollation\Services.txt'

