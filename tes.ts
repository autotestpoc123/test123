$deploy = 'D:\实际脚本目录\Deploy-PhotoImportScheduledTask.ps1'

$parameters = @{
    Environment      = 'PROD'
    SystemId         = 'DOMAIN\实际systemid'
    EnableSchedule   = $false
    AllowWriteConfig = $true
}

# 预检
& $deploy @parameters -WhatIf

# 首次注册
& $deploy @parameters
