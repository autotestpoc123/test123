@{
    ScheduleAnchor = '20260402'
    ExecutionMode = 'FailFast'
    CurrentRunWeeks = '2'
    EnforceBackupValidation = $false
    InputRoot = 'C:\addin_deploy_cert'
    LogRoot = 'C:\SysAdmin\log'
    BackupRoot = '\\cod.test.com.cn\apptest\wecom_audit_log_backup'
    SourceCleanup = @{
        Enabled      = $false
        AllowedRoots = @(
            '\\cod.test.com.cn\apptest\wecom_audit_log'
        )
    }
    BackupValidationRules = @{
        CommonFixedFiles = @(
            @{ File = 'COD WeCom Login to Non-Approved Devices FID BU - Report({startDate} - {endDate}).msg'; ReadyBy = 'Validate' }
        )
        DynamicFiles = @(
            @{
                SummaryTaskName = 'device-msms'
                BaseName = 'COD WeCom Login to Non-Approved Devices IM BU - Report({startDate} - {endDate}).msg'
            }
            @{
                SummaryTaskName = 'mail-msms'
                BaseName = 'COD WeCom Mail Data Leakage Manual Review - from {startDate} to {endDate}.msg'
            }
        )
        TwoWeekFixedFiles = @(
            @{ File = 'MSMS WeCom Mail Log_{startDate}_{endDate}.csv'; ReadyBy = 'Analysis' }
            @{ File = "中国's member operation records{endDatePlus1MMdd}.xlsx"; ReadyBy = 'Analysis' }
            @{ File = "摩根士丹利国际银行（中国）'s member operation records{endDatePlus1MMdd}.xlsx"; ReadyBy = 'Analysis' }
        )
        FourWeekFixedFiles = @(
            @{ File = 'Conduct WeCom Log Audit file uploaded.msg'; ReadyBy = 'Validate' }
            @{ File = 'msbic-miniapp.png'; ReadyBy = 'Analysis' }
            @{ File = 'msms-miniapp.png'; ReadyBy = 'Analysis' }
            @{ File = 'MSMS WeCom Mail Log_{startDate}_{endDate}.csv'; ReadyBy = 'Analysis' }
            @{ File = "中国's member operation records{endDatePlus1MMdd}.xlsx"; ReadyBy = 'Analysis' }
            @{ File = "中国's member test operation records{endDatePlus1MMdd}.xlsx"; ReadyBy = 'Analysis' }
        )
    }
    Notification = @{
        PROD = @{
            SmtpServer   = 'mailhost.ms.com'
            Port         = 2587
            From         = 'wecom-audit-prod'
            CertName     = 'wecom-audit-prod-cert'
            OpsTeam      = @('ops-team@corp.com')
            CcRecipients = @('admin@corp.com')
        }
        QA = @{
            SmtpServer   = 'mailhost.ms.com'
            Port         = 2587
            From         = 'wecom-audit-qa'
            CertName     = 'wecom-audit-qa-cert'
            OpsTeam      = @('ling.gu@infradev.mocktest.com.cn')
            CcRecipients = @('ling.gu@infradev.mocktest.com.cn')
        }
    }
    Tasks = @(
        @{
            Name = 'mail-msms'
            Type = 'mail'
            BU = 'MSMS'
            Enabled = $true
            InputDirectory = '{InputRoot}\wecom_audit_log'
            FileNamePattern = 'MSMS WeCom Mail Log_{startDate}_{endDate}.csv'
        }
        @{
            Name = 'device-msms'
            Type = 'device'
            BU = 'MSMS'
            Enabled = $false
            InputDirectory = '{InputRoot}\incoming'
            FileNamePattern = 'msms_device_log.xlsx'
        }
        @{
            Name = 'device-msbic'
            Type = 'device'
            BU = 'MSBIC'
            Enabled = $false
            InputDirectory = '{InputRoot}'
            FileNamePattern = 'test_msbic_records1028.xlsx'
        }
        @{
            Name = 'device-msimc'
            Type = 'device'
            BU = 'MSIMC'
            Enabled = $false
            InputDirectory = '{InputRoot}\incoming'
            FileNamePattern = 'msimc_device_log.xlsx'
        }
        @{
            Name = 'mail-msbic'
            Type = 'mail'
            BU = 'MSBIC'
            Enabled = $false
            InputDirectory = '{InputRoot}\incoming'
            FileNamePattern = 'MSBIC WeCom Mail Log _{startDate}_{endDate}.csv'
        }
        @{
            Name = 'mail-msimc'
            Type = 'mail'
            BU = 'MSIMC'
            Enabled = $false
            InputDirectory = '{InputRoot}\wecom_audit_log'
            FileNamePattern = 'MSIMC WeCom Mail Log _{startDate}_{endDate}.csv'
        }
        @{
            Name = 'device-msms-member-records'
            Type = 'device'
            BU = 'MSMS'
            Enabled = $true
            InputDirectory = '{InputRoot}\wecom_audit_log'
            FileNamePattern = "中国's member operation records{endDatePlus1MMdd}.xlsx"
        }
    )
}
