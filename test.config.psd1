@{
    ExecutionMode = 'FailFast'
    CurrentRunWeeks = '2'
    EnforceBackupValidation = $false
    BackupValidationRules = @{
        CommonFixedFiles = @(
            'COD WeCom Login to Non-Approved Devices FID BU - Report({startDate} - {endDate}).msg'
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
            'MSMS WeCom Mail Log_{startDate}_{endDate}.csv'
            "中国's member operation records{endDatePlus1MMdd}.xlsx"
            "摩根士丹利国际银行（中国）'s member operation records{endDatePlus1MMdd}.xlsx"
        )
        FourWeekFixedFiles = @(
            'Conduct WeCom Log Audit file uploaded.msg'
            'msbic-miniapp.png'
            'msms-miniapp.png'
            'MSMS WeCom Mail Log_{startDate}_{endDate}.csv'
            "中国's member operation records{endDatePlus1MMdd}.xlsx"
            "中国's member test operation records{endDatePlus1MMdd}.xlsx"
        )
    }
    Tasks = @(
        @{
            Name = 'mail-msms'
            Type = 'mail'
            BU = 'MSMS'
            Enabled = $true
            InputDirectory = 'C:\addin_deploy_cert\wecom_audit_log'
            FileNamePattern = 'MSMS WeCom Mail Log_{startDate}_{endDate}.csv'
        }
        @{
            Name = 'device-msms'
            Type = 'device'
            BU = 'MSMS'
            Enabled = $false
            InputDirectory = 'C:\addin_deploy_cert\incoming'
            FileNamePattern = 'msms_device_log.xlsx'
        }
        @{
            Name = 'device-msbic'
            Type = 'device'
            BU = 'MSBIC'
            Enabled = $false
            InputDirectory = 'C:\addin_deploy_cert'
            FileNamePattern = 'test_msbic_records1028.xlsx'
        }
        @{
            Name = 'device-msimc'
            Type = 'device'
            BU = 'MSIMC'
            Enabled = $false
            InputDirectory = 'C:\addin_deploy_cert\incoming'
            FileNamePattern = 'msimc_device_log.xlsx'
        }
        @{
            Name = 'mail-msbic'
            Type = 'mail'
            BU = 'MSBIC'
            Enabled = $false
            InputDirectory = 'C:\addin_deploy_cert\incoming'
            FileNamePattern = 'MSBIC WeCom Mail Log _{startDate}_{endDate}.csv'
        }
        @{
            Name = 'mail-msimc'
            Type = 'mail'
            BU = 'MSIMC'
            Enabled = $false
            InputDirectory = 'C:\addin_deploy_cert\wecom_audit_log'
            FileNamePattern = 'MSIMC WeCom Mail Log _{startDate}_{endDate}.csv'
        }
        @{
            Name = 'device-msms'
            Type = 'device'
            BU = 'MSMS'
            Enabled = $true
            InputDirectory = 'C:\addin_deploy_cert\wecom_audit_log'
            FileNamePattern = "中国's member operation records{endDatePlus1MMdd}.xlsx"
        }
    )
}
