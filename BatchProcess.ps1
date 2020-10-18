# 管理员账号
$admin = "ADMIN"
# 管理员密码
$pass = "ZX..333"
# 安装脚本位置
$cmd = "\\ecmxt\Agilent\AUTOINST\AUTO_INST_CDS24_SHARE.ps1"
$user = $env:USERDOMAIN + "\" + $admin
$PCs = Get-Content ./PCs.txt
foreach ($pc in $PCs) {
    # & cmdkey /generic $pc /user:$user /pass:$pass
    if (Test-Connection -ComputerName $pc -Quiet -Count 2) {
        # & mstsc /v $pc
        # Restart-Computer -ComputerName $pc -Force
        &.\PsExec.exe \\$pc -u $user -p $pass -h -d -accepteula powershell $cmd 2>&1 | ForEach-Object { "$_" }
    }
    else {
        Write-Warning -Message "Timeout when ping $pc"
    }
}