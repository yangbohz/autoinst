#################################
# 配置部分

# 是否生成啰嗦版SVT报告，默认不生成
$SVTAll = $false

# 配置结束
#################################


# 获取配置文件
# $conf = (Get-Content -Path ($PSScriptRoot + "\conf.json") -Encoding UTF8 | ConvertFrom-Json)
# 自用调试
# $conf = (Get-Content -Path ./conf.json -Encoding UTF8 | ConvertFrom-Json)
# 共享目录路径
# $shareBase = "\\" + $conf.ServerName + "\Agilent\"
$shareBase = (Get-ItemProperty $PSScriptRoot).Parent.FullName
# 更新文件文件夹名称
$updateBase = Join-Path $shareBase "update"

# 需要安装/更新的驱动，LabAdvisor也要放到这里面，注意要把MSI的安装文件放过来，不是散文件,安装时按按文件名称排序安装。
$drvbase = Join-Path $updateBase "Drivers"

# 补丁包执行程序名
$cdshf = Join-Path $updateBase "OpenLAB_CDS_Update.exe"


# 固定目录位置，不用管
$logbase = Join-Path $env:ProgramData "Agilent\InstallLogs"
$docpath = Join-Path $env:USERPROFILE "Documents"

# 检查必要的文件路径，如果安装程序或者属性文件不存在则退出执行
Write-Host ((Get-Date).ToString() + "  Get update path $($cdshf)") -ForegroundColor Cyan

if (-not (Test-Path -Path $logbase)) {
    New-Item -ItemType Directory -Path $logbase | Out-Null
}

# 如果补丁文件存在，则安装
if (Test-Path $cdshf) {
    Write-Host ((Get-Date).ToString() + "  Start to install CDS Update") -ForegroundColor Green
    Start-Process $cdshf -ArgumentList "-s LicenseAccepted=True -norestart" -Wait
    # 检测补丁安装状态
    $lastpatchdir = Get-ChildItem (Join-Path $logbase SoftwareUpdate) | Sort-Object -Property CreationTime -Descending | Select-Object -First 1
    $lastpatchlog = Get-ChildItem -Path $lastpatchdir.FullName -Filter ("*" + $lastpatchdir.BaseName + ".log")
    $patchStatus = Get-Content -Path $lastpatchlog.FullName -Encoding utf8 -Tail 1
    if ($patchStatus.Contains("0x0")) {
        Add-Content -Path (Join-Path $shareBase "Update_Summary.log") -Value ($env:COMPUTERNAME + "`t" + $patchStatus.Substring(11))
    }
    else {
        Add-Content -Path (Join-Path $shareBase "Exception.log") -Value ($env:COMPUTERNAME + "`tUpdate " + $patchStatus.Substring(11))
    }
}
else {
    Write-Host ((Get-Date).ToString() + "  Update not present. Install driver directly") -ForegroundColor Magenta 
}
# 驱动
Get-ChildItem -Path $drvbase *.msi -Recurse | Sort-Object -Property Name | ForEach-Object {
    $msi = $_.FullName
    Write-Host ((Get-Date).ToString() + "  Start to install " + $_.BaseName) -ForegroundColor Green
    Start-Process MSIEXEC -ArgumentList "/qn /i `"$msi`"" -Wait
}

# # PalXT驱动
# if (Test-Path (Join-Path $drvbase palxt.exe)) {
#     Write-Host ((Get-Date).ToString() + "  Start to install Palxt driver") -ForegroundColor Green
#     Start-Process "$drvbase\palxt.exe" -ArgumentList "/S /v/qn" -Wait
# }

# 运行SVT工具，默认为监测到的产品生成简短报告
$svthome = (Get-ItemProperty -Path "HKLM:SOFTWARE\WOW6432Node\Agilent Technologies\IQTool").InstallLocation
ForEach ($reffile in (Get-ChildItem -Path ($svthome + "\IQProducts") -Recurse *.xml)) {
    [xml] $isbase = Get-Content $reffile.FullName -Encoding UTF8
    if ($null -ne $isbase.PRODUCT.BASE) {
        if ($null -ne $isbase.PRODUCT.PRODUCTINFO.PRODUCTNAME) {
            $proname = $proname + $isbase.PRODUCT.PRODUCTINFO.PRODUCTNAME + ","
        }
        else {
            $proname = $proname + $isbase.PRODUCT.NAME.'#text' + ","
        }
    }
}
if ($SVTAll) {
    Start-Process ($svthome + "\Bin\SFVTool.exe") -ArgumentList "-qt -slient -showall -p:`"$proname`" -pdf -xml"  -Wait
}
else {
    Start-Process ($svthome + "\Bin\SFVTool.exe") -ArgumentList "-qt -slient -shownothing -p:`"$proname`" -pdf -xml" -Wait
}

# 将SVT工具生成的PDF复制到我的文档中
$sv = 'C:\SVReports'
# 在桌面上生成IQOQ文件夹并将svreport复制过去
# $dest = [Environment]::GetFolderPath("Desktop") + '\' + $iqoq + '\'
# New-Item -Path $dest -ItemType Directory -ErrorAction SilentlyContinue
# 在将svreport直接复制到“我的文档”中，ACE默认打开我的文档目录，放到这里会稍微省点麻烦。
$dest = $docpath
$rptnum = Get-ChildItem $sv | Measure-Object
Get-ChildItem $sv *.pdf -Recurse | Sort-Object -Property CreationTime -Descending | Select-Object -First $rptnum.Count |
ForEach-Object {
    $dirname = (Get-ItemProperty $_.FullName).Directory.Name
    $crttm = (Get-ItemProperty $_.FullName).CreationTime.ToString("yyyyMMdd-HHmm") 
    $newname = 'SVReport_{0}_{1}.pdf' -f $dirname, $crttm
    Rename-Item $_.fullname -NewName $newname
    $path = $_.directoryname + '\' + $newname
    Copy-Item $path $dest
}

# 检测生成的svt报告是否均为pass，生成汇总报告
Get-ChildItem $sv *.xml -Recurse | Sort-Object -Property CreationTime -Descending | Select-Object -First $rptnum.Count |
ForEach-Object {
    [xml] $svstate = Get-Content $_.FullName
    Add-Content -Path (Join-Path $shareBase "SVT_Summary.log") -Value ($svstate.QualifiedSuiteResultCollection.QualifiedSuiteResult.Status + "`t" + $env:COMPUTERNAME + "`t" + $svstate.QualifiedSuiteResultCollection.QualifiedSuiteResult.Name )
}

# 重启计算机
Write-Warning -Message ((Get-Date).ToString() + "  Installation completed, restart computer in 15 seconds")
Start-Sleep -Seconds 15
Restart-Computer -Force