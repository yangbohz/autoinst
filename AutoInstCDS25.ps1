﻿#################################
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
# 安装文件文件夹名称
$installBase = Join-Path $shareBase "OpenLabCDS-2.5.0.927"
# .net安装文件存放位置（Win10的sxs下面的文件也在这里）
$netBase = Join-Path $shareBase "dotNet"
# AIC的安装属性文件名
$aicprop = Join-Path $installBase "aic.properties"
# 客户端的安装属性文件
$cltprop = Join-Path $installBase "clt.properties"
# 需要安装/更新的驱动，LabAdvisor也要放到这里面，注意要把MSI的安装文件放过来，不是散文件,安装时按按文件名称排序安装。
$drvbase = Join-Path $installBase "Drivers"

# 安装程序，现场准备程序，补丁包以及Adobe Reader的执行程序名
$cdsinstaller = Join-Path $installBase "Setup\CDSInstaller.exe"
$cdshf = Join-Path $installBase "OpenLAB_CDS_Update.exe"
# 英文版adobe及现场准备路径
if ([System.Globalization.Cultureinfo]::InstalledUICulture.LCID -eq "1033") {
    $adobe = (Get-ChildItem -Path (Join-Path $installBase "Setup\Tools\Adobe\Reader\*" ) -Recurse -Include *US* -ErrorAction SilentlyContinue).FullName
}
else {
    # 中文版adobe及现场准备路径
    $adobe = (Get-ChildItem -Path (Join-Path $installBase "Setup\Tools\Adobe\Reader\*" ) -Recurse -Include *CN* -ErrorAction SilentlyContinue).FullName
}

# 固定目录位置，不用管
$logbase = Join-Path $env:ProgramData "Agilent\InstallLogs"
$docpath = Join-Path $env:USERPROFILE "Documents"

# 检查必要的文件路径，如果安装程序或者属性文件不存在则退出执行
if (-not (Test-Path $cdsinstaller)) {
    Write-Warning -Message "Doesn't detected installer, expected path is $($cdsinstaller)"
    break
}
Write-Host ((Get-Date).ToString() + "  Get installer path $($cdsinstaller)") -ForegroundColor Cyan
if (-not (Test-Path $cltprop)) {
    Write-Warning -Message "Doesn't detected client config file, expected path is $($cltprop)"
    break
}
if (-not (Test-Path $aicprop )) {
    Write-Warning -Message "Doesn't detected AIC config file, expected path is $($aicprop)"
    break
}
if (-not (Test-Path $netBase)) {
    Write-Warning -Message "Doesn't detected .Net source file, expected path is $($netBase), Please confirm this is the expected behavior, it will continue after 10 seconds, if not, press CTRL+C to terminate the installation"
    Start-Sleep -Seconds 10
}
Write-Host  ((Get-Date).ToString() + "  All needed files present, installation start immediately") -ForegroundColor Cyan

# 安装NetFX
$OSVersion = (Get-Item "HKLM:SOFTWARE\Microsoft\Windows NT\CurrentVersion").GetValue('ReleaseID')
$netsxs = Join-Path $netBase $OSVersion
if ((Get-WindowsOptionalFeature -Online | Where-Object { $_.FeatureName -eq "WCF-NonHTTP-Activation" }).State -eq "disabled") {
    Write-Host ((Get-Date).ToString() + "  Enable netFx3.5") -ForegroundColor Green
    Enable-WindowsOptionalFeature -Online -FeatureName NetFx3, WCF-NonHTTP-Activation -All -LimitAccess -Source $netsxs | Out-Null
}

# 检测NETFX3是否启用成功
if ((Get-WindowsOptionalFeature -Online -FeatureName "NetFX3").State -eq "Disabled") {
    Write-Warning -Message "Detected .NetFX3 was not successfully activated, installation was terminated"
    Add-Content -Path (Join-Path $shareBase "Exception.log") -Value ($env:COMPUTERNAME + "`tNetFX3")
    break
}

if (-not (Test-Path -Path $logbase)) {
    New-Item -ItemType Directory -Path $logbase | Out-Null
}
Start-Process EXPLORER -ArgumentList $logbase

# 安装VC运行库
Write-Host ((Get-Date).ToString() + "  Start to install C++ runtime lib") -ForegroundColor Green
Start-Process -FilePath "$installBase\Setup\redist\vc_redist80sp1_x86.EXE" -ArgumentList "/q:a" -Wait
Start-Process -FilePath "$installBase\Setup\redist\vc_redist90sp1_x86.EXE" -ArgumentList "/q:a" -Wait
Start-Process -FilePath "$installBase\Setup\redist\vcredist_x86.exe" -ArgumentList "-install -quiet -norestart" -Wait
Start-Process -FilePath "$installBase\Setup\redist\vc_redist12.x64.exe" -ArgumentList "-install -quiet -norestart" -Wait
Start-Process -FilePath "$installBase\Setup\redist\vc_redist12.x86.EXE" -ArgumentList "-install -quiet -norestart" -Wait
Start-Process -FilePath "$installBase\Setup\redist\vc_redist15.x64.EXE" -ArgumentList "-install -quiet -norestart" -Wait
Start-Process -FilePath "$installBase\Setup\redist\vc_redist15.x86.EXE" -ArgumentList "-install -quiet -norestart" -Wait
# Start-Process -FilePath "$installBase\Setup\redist\VisualCppRedist_AIO_x86_x64.exe" -ArgumentList "/ai5839" -Wait

# 判断是否AIC
$aicNameList = Get-Content -Path (Join-Path (Get-ItemProperty $PSScriptRoot).FullName AICList.txt)

if ($aicNameList -notcontains "normal") {
    $isAIC = ($aicNameList -contains $env:COMPUTERNAME)
}
else {
    $isAIC = $env:COMPUTERNAME -match "AIC"
}

# 根据上一步判断结果安装工作站
if ($isAIC) {
    Write-Host ((Get-Date).ToString() + "  Start to install as AIC") -ForegroundColor Green
    Start-Process $cdsinstaller -ArgumentList "-s -c $aicprop" -Wait
    # 卸载Sample Scheduler插件，不卸载会在系统里报告大量错误
    # Start-Process msiexec -ArgumentList "/x {645F3E18-1ED9-458F-A8A9-2EF44104B074} /qn" -wait
}
else {
    Write-Host ((Get-Date).ToString() + "  Start to install as Client") -ForegroundColor Green
    Start-Process $cdsinstaller -ArgumentList "-s -c $cltprop" -Wait
}

# 检测安装状态
# 搜安装日志目录，只搜索20开头的目录，防止找到软件升级日志
$lastinstdir = Get-ChildItem $logbase -Filter 20* | Sort-Object -Property CreationTime -Descending | Select-Object -First 1
$lastinstlog = Join-Path $lastinstdir.FullName ("Agilent_OpenLab_CDS_" + $lastinstdir.BaseName + ".log")
$instStatus = Get-Content -Path $lastinstlog -Encoding utf8 -Tail 1
# 安装日志最后一行的后半部分内容。预期输出为类似下面的内容,如果检测不到安装成功的标志，写入异常日志。
# [2020-04-21T15:47:59]i007: Exit code: 0x0, restarting: No
if ($instStatus.Contains("0x0")) {
    Add-Content -Path (Join-Path $shareBase "Install_Summary.log") -Value ($env:COMPUTERNAME + "`t" + $instStatus.Substring(11))
}
else {
    Add-Content -Path (Join-Path $shareBase "Exception.log ") -Value ($env:COMPUTERNAME + "`tInstallation " + $instStatus.Substring(11))
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
        Add-Content -Path (Join-Path $shareBase "Exception.log ") -Value ($env:COMPUTERNAME + "`tUpdate " + $patchStatus.Substring(11))
    }
}

# 安装QualA
$qualabase = Join-Path $installBase "Setup\Tools\QualA"
$quala = (Get-ChildItem -Path $qualabase *.msi).FullName
$qualaplugin = (Get-ChildItem -Path (Join-Path $qualabase "CDS Plugin") *.msi).FullName

Write-Host ((Get-Date).ToString() + "  Start to install QualA") -ForegroundColor Green
if ([System.Globalization.Cultureinfo]::InstalledUICulture.LCID -eq "2052") {
    Start-Process MSIEXEC -ArgumentList "/qn /i `"$quala`" TRANSFORMS=`":zh-CN.mst`" " -Wait
    Start-Process MSIEXEC -ArgumentList "/qn /i `"$qualaplugin`" TRANSFORMS=`":zh-CN.mst`" " -Wait
}
else {
    
    Start-Process MSIEXEC -ArgumentList "/qn /i `"$quala`"" -Wait
    Start-Process MSIEXEC -ArgumentList "/qn /i `"$qualaplugin`"" -Wait
}

# 安装adobe阅读器
Write-Host ((Get-Date).ToString() + "  Start to install Adobe Reader") -ForegroundColor Green
Start-Process $adobe -ArgumentList "/sAll /rs EULA_ACCEPT=YES REMOVE_PREVIOUS=YES" -Wait

# 驱动
Get-ChildItem -Path $drvbase *.msi -Recurse | Sort-Object -Property Name | ForEach-Object {
    $msi = $_.FullName
    Write-Host ((Get-Date).ToString() + "  Start to install" + $_.BaseName) -ForegroundColor Green
    Start-Process MSIEXEC -ArgumentList "/qn /i `"$msi`"" -Wait
}

# # PalXT驱动
# if (Test-Path (Join-Path $drvbase palxt.exe)) {
#     Write-Host ((Get-Date).ToString() + "  Start to install Palxt driver") -ForegroundColor Green
#     Start-Process "$drvbase\palxt.exe" -ArgumentList "/S /v/qn" -Wait
# }

# 运行SVT工具，默认为监测到的产品生成简短报告
$svthome = (Get-ItemProperty -Path "HKLM:SOFTWARE\WOW6432Node\Agilent Technologies\IQTool").InstallLocation
ForEach ($reffile in (Get-ChildItem -path ($svthome + "\IQProducts") -Recurse *.xml)) {
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

# WIN10 1903服务修正
if ((Get-WindowsOptionalFeature -Online -FeatureName "WCF-TCP-Activation45").State -eq "Disabled") {
    Enable-WindowsOptionalFeature -Online -FeatureName "WCF-TCP-Activation45" -All -NoRestart | Out-Null
}
if ((Get-Service -Name "NetTcpPortSharing").StartType -ne "Automatic" ) {
    Set-Service -Name "NetTcpPortSharing" -StartupType Automatic
}

# 重启计算机
Write-Warning -Message ((Get-Date).ToString() + "  Installation completed, restart computer in 15 seconds")
Start-Sleep -Seconds 15
Restart-Computer -Force