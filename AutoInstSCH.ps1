#################################
# 配置部分

# 是否生成啰嗦版SVT报告，默认不生成
$SVTAll = $false

# 安装语言
$lang = "zh-Hans"

# ECMXT计算机名
$OLSS = "ECMXT"

# Automation上传缓存位置
$SSCache = "D:\AgilentCache\SecureFileSystem"

# Scheduler 网页端口
$AUUPort = "80"


# 配置结束
#################################


# 共享目录路径
# $shareBase = "\\olss\Agilent\"
$shareBase = (Get-ItemProperty $PSScriptRoot).Parent.FullName
# 安装文件文件夹名称
$installBase = Join-Path $shareBase "OpenLabCDS-2.7.0.787"
# .net安装文件存放位置（Win10的sxs下面的文件也在这里）
$netBase = Join-Path $installBase "dotNet"

# 安装程序，现场准备程序，补丁包以及Adobe Reader的执行程序名
$cminstaller = Join-Path $installBase "Setup\DatastoreClient.exe"
$schinstaller = Join-Path $installBase "Setup\AutoUploaderSetup.exe"
$cmupdate = Join-Path $installBase "CMUpdate\OpenLAB_CmClient_Update.exe"

# 固定目录位置，不用管
$logbase = Join-Path $env:ProgramData "Agilent\InstallLogs"

# 检查是否存在正在运行的安装程序，有时候意外重启会这样，防止冲突。
if ((Get-Process -Name DatastoreClient -ErrorAction SilentlyContinue).Count -ne 0) {
    Write-Warning -Message "Detect running DatastoreClient, maybe there was a force reboot in last installation."
    break
}

# 检查必要的文件路径，如果安装程序或者属性文件不存在则退出执行
if (-not (Test-Path $cminstaller)) {
    Write-Warning -Message "Doesn't detected installer, expected path is $($cminstaller)"
    break
}
if (-not (Test-Path $schinstaller)) {
    Write-Warning -Message "Doesn't detected installer, expected path is $($schinstaller)"
    break
}


if (-not (Test-Path $netBase)) {
    Write-Warning -Message "Doesn't detected .Net source file, expected path is $($netBase), Please confirm this is the expected behavior, it will continue after 10 seconds, if not, press CTRL+C to terminate the installation"
    Start-Sleep -Seconds 10
}
Write-Host ((Get-Date).ToString() + "  All needed files present, installation start immediately") -ForegroundColor Cyan

# 安装NetFX, 使用Windows的数字Build ID，例如19044（对应于win10 21H2）
$OSVersion = [System.Environment]::OSVersion.Version.Build

Write-Host ((Get-Date).ToString() + "  Detect OS Build is $($OSVersion)") -ForegroundColor Cyan

if ((Get-WindowsOptionalFeature -Online -FeatureName "WCF-NonHTTP-Activation" ).State -eq "disabled") {
    Get-ChildItem $netBase -Directory | ForEach-Object {
        if ($_.BaseName -match $OSVersion) {
            Write-Host ((Get-Date).ToString() + "  Enable netFx3.5, use sxs folder is $($_.BaseName)") -ForegroundColor Green
            try {
                Enable-WindowsOptionalFeature -Online -FeatureName NetFx3, WCF-NonHTTP-Activation -All -LimitAccess -Source $_.FullName | Out-Null
            }
            catch {
                Write-Warning -Message "Not correct version SxS source, use next source"
            }
        }
    }
}

# 检测NETFX3是否启用成功
if ((Get-WindowsOptionalFeature -Online -FeatureName "NetFX3").State -ne "Enabled") {
    Write-Warning -Message "Detected .NetFX3 was not successfully activated, installation was terminated"
    Add-Content -Path (Join-Path $shareBase "Exception.log") -Value ($env:COMPUTERNAME + "`tNetFX3`t" + $OSVersion)
    break
}

if (-not (Test-Path -Path $logbase)) {
    New-Item -ItemType Directory -Path $logbase | Out-Null
}

# 检查是否应用了预置组策略，如果没有检测到自制安捷伦壁纸，则认为没有应用组策略，将使用SPT进行系统设定
$spt = Join-Path $installBase "Setup\Tools\SPT\SystemPreparationTool.exe"
if (-not (Test-Path -Path "C:\Windows\Agilent.png")) {
    Write-Host ((Get-Date).ToString() + "  +_+ Seemed Agilent Group Policy not be applied. SPT will run full configure") -ForegroundColor Yellow
    Start-Process -FilePath $spt -ArgumentList "-silent -norestart ConditionRecommended=True ConfigurationName=`"IES Customerzed for CDS 2.7`""
}
else {
    Write-Host ((Get-Date).ToString() + "  ^_^ Agilent Group Policy detected. Run SPT with lite configure") -ForegroundColor Yellow
    Start-Process -FilePath $spt -ArgumentList "-silent -norestart ConditionRecommended=True ConfigurationName=`"IES Customerzed for CDS 2.7 LITE`""
 
}

# 安装VC运行库
Write-Host ((Get-Date).ToString() + "  Start to install C++ runtime lib") -ForegroundColor Green
Start-Process -FilePath "$installBase\Setup\redist\vc_redist80sp1_x86.EXE" -ArgumentList "/q:a" -Wait
Start-Process -FilePath "$installBase\Setup\redist\vc_redist90sp1_x86.EXE" -ArgumentList "/q:a" -Wait
Start-Process -FilePath "$installBase\Setup\redist\vcredist_x86.exe" -ArgumentList "-install -quiet -norestart" -Wait
Start-Process -FilePath "$installBase\Setup\redist\vc_redist12.x64.exe" -ArgumentList "-install -quiet -norestart" -Wait
Start-Process -FilePath "$installBase\Setup\redist\vc_redist12.x86.EXE" -ArgumentList "-install -quiet -norestart" -Wait
Start-Process -FilePath "$installBase\Setup\redist\vc_redist16.x64.EXE" -ArgumentList "-install -quiet -norestart" -Wait
Start-Process -FilePath "$installBase\Setup\redist\vc_redist16.x86.EXE" -ArgumentList "-install -quiet -norestart" -Wait
# https://github.com/abbodi1406/vcredist AllInOne C++运行库,安装的似乎有点乱，先不用。
# Start-Process -FilePath "$installBase\Setup\redist\VisualCppRedist_AIO_x86_x64.exe" -ArgumentList "/ai58239" -Wait

# 安装dotNetCore库
Write-Host ((Get-Date).ToString() + "  Start to install dotNetCore runtime lib") -ForegroundColor Green
Get-ChildItem -Path $netBase\core -Recurse -Include *.exe | ForEach-Object {
    Write-Host ((Get-Date).ToString() + "  Start to install " + $_.BaseName) -ForegroundColor Green
    Start-Process -FilePath $_.FullName -ArgumentList "/install /quiet /norestart" -Wait
}


# 打开日志文件夹
Start-Process EXPLORER -ArgumentList $logbase

# 根据上一步判断结果安装工作站

    Write-Host ((Get-Date).ToString() + "  Start to install CM Client") -ForegroundColor Green
    Start-Process $cminstaller -ArgumentList " -LicenseAccepted=True -AGILENTHOME=`"C:\Program Files (x86)\Agilent Technologies`" -LanguageCode=$($lang) -OlssHostName=$($OLSS) -OLSS_FILE_STORAGE=$($SSCache) -s" -Wait

    
    Write-Host ((Get-Date).ToString() + "  Start to install Scheduler") -ForegroundColor Green
    Start-Process $schinstaller -ArgumentList "-LicenseAccepted=True -AGILENTHOME=`"C:\Program Files (x86)\Agilent Technologies`" -LanguageCode=$($lang) -AUU_PORT=$($AUUPort) -s" -Wait

# 如果补丁文件存在，则安装
if (Test-Path $cmupdate) {
    Write-Host ((Get-Date).ToString() + "  Start to install CM Client Update") -ForegroundColor Green
    Start-Process $cmupdate -ArgumentList "-s LicenseAccepted=True -norestart" -Wait
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

# 运行SVT工具，默认为监测到的产品生成简短报告
$svthome = (Get-ItemProperty -Path "HKLM:SOFTWARE\WOW6432Node\Agilent Technologies\IQTool").InstallLocation
# ForEach ($reffile in (Get-ChildItem -Path ($svthome + "\IQProducts") -Recurse *.xml)) {
#     [xml] $isbase = Get-Content $reffile.FullName -Encoding UTF8
#     if ($null -ne $isbase.PRODUCT.BASE) {
#         if ($null -ne $isbase.PRODUCT.PRODUCTINFO.PRODUCTNAME) {
#             $proname = $proname + $isbase.PRODUCT.PRODUCTINFO.PRODUCTNAME + ","
#         }
#         else {
#             $proname = $proname + $isbase.PRODUCT.NAME.'#text' + ","
#         }
#     }
# }
if ($SVTAll) {
    Start-Process ($svthome + "\Bin\SFVTool.exe") -ArgumentList "-qt -slient -showall -p:`"all`" -pdf -xml"  -Wait
}
else {
    Start-Process ($svthome + "\Bin\SFVTool.exe") -ArgumentList "-qt -slient -shownothing -p:`"all`" -pdf -xml" -Wait
}

# 将SVT工具生成的PDF复制到我的文档中
$sv = 'C:\SVReports'
# 在桌面上生成IQOQ文件夹并将svreport复制过去
$dest = Join-Path ([Environment]::GetFolderPath("Desktop")) IQOQ
New-Item -Path $dest -ItemType Directory -ErrorAction SilentlyContinue
# 在将svreport直接复制到“我的文档”中，ACE默认打开我的文档目录，放到这里会稍微省点麻烦。ACE3改了工作机制，不再使用文档作为默认位置，使用上面的方式。
# $docpath = Join-Path $env:USERPROFILE "Documents"
# $dest = $docpath
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