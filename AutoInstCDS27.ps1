#################################
# 配置部分

# 是否生成啰嗦版SVT报告，默认不生成
$SVTAll = $false

# 安装waters驱动，默认是安装alliance，如果需要安装HClass，把下面的installHClass改成true,否则仅改这里即可
$installWaters = $false
# 安装Waters H-Class
$installHClass = $false

# 安装Thermo驱动
$installThermo = $false

# 配置结束
#################################


# 共享目录路径
# $shareBase = "\\olss\Agilent\"
$shareBase = (Get-ItemProperty $PSScriptRoot).Parent.FullName
# 安装文件文件夹名称
$installBase = Join-Path $shareBase "OpenLabCDS-2.7.0.787"
# .net安装文件存放位置（Win10的sxs下面的文件也在这里）
$netBase = Join-Path $installBase "dotNet"
# AIC的安装属性文件名
$aicprop = Join-Path $installBase "aic.properties"
# 客户端的安装属性文件
$cltprop = Join-Path $installBase "clt.properties"
# 需要安装/更新的驱动，LabAdvisor也要放到这里面，注意要把MSI的安装文件放过来，不是散文件,安装时按按文件名称排序安装。
$drvbase = Join-Path $installBase "Drivers"

# 安装程序，现场准备程序，补丁包以及Adobe Reader的执行程序名
$cdsinstaller = Join-Path $installBase "Setup\CDSInstaller.exe"
$cdshf = Join-Path $installBase "Update\OpenLAB_CDS_Update.exe"
# adobe,WATERS自动安装响应文件等与语言有关的东西,英文版
if ([System.Globalization.Cultureinfo]::InstalledUICulture.LCID -eq "1033") {
    $rspfile = Join-Path $drvbase "3P\Waters\Push Install\en\ICS_Response_EN_InstallAll_Agilent.rsp"
    # $adobe = (Get-ChildItem -Path (Join-Path $installBase "Setup\Tools\Adobe\Reader\*" ) -Recurse -Include *US* -ErrorAction SilentlyContinue).FullName
}
else {
    # 中文版
    $rspfile = Join-Path $drvbase "3P\Waters\Push Install\zh\ICS_Response_ZH_InstallAll_Agilent.rsp"
    # $adobe = (Get-ChildItem -Path (Join-Path $installBase "Setup\Tools\Adobe\Reader\*" ) -Recurse -Include *CN* -ErrorAction SilentlyContinue).FullName
}

# 固定目录位置，不用管
$logbase = Join-Path $env:ProgramData "Agilent\InstallLogs"

# 检查是否存在正在运行的安装程序，有时候意外重启会这样，防止冲突。
if ((Get-Process -Name CDSInstaller -ErrorAction SilentlyContinue).Count -ne 0) {
    Write-Warning -Message "Detect running CDSInstaller, maybe there was a force reboot in last installation."
    break
}

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
# 安装dotNetCore库
# Write-Host ((Get-Date).ToString() + "  Start to install dotNetCore runtime lib") -ForegroundColor Green
# Start-Process -FilePath "$installBase\Setup\redist\DotnetCore\windowsdesktop-runtime-3.1.10-win-x64.exe" -ArgumentList "/install /quiet /norestart" -Wait
# Start-Process -FilePath "$installBase\Setup\redist\DotnetCore\aspnetcore-runtime-3.1.10-win-x64.exe" -ArgumentList "/install /quiet /norestart" -Wait

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

# 判断是否AIC
$aicList = Join-Path (Get-ItemProperty $PSScriptRoot).FullName AICList.txt

if (Test-Path -Path $aicList) {
    $aicNameList = Get-Content -Path $aicList -Encoding utf8
    $isAIC = ($aicNameList -contains $env:COMPUTERNAME)
}
else {
    $isAIC = $env:COMPUTERNAME -match "AIC"
}

# 打开日志文件夹
Start-Process EXPLORER -ArgumentList $logbase

# 根据上一步判断结果安装工作站
if ($isAIC) {
    Write-Host ((Get-Date).ToString() + "  Start to install as AIC") -ForegroundColor Magenta
    Start-Process $cdsinstaller -ArgumentList "-s -config $aicprop" -Wait
    # 卸载Sample Scheduler插件，不卸载会在系统里报告大量错误
    # Start-Process msiexec -ArgumentList "/x {645F3E18-1ED9-458F-A8A9-2EF44104B074} /qn" -wait
}
else {
    Write-Host ((Get-Date).ToString() + "  Start to install as Client") -ForegroundColor Green
    Start-Process $cdsinstaller -ArgumentList "-s -config $cltprop" -Wait
}

# 检测安装状态
# 搜安装日志目录，只搜索20开头的目录，防止找到软件升级日志 2.7安装到半截会生成一个空日志文件夹，改为找倒数第一个正常的log文件
Get-ChildItem $logbase -Filter 20* | Sort-Object -Property CreationTime -Descending | ForEach-Object {
    $lastinstlog = Join-Path $_.FullName ("Agilent_OpenLab_CDS_" + $_.BaseName + ".log")
    if (Test-Path $lastinstlog) {
        continue
    }
} 

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
        Add-Content -Path (Join-Path $shareBase "Exception.log") -Value ($env:COMPUTERNAME + "`tUpdate " + $patchStatus.Substring(11))
    }
}

# 卸载备份/还原实用程序, 2.6Update05会在客户端/AIC上额外安装
# Write-Host ((Get-Date).ToString() + "  Start to uninstall B&R") -ForegroundColor Magenta
# Start-Process msiexec -ArgumentList "/x {A8327120-9557-4FB9-A38D-2703ED79B7C5} /qn" -Wait #备份
# Start-Process msiexec -ArgumentList "/x {01C23600-21B6-41B9-8CFD-6ED554CE268C} /qn" -Wait #还原

# 安装adobe阅读器 2.7不再集成安装包
# Write-Host ((Get-Date).ToString() + "  Start to install Adobe Reader") -ForegroundColor Green
# Start-Process $adobe -ArgumentList "/sAll /rs EULA_ACCEPT=YES REMOVE_PREVIOUS=YES" -Wait

# 标准msi驱动
Get-ChildItem -Path $drvbase -Include *.msi -Recurse | Where-Object { $_.FullName -notlike "*\AIC\*" -and $_.FullName -notlike "*\3P\*" } | Sort-Object -Property Name | ForEach-Object {
    $msi = $_.FullName
    Write-Host ((Get-Date).ToString() + "  Start to install " + $_.BaseName) -ForegroundColor Green
    Start-Process msiexec -ArgumentList "/qn /i `"$msi`" /norestart" -Wait
}

# AIC额外安装的msi包
if ($isAIC) {
    Get-ChildItem -Path (Join-Path $drvbase "AIC") -Include *.msi | Sort-Object -Property Name | ForEach-Object {
        $msi = $_.FullName
        Write-Host ((Get-Date).ToString() + "  Start to install " + $_.BaseName) -ForegroundColor Green
        Start-Process msiexec -ArgumentList "/qn /i `"$msi`" /norestart" -Wait
    }
}

# msp补丁
Get-ChildItem -Path $drvbase -Include *.msp -Recurse | Where-Object { $_.FullName -notlike "*AIC*" -and $_.FullName -notlike "*\3P\*" } | Sort-Object -Property Name | ForEach-Object {
    $msp = $_.FullName
    Write-Host ((Get-Date).ToString() + "  Start to install " + $_.BaseName) -ForegroundColor Green
    Start-Process msiexec -ArgumentList "/qn /p `"$msp`" /norestart" -Wait
}

# # PalXT驱动
# if (Test-Path (Join-Path $drvbase palxt.exe)) {
#     Write-Host ((Get-Date).ToString() + "  Start to install Palxt driver") -ForegroundColor Green
#     Start-Process "$drvbase\palxt.exe" -ArgumentList "/S /v/qn" -Wait
# }

# Waters驱动
if ($installWaters) {
    # 检查rsp文件内容是否正确，如不正确，进行修改
    [xml]$rsp = Get-Content $rspfile
    if ($rsp.Configuration.WORKING_DIRECTORY -ne "$drvbase\3P\Waters") {
        $rsp.Configuration.WORKING_DIRECTORY = "$drvbase\3P\Waters"
        $rsp.Configuration.LOG_FILE_NETWORK_LOCATION = "$drvbase\3P\Waters\Logs"
        # 检测为英文模式
        if ($rspfile -match "ICS_Response_EN_InstallAll_Agilent") {
            # 是否是HClass,如果是,安装完整版驱动包,否则仅安装alliance驱动包
            if ($installHClass) {
                $rsp.Configuration.ICS_LIST = "$drvbase\3P\Waters\Push Install\en\ICS_List_EN.txt"
            }
            else {
                $rsp.Configuration.ICS_LIST = "$drvbase\3P\Waters\Push Install\en\ICS_Agilent_Alliance_EN.txt"
            }
        }
        # 中文模式
        else {
            if ($installHClass) {
                $rsp.Configuration.ICS_LIST = "$drvbase\3P\Waters\Push Install\zh\ICS_List_ZH.txt"
            }
            else {
                $rsp.Configuration.ICS_LIST = "$drvbase\3P\Waters\Push Install\zh\ICS_Agilent_Alliance_ZH.txt"
            }
        }
        $rsp.Save($rspfile)
    }
    Write-Host ((Get-Date).ToString() + "  Start to install Waters Driver Pack") -ForegroundColor Green
    Start-Process "$drvbase\3P\Waters\Setup.exe" -ArgumentList "/responseFile `"$rspFile`"" -Wait
    Write-Host ((Get-Date).ToString() + "  Start to install Alliance Driver") -ForegroundColor Green
    Start-Process msiexec -ArgumentList "/qn /i `"$drvbase\3P\Waters.Alliance.Drivers.OLCDS2.Setup.msi`" /norestart" -Wait
    if ($installHClass) {
        Write-Host ((Get-Date).ToString() + "  Start to install H-Class Driver") -ForegroundColor Green
        Start-Process msiexec -ArgumentList "/qn /i `"$drvbase\3P\Agilent_OpenLabCDS_Waters_Acquity_Drivers.msi`" /norestart" -Wait
    }
}

# Thermo 驱动
if ($installThermo) {
    Write-Host ((Get-Date).ToString() + "  Start to install Thermo driver") -ForegroundColor Green
    # 赛默飞安装后会启动仪器服务，常规的wait参数会卡住，用另外的一种等待方式
    $thermoInstProc = Start-Process "$drvbase\3P\Thermo\Install.exe" -ArgumentList "/q /norestart"
    $thermoInstProc.WaitForExit()
    # 把变色龙仪器服务设为自动运行，否则第一次启动赛默飞仪器会要求管理员输入账号密码
    Set-Service -Name "ChromeleonRealTimeKernel" -StartupType Automatic
    Set-Service -Name "ChromeleonInstrumentService" -StartupType Automatic
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