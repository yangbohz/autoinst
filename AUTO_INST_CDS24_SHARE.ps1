# 获取配置文件
$conf = (Get-Content -Path ($PSScriptRoot + "\conf.json") -Encoding UTF8 | ConvertFrom-Json)
# 自用调试
# $conf = (Get-Content -Path ./conf.json -Encoding UTF8 | ConvertFrom-Json)
# 共享目录路径
$sharebase = "\\" + $conf.ServerName + "\Agilent\"
# 安装文件文件夹名称
$installbase = $sharebase + "M8490-60006_OpenLABCDS-2.4.0.695\"
# .net安装文件存放位置（Win10的sxs下面的文件也在这里）
$netbase = $sharebase + "dotnet\"
# AIC的安装属性文件名
$aicprop = $installbase + "aic.properties"
# 客户端的安装属性文件
$cltprop = $installbase + "clt.properties"
# 补丁安装属性文件,暂时禁用，目前可以使用客户端安装属性文件正常安装
# $updprop = $cltprop
# 需要安装/更新的驱动，LabAdvisor也要放到这里面，注意要把MSI的安装文件放过来，不是散文件,安装时按按文件名称排序安装。
$drvbase = $installbase + "Drivers\"

# 安装程序，现场准备程序，补丁包以及Adobe Reader的执行程序名
$cdsinstaller = $installbase + "Setup\CDSInstaller.exe"
$cdshf = $installbase + "OpenLAB_CDS_Update.exe"
# 英文版adobe及现场准备路径
if ([System.Globalization.Cultureinfo]::InstalledUICulture -eq "1033") {
    $adobe = (Get-ChildItem -Path ($installbase + "Setup\Tools\Adobe\Reader\*") -Include *US*).FullName
    $siteprep = $installbase + "Setup\Tools\SitePrep\ENU\"
}
else {
    # 中文版adobe及现场准备路径
    $adobe = (Get-ChildItem -Path ($installbase + "Setup\Tools\Adobe\Reader\*") -Include *CN*).FullName
    $siteprep = $installbase + "Setup\Tools\SitePrep\CHS\"
}

$sitepreptool = $siteprep + "Bin\SitePrepTool.exe"
# 现场准备AIC配置文件
$aicprep = $siteprep + "ProductChecks\OpenLAB CDS.xml"
# 现场准备AIC产品名
$siteprep_aicpn = "OpenLAB CDS"
# 现场准备CLIENT配置文件
$cltprep = $siteprep + "ProductChecks\OpenLAB CDS.xml"
# 现场准备CLIENT产品名
$siteprep_cltpn = "OpenLAB CDS"
# 产品版本(xml中的version项目，似乎可以不填)
$pdcv = ""

# 日志位置，不用管
$logbase = $env:ProgramData + "\Agilent\InstallLogs"


# 安装NetFX，检测到WIN10安装3.5，检测到WIN7安装4
if ([System.Environment]::OSVersion.Version.Major -eq 10) {
    if ((Get-WindowsOptionalFeature -Online | Where-Object { $_.FeatureName -eq "WCF-NonHTTP-Activation" }).State -eq "disabled") {
        Enable-WindowsOptionalFeature -Online -FeatureName NetFx3, WCF-NonHTTP-Activation -All -LimitAccess -Source $netbase
    }
}
else {
    $netexes = Get-ChildItem -Path $netbase *.exe | Sort-Object -Property Length -Descending
    foreach ($netexe in $netexes) {
        Start-Process  -FilePath $netexe -ArgumentList "/q /norestart" -Wait
    }
}

# 首先进行现场准备检查，之后安装工作站
$docpath = $env:USERPROFILE + "\Documents\"
New-Item -ItemType Directory -Path C:\ProgramData\Agilent\InstallLogs -ErrorAction SilentlyContinue
Start-Process EXPLORER -ArgumentList 'C:\ProgramData\Agilent\InstallLogs'

# 判断是否AIC
if ($conf.SpecifyComputerName.Specify) {
    $isAIC = ($conf.SpecifyComputerName.AICNameList -contains $env:COMPUTERNAME)
}
else {
    $isAIC = $env:COMPUTERNAME -match "AIC"
}

# 根据上一步判断结果安装工作站，安装前首先进行系统检查
if ($isAIC) {
    if ($conf.SitePrep.Check) {
        Start-Process $sitepreptool  -ArgumentList "configpath=`"$($aicprep)`" reportpath=$docpath orgname=`"$($conf.SitePrep.Info.OrgName)`" orgloc=`"$($conf.SitePrep.Info.OrgLoc)`" contactname=`"$($conf.SitePrep.Info.ContactorName)`" contactjobtitle=`"$($conf.SitePrep.Info.ContactorTitle)`"  mode=silent productname=`"$siteprep_aicpn`" productversion=$pdcv connectmode=local checktype=preinstall textreportname=stieprepreport.txt correctivexmlpath=OLAIC_CorrectiveActions.xml" -Wait

    }
    Start-Process $cdsinstaller -ArgumentList "-s -c $aicprop" -Wait
    # 卸载Sample Schedule插件，不卸载会在系统里报告大量错误
    # Start-Process msiexec -ArgumentList "/x {645F3E18-1ED9-458F-A8A9-2EF44104B074} /qn" -wait
    
}
else {
    if ($conf.SitePrep.Check) {
        Start-Process $sitepreptool  -ArgumentList "configpath=`"$($cltprep)`" reportpath=$docpath orgname=`"$($conf.SitePrep.Info.OrgName)`" orgloc=`"$($conf.SitePrep.Info.OrgLoc)`" contactname=`"$($conf.SitePrep.Info.ContactorName)`" contactjobtitle=`"$($conf.SitePrep.Info.ContactorTitle)`"  mode=silent productname=`"$siteprep_cltpn`" productversion=$pdcv connectmode=local checktype=preinstall textreportname=stieprepreport.txt correctivexmlpath=OLC_CorrectiveActions.xml" -Wait

    }
    Start-Process $cdsinstaller -ArgumentList "-s -c $cltprop" -Wait
}

# 如果补丁文件存在，则安装
if (Test-Path  $cdshf) {
    Start-Process $cdshf  -ArgumentList "-s -c $cltprop" -Wait
}
# 安装aodbe阅读器
Start-Process $adobe -ArgumentList "/sAll" -Wait

# 驱动
Get-ChildItem -Path $drvbase *.msi -Recurse | Sort-Object -Property Name | ForEach-Object {
    $msi = $_.FullName
    Start-Process MSIEXEC -ArgumentList "/qn /i `"$msi`"" -Wait
}

# 生成总结报告
$lastinst = Get-ChildItem $logbase | Sort-Object -Property CreationTime -Descending | Select-Object -First 1
$instlog = Get-ChildItem $lastinst.FullName
$inststate = (Get-Content $instlog.fullname | Where-Object { $_ -Match "已安装产品" }).Substring(62)
Add-Content -Path ($sharebase + "Install_Summary.log") -Value ($env:COMPUTERNAME + "__" + $inststate)

# 运行SVT工具，默认为监测到的产品生成简短报告
$svthome = (Get-ItemProperty -Path "HKLM:SOFTWARE\WOW6432Node\Agilent Technologies\IQTool").InstallLocation
ForEach ($reffile in (Get-ChildItem -path ($svthome + "\IQProducts") -Recurse *.xml)) {
    [xml] $isbase = Get-Content $reffile.FullName -Encoding UTF8
    if ($null -ne $isbase.PRODUCT.BASE) {
        $proname = $proname + "," + $isbase.PRODUCT.NAME.'#text'
    }
}
if ($conf.SVTAll) {
    Start-Process ($svthome + "\Bin\SFVTool.exe") -ArgumentList "-qt -showall -p:`"$proname`" -pdf -xml"  -Wait
}  
else {
    Start-Process ($svthome + "\Bin\SFVTool.exe") -ArgumentList "-qt -shownothing -p:`"$proname`" -pdf -xml" -Wait
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

# 检测生成的svt报告是否均为pass，在桌面上生成一个汇总报告
Get-ChildItem $sv *.xml -Recurse | Sort-Object -Property CreationTime -Descending | Select-Object -First $rptnum.Count |
ForEach-Object {
    [xml] $svstate = Get-Content $_.FullName
    Add-Content -Path ($sharebase + "SVT_Summary.log") -Value ($svstate.QualifiedSuiteResultCollection.QualifiedSuiteResult.Status + "`t" + $env:COMPUTERNAME + "____" + $svstate.QualifiedSuiteResultCollection.QualifiedSuiteResult.Name )
}

# WIN10 1903服务修正
if ((Get-WindowsOptionalFeature -Online -FeatureName "WCF-TCP-Activation45").State -eq "Disabled") {
    Enable-WindowsOptionalFeature -Online -FeatureName "WCF-TCP-Activation45" -All -NoRestart
}
if ((Get-Service -Name "NetTcpPortSharing").StartType -ne "Automatic" ) {
    Set-Service -Name "NetTcpPortSharing" -StartupType Automatic
}

# 重启计算机
Restart-Computer -Force