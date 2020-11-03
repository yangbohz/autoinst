#################################
# 配置部分

# 是否生成啰嗦版SVT报告，默认不生成
$SVTAll = $false

# 配置结束
#################################

$docpath = Join-Path $env:USERPROFILE "Documents"

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

# 删除旧的svt报告
Get-ChildItem $dest SVR*.PDF | Remove-Item -Force

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
