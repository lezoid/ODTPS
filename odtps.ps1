## Office Deployment Toolsの存在を確認
function check_odt() {

    if (Test-Path .\setup.exe) {
        menu1
    } else {
        Write-Host 'Setup.exe of Office Deployment Tool can not found.'
        Write-Output 'Run setup.exe expansion installer. Please expand to the same directory.'
        Invoke-WebRequest -Uri https://download.microsoft.com/download/2/7/A/27AF1BE6-DD20-4CB4-B154-EBAB8A7D4A7E/officedeploymenttool_11107-33602.exe -OutFile ./odt.exe
        ./odt.exe
        Read-Host "完了したら何かキーを入力してください。"
        Remove-Item .\configuration-*
        Remove-Item .\odt.exe
        check_odt

    }
        Write-Output ' '
}


function menu1() {
cls
$mode = Read-Host "Office Deploy PowerShell Script
Source Image Download Mode

1. Download Office365 lasted 32bit Image
2. Download Office365 lasted 64bit Image
3. Download Office2019 ProPlus2019Volume 32bit Image
4. Download Office2019 ProPlus2019Volume 64bit Image
9. Switch Install and Configuration Mode

Z. Delete All Download Images

Enter the number："

if($mode -eq "1"){
   download_office 365_32

}elseif($mode -eq "2"){
   download_office 365_64

}elseif($mode -eq "3"){
   download_office 2019vl_32

}elseif($mode -eq "4"){
   download_office 2019vl_64

}elseif($mode -eq "9"){
   menu2

}elseif($mode -eq "Z"){
   Remove-Item ./Source/* -Recurse
   Read-Host "All in the Source folder have been deleted.
Please enter any key to continue."
   menu1

}else{
   Write-Output 'Unrecognized command'
   menu1
}

}

## Install and Configuration Mode
function menu2() {
   Write-Output 'function_menu2'
   cls
$mode = Read-Host "Office Deploy PowerShell Script
Source Image Install Mode

In order to use the installation mode, 
it is necessary to download the Office image in the Source folder beforehand.

1. Install Office365 lasted 32bit Image
2. Install Office365 lasted 64bit Image
3. Install Office2019 ProPlus2019Volume 32bit Image
4. Install Office2019 ProPlus2019Volume 64bit Image
5. Install Office2019 ProPlus2019 with Visio and Project Volume 32bit Image
5. Install Office2019 ProPlus2019 with Visio and Project Volume 64bit Image
9. Switch Download Mode

Enter the number："

if($mode -eq "1"){
   Copy-Item .\xmls\configuration-Office365-x86.xml .\xmls\install.xml
   replace_config O365_x86\
   install_office

}elseif($mode -eq "2"){
   Copy-Item .\xmls\configuration-Office365-x64.xml .\xmls\install.xml
   replace_config O365_x64\
   install_office

}elseif($mode -eq "3"){
   Copy-Item .\xmls\configuration-Office2019Enterprise-x86.xml .\xmls\install.xml
   replace_config O2019VL_x86\
   install_office

}elseif($mode -eq "4"){
   Copy-Item .\xmls\configuration-Office2019Enterprise-x64.xml .\xmls\install.xml
   replace_config O2019VL_x64\
   install_office

}elseif($mode -eq "5"){
   Copy-Item .\xmls\configuration-Office2019Enterprise_for_visioandproject-x86.xml .\xmls\install.xml
   replace_config O2019VL_x86\
   install_office

}elseif($mode -eq "6"){
   Copy-Item .\xmls\configuration-Office2019Enterprise_for_visioandproject-x64.xml .\xmls\install.xml
   replace_config O2019VL_x64\
   install_office

}elseif($mode -eq "9"){
   menu1

}else{
   Write-Output 'Unrecognized command'
   menu2
}
}

## Office のダウンロード処理
function download_office($i) {
    if($i -eq "365_32"){
        Write-Output 'Office365 lasted 32bit Image Image Downloading... '
        ./setup.exe /download ./xmls/download-Office365-x86.xml
        Read-Host  'Download Complete!
         
Please enter any key to continue.'
        menu1

    }elseif($i -eq "365_64"){
        Write-Output 'Office365 lasted 64bit Image Image Downloading... '
        ./setup.exe /download ./xmls/download-Office365-x64.xml
        Read-Host  'Download Complete!
         
Please enter any key to continue.'
        menu1

    }elseif($i -eq "2019vl_32"){
        Write-Output 'Office2019 ProPlus2019Volume 32bit Image Downloading... '
        ./setup.exe /download ./xmls/download-Office2019Enterprise-x86.xml
        Read-Host 'Download Complete!
         
Please enter any key to continue.'
        menu1

    }elseif($i -eq "2019vl_64"){
        Write-Output 'Office2019 ProPlus2019Volume 64bit Image Downloading... '
        ./setup.exe /download ./xmls/download-Office2019Enterprise-x64.xml
        Read-Host 'Download Complete!
         
Please enter any key to continue.'
        menu1

    }else{
        Write-Output 'else $i'
    }

}

## Office のインストール処理
function install_office() {
   Write-Output 'Office installation in progress...'
   ./setup.exe /configure .\xmls\install.xml
   Read-Host "Install Complete! 
Please enter any key to continue."
   menu2
}

## Replace Configuration File
function replace_config($i) {
    $str_path = (Convert-Path .)
    $source_dir = "\Source\"
    $source_path = $str_path + $source_dir + $i

    if (Test-Path .\Source\$i) {
        $install_config= Get-Content .\xmls\install.xml | foreach { $_ -replace "OFFICE-DIR", $source_path }
        $install_config | Out-File .\xmls\install.xml -Encoding UTF8
    }
    else {
        Read-Host "Source file does not exist.
Please execute download in download mode.

Please enter any key to continue."
        menu2

    }
        Write-Output ' '
}

check_odt