# ODTPS
Office Deployment Tool PowerShell Script  
  
# 説明
Office Deployment Tools (略.ODT)を利用したOfficeイメージのダウンローダー 兼 インストーラー PowerShell Scriptです。  
Office2019からISOが提供されなくなったのですが、毎回setup.exeを入手してXMLを記述するのもめんどくさいので自分用に作りました。  

# 使い方
odtps.ps1を実行してください。  
初回起動時にODTをMicrosoftからダウンロードを実行し、展開用インストーラーが実行されるので、同ディレクトリ内に展開をしてください。  
setup.exeがODTPSと同じディレクトリ内に存在する状態になると、Officeイメージのダウンロード及びインストールができます。  

# ディレクトリ及びファイル構成の説明
.\Source\	Officeのダウンロードしたイメージの展開先のディレクトリとして利用します。  
.\xmls\		ダウンロード及びインストール用のXML定義テンプレートが配置されています。  
odtps.ps1	odtpsの実行ファイルです。  
setup.exe	初回実行時に展開用インストーラーが開くので、odtpsが存在するディレクトリを指定して展開してください。  
readme.txt	今読んでいるファイルです。  
  
  
# Description
Office image downloader and installer PowerShell Script using Office Deployment Tools (.ODT).  
ISO is no longer being offered from Office 2019, but it is troublesome to write setup.exe every time and write XML, so I made it for myself.  

# How to use
Please execute odtps.ps1.  
Please download the ODT from Microsoft at the first startup and execute the deployment installer, so please deploy it in the same directory.  
When setup.exe is in the same directory as ODTPS, you can download and install the Office image.  

# Description of directory and file structure
. \ Source \ Office Used as the destination directory of downloaded images.  
. \ xmls \ XML definition template for download and installation is located.  
odtps.ps1 Executable file of odtps,  
setup.exe Since the deployment installer opens when you run it for the first time, please specify and expand the directory where odtps exists.  
readme.txt This is the file you are reading now.  
