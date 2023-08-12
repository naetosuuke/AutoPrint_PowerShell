
#最後のダイアログ用
Add-Type -AssemblyName System.Windows.Forms 
Add-Type -AssemblyName System.Drawing 

#文字化け対策?
$OutputEncoding = [Console]::OutputEncoding 

#フォルダ指定ダイアログ
#[void] [Reflection.Assembly] ::LoadWithPartialName ("System.Windows.Forms") 
$FolderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog -Property @{ 
    RootFolder = Desktop
    Description = 印刷対象のフォルダを選択してください
}

if($FolderBrowser.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK){ 
}else{
    exit 
}

if($FolderBrowser -eq $null) {
    Write-Host 変数はNULLです。
    pause 
    exit 
}else{
    Write-Host passed 
}

#選択したフォルダのパスを宣言 
Write-Host $FolderBrowser.SelectedPath 
$folder = $FolderBrowser.SelectedPath 

#各種宣言 
$Kari = (Join-Path $folder ¥Kari) 
$PrintComplete = (Join-Path $folder ¥Allprint complete) 

#>>>>>>>>>>印刷処理部分>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> 

#操作先ディレクトリ選択 
Set-Location $folder 

Write-Host "$_ フォルダ直下のファイルをすべて印刷する!" 
Write-Host "※ファイル間のスリ-プ処理:2秒`r`n"
#ドラフト印刷よけるための宣言 

$DRIV = @(Dir *invoice* -include *draft* | Sort-Object -property LastWriteTime -Descending) 
$DRIV2 = @(Dir *invoice* ゙ -inc1ude  *通関* | Sort-Object -property LastWriteTime -Descending) 
$DRPL = @(Dir *packinglist* -include *draft* | Sort-Object -property LastWriteTime -Descending) 
$DRPL2 = @(Dir *packinglist* -include・*通関* | Sort-Object -property LastWriteTime -Descending) 
Write-Host "除外対象`r`n$DRIV`r`n$DRIV2`r`n$DRPL`r`n$DRPL2`r`n "

#順番が決まっているファイルの印刷を一回にがす 
New-Item Kari -ItemType Directory 


<#ファイル名の昇順で印刷実行(Dir=Gct-Childitem) 
    手段① 
    ファイル名に特定の単語(PO)をもつ者を順番でGetchildでひろい、 
    あまりものは一番最初に印刷するようにする
    (添付書類→PO→SI→納品書→安全確認書→通関指示書→出荷指示書→PL→IV→託送PL→託送 IV→BL/AWB/SWB-,受領書→FLIGHT SCHEDULE→船完) 
#>

Write-Host ">>>>>>>>>>>>>>>>>>自動で整列可能な書類>>>>>>>>>>>>>>>>>>>>>>>>>>"

$Aligned_docs = @()

$Aligned_docs += (Dir *SB1* -include *.pdf -exclude *SI_* | Write-Host | Move-Item -Destination $Kari) #| ForEach-Object { $_ }
$Aligned_docs += (Dir *SF1* -include *.pdf -exclude *SI_* | Write-Host | Move-Item -Destination $Kari) #| ForEach-Object { $_ }
$Aligned_docs += (Dir *SI1* -include *.pdf -exclude *SI_* | Write-Host | Move-Item -Destination $Kari) #| ForEach-Object { $_ }
$Aligned_docs += (Dir *SM1* -include *.pdf -exclude *SI_* | Write-Host | Move-Item -Destination $Kari) #| ForEach-Object { $_ }
$Aligned_docs += (Dir *SO1* -include *.pdf -exclude *SI_* | Write-Host | Move-Item -Destination $Kari) #| ForEach-Object { $_ }
$Aligned_docs += (Dir *SA1* -include *.pdf -exclude *SI_* | Write-Host | Move-Item -Destination $Kari) #| ForEach-Object { $_ }
$Aligned_docs += (Dir *SQ1* -include *.pdf -exclude *SI_* | Write-Host | Move-Item -Destination $Kari) #| ForEach-Object { $_ }
$Aligned_docs += (Dir *SP1* -include *.pdf -exclude *SI_* | Write-Host | Move-Item -Destination $Kari) #| ForEach-Object { $_ }

$Aligned_docs += (Dir *SI_S* -include *.pdf | Write-Host | Move-Item -Destination $Kari) #| ForEach-Object { $_ }
$Aligned_docs += (Dir *納品* -include *.pdf | Write-Host | Move-Item -Destination $Kari) #| ForEach-Object { $_ }
$Aligned_docs += (Dir *安全確認* -include *.pdf | Write-Host | Move-Item -Destination $Kari) #| ForEach-Object { $_ }
$Aligned_docs += (Dir *通関指示* -include *.pdf | Write-Host | Move-Item -Destination $Kari) #| ForEach-Object { $_ }
$Aligned_docs += (Dir *ブッキング依頼* -include *.pdf | Write-Host | Move-Item -Destination $Kari) #| ForEach-Object { $_ }
$Aligned_docs += (Dir *出荷指示* -include *.pdf | Write-Host | Move-Item -Destination $Kari) #| ForEach-Object { $_ }
$Aligned_docs += (Dir *packinglist* -include *.pdf -exclude *託送*, *通関*,*draft*| Write-Host | Move-Item -Destination $Kari) #| ForEach-Object { $_ }
$Aligned_docs += (Dir *invoice* -include *.pdf -exclude *託送*,*通関*,*draft*| Write-Host | Move-Item -Destination $Kari) #| ForEach-Object { $_ }
$Aligned_docs += (Dir *packinglist* -include *.pdf, *託送* | Write-Host | Move-Item -Destination $Kari) #| ForEach-Object { $_ }
$Aligned_docs += (Dir *invoice* -include *.pdf, *託送* | Write-Host | Move-Item -Destination $Kari) #| ForEach-Object { $_ }
$Aligned_docs += (Dir *SIS* -include *.pdf | SOrt-Object -property LastWriteTime -Descending | Write-Hos| Move-Item -Destination $Karit ) #| ForEach-Object { $_ }
$Aligned_docs += (Dir *受領書* -include *.pdf | Write-Host | Move-Item -Destination $Kari) #| ForEach-Object { $_ }
$Aligned_docs += (Dir *FLIGHT* -include *.pdf | Write-Host | Move-Item -Destination $Kari) #| ForEach-Object { $_ }
$Aligned_docs += (Dir *SIS* -include *.pdf | Write-Host | Move-Item -Destination $Kari) #| ForEach-Object { $_ }


#順番きまってる書類の印刷 
Set-Location $Kari 

Write-Host "`r`n>>>>>>>>>>>>>>>整列分印刷中.....>>>>>>>>>>>>>>>>>>>>>>>`r`n"

$Aligned_docs | ForEach-Object {
    Write-Host $_.Name 
    #ファイル名表示後、印刷 
    Start-Process $_.FulIName -VerbPrint | Stop-Process 
    Start-Sleep -s 2 
}


<#

$FUNA = Dir *SIS* | Sort-Object -property LastWriteTime -Descending | ForEach { 
    Write-Host $_.Name 
    #ファイル名表示後、印刷 
    Start-Process $_.FulIName -VerbPrint | Stop-Process 
    Start-Sleep -s 2 
}

$FLIGHT = Dir *FLIGHT* | Sort-Object -property LastWriteTime -Descending | ForEach { 
    Write-Host $_.Name 
    #ファイル名表示後、印刷 
    Start-Process $_.FulIName -VerbPrint | Stop-Process 
    Start-Sleep -s 2 
}

$JURYOSHO= Dir *受領書* | Sort-Object -property LastWriteTime -Descending | ForEach { 
    Write-Host $_.Name 
    #ファイル名表示後、印刷 
    Start-Process $_.FulIName -VerbPrint | Stop-Process 
    Start-Sleep -s 2 
}

$WB = Dir *AWB*, *SWB*, *BL* | Sort-Object -property LastWriteTime -Descending | ForEach { 
    Write-Host $_.Name 
    #ファイル名表示後、印刷 
    Start-Process $_.FulIName -VerbPrint | Stop-Process 
    Start-Sleep -s 2 
}

$TAKUSOIV = Dir *invoice*, -include *託送゙* | Sort-Object -property LastWriteTime -Descending | ForEach{ 
    Write-Host $_-Name 
    #ファイル名表示後、印刷 
    Start-Process $_FullName-Verb Print | Stop-Process 
    Start-Sleep -s 2 
}

$TAKUSOPL = Dir *packinglist* -include *託送* | Sort-Object -property LastWriteTime -Descending | ForEach{ 
    Write-Host $_-Name 
    #ファイル名表示後、印刷 
    Start-Process $_FullName-Verb Print | Stop-Process 
    Start-Sleep -s 2 
}

$IV = Dir *invoice* -exclude *draft*, *託送* | Sort-Object -property LastWriteTime -Descending | ForEach { 
    Write-Host $_.Name 
    #ファイル名表示後、印刷 
    Start-Process $_.FulIName -VerbPrint | Stop-Process 
    Start-Sleep -s 2 
}
$PL = Dir *packinglist* -exclude *draft*, *託送* | Sort-Object -property LastWriteTime -Descending | ForEach{ 
    Write-Host $_.Name 
    #ファイル名表示、印刷 
    Start-Process $_FullName-Verb Print | Stop-Process 
    Start-Sleep -s 2 
}

$BOOKING = Dir *ブッキング依頼* | Sort-Object -property LastWriteTime -Descending | ForEach{ 
    Write-Host $_-Name 
    #ファイル名表示後、印刷 
    Start-Process $_FullName-Verb Print | Stop-Process 
    Start-Sleep -s 2 
}

$SHUKKA - Dir ゙出荷指示書・| Sort-Object -property LastWriteTime -Descending | ForEach{ 
    Write-Host $_-Name 
    #ファイル名表示後、印刷 
    Start-Process $_FullName-Verb Print | Stop-Process 
    Start-Sleep -s 2 
}

$TSUKAN = Dir *通関指示* | Sort-Object -property LastWriteTime -Descending | ForEach{ 
    Write-Host $_-Name 
    #ファイル名表示後、印刷 
    Start-Process $_FullName-Verb Print | Stop-Process 
    Start-Sleep -s 2 
}


$ANZEN = Dir *安全確認書* | Sort-Object -property LastWriteTime -Descending | ForEach{ 
    Write-Host $_-Name 
    #ファイル名表示後、印刷 
    Start-Process $_FullName-Verb Print | Stop-Process 
    Start-Sleep -s 2 
}

$NOUHIN = Dir *納品書* | Sort-Object -property LastWriteTime -Descending | ForEach{ 
    Write-Host $_-Name 
    #ファイル名表示後、印刷 
    Start-Process $_FullName-Verb Print | Stop-Process 
    Start-Sleep -s 2 
}


$POB - Dir *SB1*-exclude *SI_*, *LC*| Sort-Object -property LastWriteTime -Descending | ForEach{ 
    Write-Host $_-Name 
    #ファイル名表示後、印刷 
    Start-Process $_FullName-Verb Print | Stop-Process 
    Start-Sleep -s 2 
}

$POF :- Dir *SF 1*-exclude *SI_*, *LC*| Sort-Object -property LastWriteTime -Descending | ForEach{ 
    Write-Host $_-Name 
    #ファイル名表示後、印刷 
    Start-Process $_FullName-Verb Print | Stop-Process 
    Start-Sleep -s 2 
}

$POI - Dir *SII*-exclude *SI_*, *LC*| Sort-Object -property LastWriteTime -Descending | ForEach{ 
    Write-Host $_-Name 
    #ファイル名表示後、印刷 
    Start-Process $_FullName-Verb Print | Stop-Process 
    Start-Sleep -s 2 
}

$POM - Dir *SM1*-exclude *SI_* *LC*[ Sort-Object -property LastWriteTime -Descending lForEach{ Write-Host $_-Name 
#ファイル名表示後、印刷Start-Process $_.FullName -Verb Print l Stop-Process Start-Sleep -s 2 

$POO -- Dir *SOI*-exclude *SI_*, *LC*| Sort-Object -property LastWriteTime -Descending | ForEach{ 
    Write-Host $_-Name 
    #ファイル名表示後、印刷 
    Start-Process $_FullName-Verb Print | Stop-Process 
    Start-Sleep -s 2 
}


$POQ - Dir *SQ1*-exclude *SI_*, *LC* | Sort-Object -property LastWriteTime -Descending | ForEach{ 
    Write-Host $_-Name 
    #ファイル名表示後、印刷 
    Start-Process $_FullName-Verb Print | Stop-Process 
    Start-Sleep -s 2 
}

$POP = Dir *SP1*-exclude *SI_*, *LC* | Sort-Object -property LastWriteTime -Descending | ForEach{ 
    Write-Host $_-Name 
    #ファイル名表示後、印刷 
    Start-Process $_FullName-Verb Print | Stop-Process 
    Start-Sleep -s 2 
}

#>


#操作 Directry $Kari→$folderに戻す) 
Set-Location $folder 

#LC書類を自動印刷しないよう、$Kariに逃がす 
Dir *LC* -inc1ude *.pdf | Move-ltem -Destination $Kari 
Dir *SDS* -include *.pdf | Move-ltem -Destination $Kari  

#順番が決まっていないファイルの印刷 

Write-Host "`r`n>>>>>>>>>>>>>>>>>>未整列分印刷中>>>>>>>>>>>>>>>>>>>>>>>>>>>`r`n"

Dir *.pdf | Sort-Object -property LastWriteTime -Descending | ForEach{ 
    Write-Host $_.Name 
    #ドラフトInvoice, PDFは印刷回避 
    if( ($DRIV -match $_Name) -or ($DRIV2 -match $_.Name)){
        Write-Host "↑↑ドラフトInvoiceが合まれていた為、除外しました。" 
    }elself(($DRPL -match $_.Name) -or ($DRPL2 -match $_.Name)){ 
        Write-Host"↑↑ドラフトPackin8 Listが合まれていた為、除外しました。" 
    }else{
        #ファイル名表示後、実行 
        Start-Process $_.FulIName -Verb Print | Stop-Process 
        #2秒スリ-プ 
        Start-Sleep -s 2 
    }
}
    
    
#一旦逃がしたLDISDSを$folderにもどす 
Set-Location $Kari 
Dir *LC* -include *.pdf | Move-Item -Destination $folder 
Dir *SDS* -include *.pdf | Move-Item -Destination $folder 
Set-Location $folder 

#>>>>>>>>>>>メ-ルひらいて印刷させる>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> 
Write-Host #`r`n>>>>>>>>>>>>>>>>>>>>>>以下メ-ル文章jLCjSDS (手動で印刷してください。)〉>>>>>>>>>>>>>>>>>>>>`r`n# 
#メ-ル、LC、SDS 全件展開
Dir *.msg | Sort-Object -property LastWriteTime -Descending | forEach { 
Write-Host $_.Name Start-Process $_.FullName
}

Dir *LC* -include *.pdf | Sort-Object -property LastWriteTime -Descending | forEach { 
Write-Host $_.Name Start-Process $_.FullName
}

Dir *SDS* -include *.pdf | Sort-Object -property LastWriteTime -Descending | forEach { 
Write-Host $_.Name Start-Process $_.FullName
}

Write-Host "`r`n印刷が完了したら、Enterを押してください。" 
pause 

#印刷済みフォルダの作成、投入 
Write-Host "******************************`r`n*******お待ちください******`r`n****************************" 
Start-Sleep -s 1

New-Item Allprint_complete -ItemType Directory
Dir *.pdf | Move-Item -Destination $PrintComplete 
Dir *.msg | Move-Item -Destination $PrintComplete

#$Kariファイルに入れたものをもとにもどす 
Set-Location $KariDir | Move-Item -Destination $PrintComplete 
Set-Location $folder 

#名称変更処理 
$foldername = (Split-Path. -leaf) 
Set-Location "C:¥":
Stop-Process -Name AcroRd32 -Force 
Stop-Process -Name RdrCEF -Force 
Sleep -s 3 
Remove-Item $Kari 
Rename-Item "$folder" -NewName "【印刷完了】 $foldername" 
Write-Host "**********************************`r`n           COMPLETE! ! !           `r`n**********************************"

Sleep -s 2 
