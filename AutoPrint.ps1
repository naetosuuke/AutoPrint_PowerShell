OCR

〈# 
ディレクトリ中に別途フォルダがある場合、印刷処理をループさせる処理(未実装) 
おもに2件以上のP〇がある場合を想定
Set-Location $folder
While (Test-Path $folderY*-exclude *Allprint*) {Write-Host "朱印刷のフォルダがある可能性があります。" 
pause 
}
Write-host "検知しませんでした" Pause 
#>

#最後のダイアログ用
Add-Type -AssemblyName System.Windows,Forms 
Add-Type -AssemblyName System.Drawing 

#文字化け対策?
$OutputEncoding = [Console]::OutputEncoding 

#フォルダ指定ダイアログ
#[void] [Reflection.Assembly] ::LoadWithPartialName ('~System.Windows.Forms'~:) 
Add-Type -AssemblyName System.Windows.Forms
$FolderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog -Property @{ 
    RootFolder = 'Desktop'
    Description = '印刷対象のフォルダを選択してください'
}

if($FolderBrowser.ShowDialog() -eq [System.Windows.Forms.DialogResult] ::OK){ 
}else{
    exit 
}

if($FolderBrowser -eq $null) {
    Write-Host '変数はNULLです。'
    pause 
    exit 
}else{
    Write-Host 'passed' 
}

#選択したフォルダのパスを宣言 
Write-Host $FolderBrowser.SelectedPath 
$folder = $FolderBrowser.SelectedPath 

#各種宣言 
$OutputEncoding = [Console]::OutputEncoding
$Kari = (Join-Path $folder ¥Kari) 
$PrintComplete = (Join-Path $folder ¥Allprint complete) 

#>>>>>>>>>>印刷処理部分>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> 

#操作先ヂィレクトジ選択 
Set-Location $folder 

Write-Host "$_ フォルダ痕下のファイルをすべて印刷する!" 
Write-Host "※ファイル間のスリープ処理:2秒'r'n"
#ドラフト印刷よけるための宣言 

$DRIV = @(Dir *invoice* -include *draft* | Sort-Object -property LastWriteTime -Descending) 
$DRIV2 = @(Dir *invoice* ゙ -inc1ude  *通関* | Sort-Object -property LastWriteTime -Descending) 
$DRPL = @(Dir *packinglist* -include *draft* | Sort-Object -property LastWriteTime -Descending) 
$DRPL2 = @(Dir *packinglist* -include・*通関* | Sort-Objectーproperty LastWriteTime -Descending) 
Write-Host "除外対象'r'n$DRIV'r'n$DRIV2'r'n$DRPL'r'n$DRPL2'r'n "

#順番が決まっているファイルの印刷を一回にがす 
New-Item Kari -ItemType Directory 


<#ファイル名の昇順で印刷実行(Dir=Gct-Childitem) 
    手段① 
    ファイル名に特定の単語(PO)をもつ者を順番でGetchildでひろい、 
    あまりものは一番最初に印刷するようにする
    (添付書類→PO→SI→納品書→安全確認書→通関指示書→出荷指示書→PL→IV→託送PL→託送 IV→BL/AWB/SWB-,受領書→FLIGHT SCHEDULE→船完) 
#>

Write-Host ">>>>>>>>>>>>>>>>>>自動で整列可能な書類>>>>>>>>>>>>>>>>>>>>>>>>>>"

$POB = Dir *SB1* -include *.pdf -exclude *SI_* | Write-Host | Move-Item -Destination $Kari
$POF = Dir *SF1* -include *.pdf -exclude *SI_* | Write-Host | Move-Item -Destination $Kari
$POI = Dir *SI1* -include *.pdf -exclude *SI_* | Write-Host | Move-Item -Destination $Kari
$POM = Dir *SM1* -include *.pdf -exclude *SI_* | Write-Host | Move-Item -Destination $Kari
$POO = Dir *SO1* -include *.pdf -exclude *SI_* | Write-Host | Move-Item -Destination $Kari
$POA = Dir *SA1* -include *.pdf -exclude *SI_* | Write-Host | Move-Item -Destination $Kari
$POQ = Dir *SQ1* -include *.pdf -exclude *SI_* | Write-Host | Move-Item -Destination $Kari
$POP = Dir *SP1* -include *.pdf -exclude *SI_* | Write-Host | Move-Item -Destination $Kari

$SI = Dir *SI_S* -include *.pdf | Write-Host | Move-Item -Destination $Kari
$NOUHIN = Dir *納品* -include *.pdf | Write-Host | Move-Item -Destination $Kari
$ANZEN = Dir *安全確認* -include *.pdf | Write-Host | Move-Item -Destination $Kari
$TSUKAN = Dir *通関指示* -include *.pdf | Write-Host | Move-Item -Destination $Kari
$BOOKING = Dir *ブッキング依頼* -include *.pdf | Write-Host | Move-Item -Destination $Kari
$SHUKKA = Dir *出荷指示* -include *.pdf | Write-Host | Move-Item -Destination $Kari
$PL = Dir *packinglist* -include *.pdf -exclude *託送*, *通関*,*draft*| Write-Host | Move-Item -Destination $Kari
$IV = Dir *invoice* -include *.pdf -exclude *託送*,*通関*,*draft*| Write-Host | Move-Item -Destination $Kari
$TAKUSOPL = Dir *packinglist* -include *.pdf, *託送* | Write-Host | Move-Item -Destination $Kari
$TAKUSOIV = Dir *invoice* -include *.pdf, *託送* | Write-Host | Move-Item -Destination $Kari
$WB = Dir *SIS* -include *.pdf | SOrt-Object -property LastWriteTime -Descending | Write-Hos| Move-Item -Destination $Karit 
$JURYOSHO = Dir *受領書* -include *.pdf | Write-Host | Move-Item -Destination $Kari
$FLIGHT = Dir *FLIGHT* -include *.pdf | Write-Host | Move-Item -Destination $Kari
$FUNA = Dir *SIS* -include *.pdf | Write-Host | Move-Item -Destination $Kari



#順番きまってる書類の印刷 
Set-Location $Kari 

Write-Host "'r'n>>>>>>>>>>>>>>>整列分印刷中.....>>>>>>>>>>>>>>>>>>>>>>>'r'n"


$FUNA = Dir *SIS* ]Sort-Object -property LastWriteTime -Descending [ ForEach{ Write-Host $_-Name 
#ファイル名表示後、印刷 
Start-Process $_.FulIName -VerbPrint lStop-P rocess 
Start-Sleep -s 2 

$FLIGHT - Dir *FLIGHT* I Sort-Object -property LastWriteTime -Descending I ForEach{ Write-Host $_-Name 
#ファイル名表示後、印刷 
Start-Process $_.FullName -Verb Printl Stop-Pr ocess 
Start-Sleep -s 2 

$JURYOSH〇= Dir ゙受領書・I SortーObjectーproperty LastWriteTime -Descending 1 ForEach{ Write-Host $_.Name 
#ファイル名表示後、印刷 
Start-Process $_.FullName -Verb Print[ Stop-Pr ocess 
Start-Sleep -s 2 

$WB - Dir *AWB*, *SWB*, *BL* l Sort-Object -property LastWriteTime -Descending ] ForEach{ Write-Host $_-Name 
#フアイノレ名表示後、印刷 
Start-Process $_FullName-Verb Print lStop-Proce ss 
Start-Sleep -s 2 
}

$TAKUSOIV-Dir*invoice* q ーinclude゙託送゙|Sort-Object -property LastWriteTime -Descending l ForEach{ 
Write-Host $_-Name#ファイル名表示後、印刷Start-Process $_.FullName -Verb Print l Stop-Process Start-Sleep -s 2 
$TAKUSOPL-Dir*packinglist*-include *jt2*I Sort-Object -property LastWriteTime -Descending [ ForEach{ 
Write-Host $_-Name #ファイル名表示後、印刷 
Start-Process $_.FullName -Verb Print l Sto p-Process 
Start-Sleep -s 2 

$IV - Dir *invoice*-exclude *draft* *jtl*I Sort-Object -property LastWriteTime -Descending I ForEach{ Write-Host $_,Name 
#ファイル名表示後、印刷 
Start-Process $_.FullName -Verb Print [ Stop-Process 
Start-Sleep -s 2 

$PL -- Dir *packinglist*-exclude *draft*, *22* l Sort-Object -property LastWriteTime -Descending l ForEach{ Write-Host $_-Name 
#ファイル名表示後、印刷 
Start-Process $_.FullName -Verb Print 1 Stop-Process 
Start-Sleep -s 2 

$BOOKING = Dir・ブッキング依頼・I SortーObject -property LastWriteTime一Descending l ForEach{ Write-Host $_.Name 
#ファイル名表示後、印刷 
Start-Process $_.FullName -Verb Print [ Stop-Process 
Start-Sleep -s 2 

$SHUKKA - Dir ゙出荷指示書・l Sort一0bjectーproperty LastWriteTime -Descending l ForEach{ Write-Host $_-Name 
#ファイル名表示後、印刷 
Start-Process $_.FullName -Verb Print l Stop-Process 
Start-SIeep -s 2 

$TSUKAN = Dir・通関指示・j Sort-Object -property LastWriteTime -Descending l ForEach{ Write-Host $_-Name 
#ファイル名表示後、印刷 
Start-Process $_.FullName -Verb Print l Stop-Process 
Start-Sleep -s 2 

$ANZEN - Dir・安全確認書・] Sort-Object -property LastWriteTime一Descending l ForEach{ Write-Host $_-Name 
#ファィノレ名表示後、印刷 
Start-Process $_.FullName -Verb Print] Stop-Proc ess 
Start-Sleep -s 2 

$NOUHIN - Dir *M@$* I Sort-Object -property LastWriteTime -Descending l ForEach{ Write-Host $_-Name 
#ファイル名表示後、印刷Start-Process $_-FuHName -Verb Print [ Stop-Process Start-Sleep -s 2 
$SI - Dir '~'SI_S" I Sort-Object -property LastWriteTime -Descending I ForEach{ Write-Host $_-Name 
#ファイル名表示後、印刷 
Start-Process $_.FulIName -Verb Print l Stop-Process 
Start-Sleep -s 2 

$POB - Dir *SB1*-exclude *SI_*, *LC*l Sort-Object -property LastWriteTime -Descending I ForEach{ Write-Host $_.Name 
#ファイル名表示後、印刷 
Start-Process $_.FuHName -Verb Print l Stop-Process 
Start-Sleep -s 2 

$POF :- Dir *SF 1*-exclude *SI_*, *LC*I Sort-Object -property LastWriteTime -Descending I ForEach{ Write-Host $_.Name 
#ファイル名表示後、印刷 
Start-Process $_.FullName -Verb Print [ Stop-Process 
Start-Sleep -s 2 

$POI - Dir *SII*-exclude *SI_*, *LC*l Sort-Object -property LastWriteTime -Descending I ForEach{ Write-Host $_.Name 
#ファイル名表示後、印刷Start-Process $_-FullName -Verb Print l Stop-Process Start-Sleep -s 2 

$POM - Dir *SM1*-exclude *SI_* *LC*[ Sort-Object -property LastWriteTime -Descending lForEach{ Write-Host $_-Name 
#ファイル名表示後、印刷Start-Process $_.FullName -Verb Print l Stop-Process Start-Sleep -s 2 

$POO -- Dir *SOI*-exclude *SI_*, *LC*l Sort-Object -property LastWriteTime -Descending lForEach{ Write-Host $_.Name 
#ファイル名表示後、印刷Start-Process $_-FullName -Verb Print ] Stop-Process Start-Sleep -s 2 
-exclude *SI_*, *LC*I Sort-Object -property LastWriteTime -Descending ] ForEach{ Write-Host $_-Name 
#ファイル名表示後、印刷 
Start-Process $_.FullName -VerbPrint[Stop-Proces s 
Start-Sleep -s 2 

$POQ - Dir *SQ1*-exclude *SI_*, *LC*I Sort-Object -property LastWriteTime -Descending IForEach{ Write-Host $_-Name 
#ファイル名表示後、印刷 
Start-Process $_.FullName -Verb Print l Stop-Proces s 
Start-Sleep -s 2 

$POP -- Dir *SP1*-exclude *SI_*, *LC*ISort-Object -property LastWriteTime -Descending I ForEach{ Write-Host $_-Name 
#ファイル名表示後、印刷 
Start-Process $_.FullName -Verb Print ] Stop-Proces s 
Start-Sleep -s 2 
gift Directry $Kari-,*folderに戻す) Set-Location $folder 
#LC書類を自動印刷しないよう、$Kariに逃がす 
Dir *LC*-mclude *.pdf j Move-ltem -Destination $Kari 
Dir *SDS*-include *.pdf IMove-Item -Destipation $Kari 
卓順番が決まっていないファイルの印刷 
WriteーHo・t **'r'n>>>>>>>>>>>>>>>>>>未整列分印刷中,,,,,>>>>>>>>>>>>>>>>>>>>>>>>>>>'r'nl* 
Dir *.pdf l Sort-Object -property LastWriteTime -Descending l ForEach{ Write-Host $_.Name 

#ドラフトInvoicel PDFは印刷回避 
if( ($D拓V -m虹。h $_Name)Write-Host *t↑↑ドラフトInvoiceが合まれていた為、除外しました。" 
}elself( ($DRPLーmatch $_ Name )-or ( $DRPL2 -match $_.N・m°)){ Write-Host"↑↑ドラフトPackin8 Listが合まれていた為、除外しました。" 
#ファイル名表示後、実行 
Start-Process $_.FulIName -Verb Print lStop-Proc ess 
#2秒スジープ Start-Sleep -s 2 
#一旦逃がしたLDISDSを$folderにもどす 
Set-Location $Kari 
Dir *LC*-mclude *.pdfIMove-Item -Destination *folder 
Dir *SDS*-mclude *.pdfIMove-Item -Destination $folder Set-Location $folder 
#>>〉>>>>>>>>>メールひらいて印刷させる>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> 
Writ。一H。・t **'r'n>>>>>>>>>>>>>>>>>>>>>>以下メール文章jLCjSDS (手動で印刷してください。)〉>>>>>>>>>>>>>>>>>>>>'r'n** 
#メール全件展開Dir *.msg l Sort-Object -property LastWriteTime -Descending [ ForEach{ 
Write-Host $_.Name Start-Process $_.FullName 
Dir *LC *-include*.pdflSort-Object-propertyLastWriteTime -Descendinglforeach{ Write-Host $_.NameStart-Process $_-FullName 
Dir *SDS*.:rnclude *.pdfj Sort-Object -property LastWriteTime-Descendinglforeach{ 
Write-Host $_.Name Start-Process $_-FulIName 
Write-Host**'r'n印刷が完了したら、Enterを押してください。" pause 

#印刷済みフォルダの作成、投入 
Write-Host ******************************'r'n*******$sfjt>()23V\*******:'r'n****************************ll 
Start-Sleep -s l 
New-ItemAllprint complete -ItemType DirectoryDir *.pdf I Move-Item -Destination $PrintComplete Dir *.msg [ Move-Item -Destination $PrintComplete 
#$Kariファイルに入れたものをもとにもどす Set-Location $KariDir I Move-Item -Destination $PrintComplete Set-Location $folder 
##朱印刷分があった場合(複数のAWBがあった場合) ################### 
#ALLPLINT以外のフォルダがあった場合#→再度ダイアログを開く。 孝印刷しなければならないフォルダがあれぼ印刷し、なければキャンセルを押せば下のジネーム処理に入る ##ALLPLINT以外のフォルダがなかった場合#→下のジネーム処理に入る 
#名称変更処理 
$foldername - (Split-Path . -leaf) Set-Location "C:Y";Stop-Process -Name AcroRd32 -Force Stop-Process -Name RdrCEF -Force Sleep -s 3 
Remove-Item $Kari 
Rename-Item "$folder" -NewName "【印刷完了】 $foldername 
Write-Host *1**********************************\ 
Sleep -s 2 
COMPLETE! ! ! 
