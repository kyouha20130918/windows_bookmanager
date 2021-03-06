[void][Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
#[void][Reflection.Assembly]::LoadWithPartialName("System.ComponentModel.ListSortDirection")

# url : http://memo-space.blogspot.jp/2010/02/powershellwin32api.html
# PowerShellからのPlatform APIの呼び出しは、
# .NET Frameworkの提供するP/Invoke機能を介して実施することができる。
# Invoke-Win32では、P/Invokeを実行用に一時的な型を定義し、
# そこに実装したStaticなメソッドを介してPlatformAPIを実行する。
function Invoke-Win32()
{
    param
    ( [string]$dllName, [Type]$returnType, [string]$methodName, [Object[]]$parameterInfos )
   
   $parameterTypes = $parameterInfos | %{ $_[ 0 ] }
   $parameters = $parameterInfos | %{ $_[ 1 ] }

   # カレントのアプリケーションドメインに、
   # P/Invokeを実行するメソッドを持つ独自の型を定義する。
   $domain = [AppDomain]::CurrentDomain
   $name = New-Object Reflection.AssemblyName 'PInvokeAssembly'
   $assembly = $domain.DefineDynamicAssembly( $name, 'Run' )
   $module = $assembly.DefineDynamicModule( 'PInvokeModule')
   $type = $module.DefineType( 'PInvokeType', "Public,BeforeFieldInit" )

   # P/Invokeを呼び出すためのパラメーターを準備する。
   # PlatformAPIに渡すパラメーターを保持する配列
   $inputParameters = @() 
   # PSReference型パラメーターの位置を保持する配列
   $refParameters = @()
  
   for( $counter = 0; $counter -lt $parameterTypes.Length; $counter++ )
   {
      # パラメーターの型が「PSReference」のものについては、関数から戻ってきた値を拾えるようにする必要がある。
      if( $parameterTypes[ $counter ] -eq [Ref] )
      {
         # 呼び出しの際に[out]をつける必要があるため、そのパラメーターの位置を退避しておく。
         # 配列の数より1つ大きな数を保持しておく必要がある。（理由は後述）
         $refParameters += $counter

         # PSReference型を、.NETオブジェクトの参照型に書き換える。
         $parameterTypes[ $counter ] = $parameters[ $counter ].Value.GetType().MakeByRefType()
         
         # 関数呼び出し時に使用するパラメーター一覧に値を追加する
         $inputParameters += $parameters[ $counter ].Value
      }
      # そうでないものについては、関数呼び出し時に使用するパラメーター一覧にただ追加するのみ
      else
      {
         $inputParameters += $parameters[ $counter ]
      }
   }

   # アセンブリーの動的メソッドとして、PlatformAPIを定義する
   $method = $type.DefineMethod( $methodName, `
                                 'Public,HideBySig,Static,PinvokeImpl', `
                                 $returnType, `
                                 $parameterTypes )
   # PSReference型のパラメーターは、out属性のパラメーターとする。
   foreach( $refParameter in $refParameters )
   {
       # 0番目の要素は戻り値の情報になるため、配列の要素番号+1を指定する必要がある。
      $method.DefineParameter( ( $refParameter + 1 ), "Out", $null )
   }

   # P/Invokeのコンストラクターを設定する
   $ctor = [Runtime.InteropServices.DllImportAttribute].GetConstructor( [string] )
   $attr = New-Object Reflection.Emit.CustomAttributeBuilder $ctor, $dllName
   $method.SetCustomAttribute( $attr )
    
   # 一時的な型を作成し、そのメソッドとしてPlatformAPIを実行する。
   $realType = $type.CreateType()
   $returnObject = $realType.InvokeMember( $methodName, `
                                           'Public,Static,InvokeMethod', `
                                           $null, `
                                           $null, `
                                           $inputParameters )
    
   # PSReference型で受け取ったパラメーターの値を更新する。
   foreach( $refParameter in $refParameters )
   {
      $parameterInfos[ $refParameter ][ 1 ].Value = $inputParameters[ $refParameter ]
   }
   # PlatformAPIの戻り値を返す
   return $returnObject
}

function getIniValue {
param([System.Object[]]$prm)
Invoke-Win32 -dllName "kernel32.dll" -returnType ( [UInt32] ) -methodName "GetPrivateProfileString" -parameterInfos $prm | out-null
}

function writeIniValue {
param([System.Object[]]$prm)
Invoke-Win32 -dllName "kernel32.dll" -returnType ( [UInt32] ) -methodName "WritePrivateProfileString" -parameterInfos $prm | out-null
}

function makeGetParamInfo {
param([System.String]$section, [System.String]$key, [System.String]$default, [System.Text.StringBuilder]$ret, [System.String]$fpath)
return @( @( [string], [string]$section ), `
          @( [string], [string]$key ), `
          @( [string], [string]$default ), `
          @( [System.Text.StringBuilder], [System.Text.StringBuilder]$ret ), `
          @( [int], [int]$ret.Capacity ), `
          @( [string], [string]$fpath ) )
}

function makeWriteParamInfo {
param([System.String]$section, [System.String]$key, [System.String]$val, [System.String]$fpath)
return @( @( [string], [string]$section ), `
          @( [string], [string]$key ), `
          @( [string], [string]$val ), `
          @( [string], [string]$fpath ) )
}

$tag_fpath="fpath"
$tag_title="title"
$tag_author="author"
$tag_publisher="publisher"
$tag_pubdate="pubdate"

function getDesktopIniValues {
param([string]$fpath)
$vals = @{}

$valTitle = New-Object System.Text.StringBuilder 500
$paramTitle = makeGetParamInfo ".ShellClassInfo" "LocalizedResourceName" "" $valTitle $fpath
getIniValue $paramTitle
$vals[$tag_title] = $valTitle.toString()

$valAuthor = New-Object System.Text.StringBuilder 500
$paramAuthor = makeGetParamInfo "{64440492-4C8B-11D1-8B70-080036B11A03}" "Prop11" "" $valAuthor $fpath
getIniValue $paramAuthor
$vals[$tag_author] = $valAuthor.toString().split(",")[1]

$valPublisher = New-Object System.Text.StringBuilder 500
$paramPublisher = makeGetParamInfo "{64440492-4C8B-11D1-8B70-080036B11A03}" "Prop30" "" $valPublisher $fpath
getIniValue $paramPublisher
$vals[$tag_publisher] = $valPublisher.toString().split(",")[1]

$valPubdate = New-Object System.Text.StringBuilder 500
$paramPubdate = makeGetParamInfo "{DE41CC29-6971-4290-B472-F59F2E2F31E2}" "Prop100" "" $valPubdate $fpath
getIniValue $paramPubdate
$vals[$tag_pubdate] = $valPubdate.toString().split(",")[1]

$vals[$tag_fpath] = $fpath
return $vals
}

function getBookInfoJsonFname {
param([System.Collections.Hashtable]$vals)
$fpath = $vals[$tag_fpath]
$jpath = (Split-Path $fpath -Parent) + "\bookinfo.json"
return $jpath
}

function getShelfInfoJsonFname {
param([System.Collections.Hashtable]$vals)
$fpath = $vals[$tag_fpath]
$npath = (Split-Path $fpath -Parent) + "\bookshelf.json"
return $npath
}

function getBookNoteJsonFname {
param([string]$bookinfo)
$npath = (Split-Path $bookinfo -Parent) + "\booknote.json"
return $npath
}

function writeDesktopIniValues {
param([System.Collections.Hashtable]$vals)
$fpath = $vals[$tag_fpath]
if(Test-Path $fpath) {
Clear-ItemProperty -path $fpath -name Attributes
}

$paramTitle = makeWriteParamInfo ".ShellClassInfo" "LocalizedResourceName" $vals[$tag_title] $fpath
writeIniValue $paramTitle

$paramAuthor = makeWriteParamInfo "{64440492-4C8B-11D1-8B70-080036B11A03}" "Prop11" ("31," + $vals[$tag_author]) $fpath
writeIniValue $paramAuthor

$paramPublisher = makeWriteParamInfo "{64440492-4C8B-11D1-8B70-080036B11A03}" "Prop30" ("31," + $vals[$tag_publisher]) $fpath
writeIniValue $paramPublisher

$paramPubdate = makeWriteParamInfo "{DE41CC29-6971-4290-B472-F59F2E2F31E2}" "Prop100" ("31," + $vals[$tag_pubdate]) $fpath
writeIniValue $paramPubdate

Set-ItemProperty -path $fpath -name Attributes -value "Hidden,System"

$folname = (Split-Path (Split-Path $fpath -Parent) -Leaf)
$title = (nullToBlank $vals[$tag_title])
$author = (nullToBlank $vals[$tag_author])
if($isShelf[$folname]) {
if(($title -ne "") -and ($author -ne "")) {
$jpath = getShelfInfoJsonFname $vals
$shelfinfo = @{}
$shelfinfo.add("name",$title)
$shelfinfo.add("itemType",$author)
if(Test-Path $jpath) {
Remove-Item $jpath
}
ConvertTo-Json $shelfinfo | %{$_ -replace " {20}","        "} | %{$_ -replace " {16}","    "} | %{$_ -replace ":  ",":"} | %{$_ -replace " {4}","  "} | New-Item $jpath -ItemType File
}
} else {
$jpath = getBookInfoJsonFname $vals
$jid = (Split-Path (Split-Path $fpath -Parent) -Leaf)
$jsoncolophon = @{}
$jsoncolophon.Add("id", $jid)
$jsoncolophon.add("title", $vals[$tag_title])
$jsoncolophon.add("author", $vals[$tag_author])
$jsoncolophon.add("publisher", $vals[$tag_publisher])
$jsoncolophon.add("pubdate", $vals[$tag_pubdate])
$jsonval = @{"colophon"=$jsoncolophon}
if(Test-Path $jpath) {
Remove-Item $jpath
}
ConvertTo-Json $jsonval | %{$_ -replace " {20}","        "} | %{$_ -replace " {16}","    "} | %{$_ -replace ":  ",":"} | %{$_ -replace " {4}","  "} | New-Item $jpath -ItemType File

$npath = getBookNoteJsonFname $jpath
if(-not (Test-Path $npath)) {
New-Item -ItemType file $npath
}
}
}

function nullToBlank {
param([System.String]$org)
if($org -ne $null) {
return $org
}
return ""
}

### main

$rootdir = "C:\home\docs\books"
$pairs = @{}
$dirs = (dir -directory $rootdir -Recurse | ? {$_.name.Length -gt 2} | Sort-Object fullname)
$dirs | % {$pairs[$_]=get-childitem $_.fullname -force | ? {$_.name -eq "desktop.ini"}}

$iniVals = @{}
$pairs.keys | % {if($pairs[$_] -ne $null){
#get-content $pair[$_].fullname
$vals = getDesktopIniValues $pairs[$_].fullname
} else {
$vals = @{}
$vals[$tag_fpath] = ($_.fullname + "\desktop.ini")
}
$iniVals[$_.fullname] = $vals
}

$isShelf = @{}
$dirs | % {if((dir -directory $_.fullname -Recurse | ? {$_.name.Length -gt 2}).count -gt 0) {$isShelf[$_.name] = $true;}}

#$iniVals.keys | % {$iniVals[$_]}

#$dirs | % {$k = $_.name; $s = ""; if($iniVals[$k] -ne $null) { $a = $iniVals[$k]; $s = "`tfile::" + $_.fullname + "\R\C000.png`t" + $a[$tag_title] + "`t" + $a[$tag_author] + "`t" + $a[$tag_publisher] + "`t" + $a[$tag_pubdate]}; $_.name + $s} > Z:\books_props.txt

$form = New-Object System.Windows.Forms.Form
$form.Size = New-Object System.Drawing.Size(1000,1300)
$dataGridView = New-Object System.Windows.Forms.DataGridView
$dataGridView.Location=New-Object System.Drawing.Size(0,30)
$dataGridView.Size=New-Object System.Drawing.Size(1000,1230)
$dataGridView.dock=[System.Windows.Forms.DockStyle]::Bottom
$doOut = New-Object System.Windows.Forms.Button
$doOut.Location = New-Object System.Drawing.Size(100,5)
$doOut.Size = New-Object System.Drawing.Size(75,23)
$doOut.text = "&Write"
$form.Controls.Add($doOut)
$form.Controls.Add($dataGridView)

$dataGridView.ColumnCount = 7
$dataGridView.ColumnHeadersVisible = $true
$dataGridView.Columns[0].Name = "Sts"
$dataGridView.Columns[1].Name = "Dir"
$dataGridView.Columns[2].Name = "Cover"
$dataGridView.Columns[3].Name = "Title"
$dataGridView.Columns[4].Name = "Author"
$dataGridView.Columns[5].Name = "Publisher"
$dataGridView.Columns[6].Name = "PubDate"
$dataGridView.Columns[0].width = 20
$dataGridView.Columns[1].width = 200
$dataGridView.Columns[2].width = 100
$dataGridView.Columns[3].width = 300
$dataGridView.Columns[4].width = 120
$dataGridView.Columns[5].width = 100
$dataGridView.Columns[6].width = 80
$dataGridView.Columns[0].readonly = $true
$dataGridView.Columns[1].readonly = $true
$dataGridView.Columns[2].readonly = $true
$font = new-object System.Drawing.Font($dataGridView.font, [System.Drawing.FontStyle]::Underline)
$dataGridView.Columns[2].DefaultCellStyle.Font = $font
$dataGridView.AllowUserToAddRows = $false
$dataGridView.AllowUserToDeleteRows = $false
$dataGridView.EditMode = [System.Windows.Forms.DataGridViewEditMode]::EditOnEnter
#$dataGridView.Sort($dataGridView.Columns[1], [System.ComponentModel.ListSortDirection]::SortOrder.Ascending)

$dataGridView.add_CellClick({
($sender, $e) = $this, $_
if($e.ColumnIndex -eq 2) {
$cpath = $dataGridView.Rows[$e.RowIndex].Cells[2]
invoke-item $cpath.value
} elseif($e.ColumnIndex -eq 0) {
$sts = $dataGridView.Rows[$e.RowIndex].Cells[0]
if($sts.Value -eq "☆") {
$cpath = $dataGridView.Rows[$e.RowIndex].Cells[2]
&(Join-Path (Split-Path ( & { $myInvocation.ScriptName } ) -parent) "\cover_spliter.ps1") $cpath.Value
$spath = (Split-Path (Split-Path $cpath.Value -Parent) -Parent) + "\C01.png"
if(Test-Path $spath) {
$dataGridView.Rows[$e.RowIndex].Cells[0].Value = "★"
}
}
}
})
$doOut.add_Click({
$dataGridView.Rows | % {
$sts = $_.cells[0].value
$dirpath = $rootdir + "\" + $_.cells[1].value
$title = $_.cells[3].value
$author = $_.cells[4].value
$publisher = $_.cells[5].value
$pubdate = $_.cells[6].value
$a = $iniVals[$dirpath]
$changed = $false
if($title -ne (nullToBlank $a[$tag_title])) { $changed = $true; $a[$tag_title] = $title }
if($author -ne (nullToBlank $a[$tag_author])) { $changed = $true; $a[$tag_author] = $author }
if($publisher -ne (nullToBlank $a[$tag_publisher])) { $changed = $true; $a[$tag_publisher] = $publisher }
if($pubdate -ne (nullToBlank $a[$tag_pubdate])) { $changed = $true; $a[$tag_pubdate] = $pubdate }
if($sts -eq "★") {
$jsonfname = getBookInfoJsonFname $a
if(-not (Test-Path $jsonfname)) { $changed = $true }
$notefname = getBookNoteJsonFname $jsonfname
if(-not (Test-Path $notefname)) { $changed = $true }
}
if($changed) {
$dir = Get-Item (Split-Path $a[$tag_fpath] -parent)
$dir.attributes = $dir.attributes -bxor ([System.IO.FileAttributes]::ReadOnly)
writeDesktopIniValues $a
$dir.attributes = $dir.attributes -bor ([System.IO.FileAttributes]::ReadOnly)
write-host "update of" $a[$tag_fpath] ":" $a[$tag_title]
}
}
})

$dirs | % {
$a = $iniVals[$_.fullname]
$title = nullToBlank $a[$tag_title]
$author = nullToBlank $a[$tag_author]
$publisher = nullToBlank $a[$tag_publisher]
$pubdate = nullToBlank $a[$tag_pubdate]
$hasDirU = Test-Path ($_.fullname+"\U") -pathtype container
$sts = ""
$coverPath = ($_.fullname+"\U\C000.png")
if($hasDirU) {
$sts = "○"
if((Test-Path $coverPath -pathtype leaf))
{
$sts = "☆"
$splitedCoverPath = $_.fullname+"\C01.png"
if((Test-Path $splitedCoverPath -pathtype leaf)) {
$sts = "★"
}
} else {
$coverPath = ""
}
} else {
$coverPath = ""
}
$dirpath = ($_.fullname.replace(($rootdir + "\"), ""))
$dataGridView.Rows.Add($sts, $dirpath, $coverPath, $title, $author, $publisher, $pubdate) | out-null
}
[void]$form.ShowDialog()
