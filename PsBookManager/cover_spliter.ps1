Param([String]$imgPath)
if(($imgPath -eq $null) -or ($imgPath -eq "")) {
return
}

[System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") | Out-Null
[System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") | Out-Null
 
$winform = New-Object Windows.Forms.Form
$winform.Text = "Image"
#$winform.Size = New-Object Drawing.Size(1200,1000)
$winform.WindowState = [System.windows.forms.FormWindowState]::Maximized
$winform.AutoScroll = $true
#$winform.DoubleBuffered = $true

$labelnotice = New-Object System.Windows.Forms.Label
$labelnotice.Text = "線の部分は右または下の領域に含まれる"
$labelnotice.AutoSize = $True
$winform.Controls.Add($labelnotice)

$img = [System.Drawing.Image]::FromFile($imgPath)

$imgwidth = $img.Width
$imgheight = $img.Height
# 左右スライダ用
$heightupper = 120
$heightlower = 120
# 操作ボタン等 + 上下スライダ用
$widthleft = 200
# 上下スライダ用
$widthright = 80
# トータル
$widthall = $widthleft + $imgwidth + $widthright
$heightall = $heightupper + $imgheight + $heightlower

$panel = New-Object Windows.Forms.Panel
#$panelimg.AutoSize = $true
$panel.Size = New-Object Drawing.Size($widthall, $heightall)
#$panel.BackColor = [System.Drawing.Color]::Transparent
$winform.Controls.Add($panel)

#$picbox = New-Object Windows.Forms.PictureBox
#$picbox.Image = $img
#$picbox.Location = New-Object Drawing.Point($widthleft, $heightupper)
##$picbox.SizeMode = [System.Windows.Forms.PictureBoxSizeMode]::AutoSize;
#$picbox.Size = New-Object Drawing.Size(($imgwidth + 1), ($imgheight + 1))
#$winform.Controls.Add($picbox)

$vscr_handlewidth = 18
$vscr_handleheight = 18
$vscr_markerheight = 16

$hscr_handlewidth = 18
$hscr_handleheight = 18
$hscr_markerwidth = 16

function makeVscrollBar() {
param([int]$val, [int]$width, [int]$height, [int]$offsetleft, [int]$offsettop)
$vscroll = New-Object Windows.Forms.VScrollBar
$vscroll.ClientSize = New-Object Drawing.Size($width, $height)
$vscroll.Minimum = 0
$vscroll.Maximum = ($imgheight + 1)
$vscroll.Value = $val
$vscroll.LargeChange = 1
$vscroll.Location = New-Object Drawing.Point($offsetleft, $offsettop)
return $vscroll
}

function makeHscrollBar() {
param([int]$val, [int]$width, [int]$height, [int]$offsetleft, [int]$offsettop)
$hscroll = New-Object Windows.Forms.HScrollBar
$hscroll.ClientSize = New-Object Drawing.Size($width, $height)
$hscroll.Minimum = 0
$hscroll.Maximum = ($imgwidth + 1)
$hscroll.Value = $val
$hscroll.LargeChange = 1
$hscroll.Location = New-Object Drawing.Point($offsetleft, $offsettop)
return $hscroll
}

$vscr_width = $vscr_handlewidth
$vscr_height = ($imgheight + $vscr_handleheight * 2 + $vscr_markerheight + 1)

$vscr_offsettop = ($heightupper - $vscr_handleheight)
$vscr_offsetleft = ($widthleft - ($hscr_handlewidth + $vscr_handlewidth * 2 + 1))
$vscrolltopleft = makeVscrollBar 0 $vscr_width $vscr_height $vscr_offsetleft $vscr_offsettop
$panel.Controls.Add($vscrolltopleft)

$vscr_offsetleft = ($widthleft + $imgwidth + $hscr_handlewidth + $hscr_markerwidth + $vscr_handlewidth * 0)
$vscrolltopright = makeVscrollBar 0 $vscr_width $vscr_height $vscr_offsetleft $vscr_offsettop
$panel.Controls.Add($vscrolltopright)

$vscr_offsetleft = ($widthleft - ($hscr_handlewidth + $vscr_handlewidth * 1 + 1))
$vscrollbottomleft = makeVscrollBar ($imgheight + 1) $vscr_width $vscr_height $vscr_offsetleft $vscr_offsettop
$panel.Controls.Add($vscrollbottomleft)

$vscr_offsetleft = ($widthleft + $imgwidth + $hscr_handlewidth + $hscr_markerwidth + $vscr_handlewidth * 1)
$vscrollbottomright = makeVscrollBar ($imgheight + 1) $vscr_width $vscr_height $vscr_offsetleft $vscr_offsettop
$panel.Controls.Add($vscrollbottomright)

$hscr_width = ($imgwidth + $hscr_handlewidth * 2 + $hscr_markerwidth + 1)
$hscr_height = $hscr_handleheight

$hscr_offsetleft = $widthleft - ($vscr_handlewidth + 1)
$hscr_initvalue = 0
$hscr_offsettop = $heightupper - ($vscr_handleheight * 6)
$hscrolledgelefttop = makeHscrollBar $hscr_initvalue $hscr_width $hscr_height $hscr_offsetleft $hscr_offsettop
$panel.Controls.Add($hscrolledgelefttop)

$hscr_offsettop = ($heightupper + $imgheight + 1 + $vscr_handleheight * 0)
$hscrolledgeleftbottom = makeHscrollBar $hscr_initvalue $hscr_width $hscr_height $hscr_offsetleft $hscr_offsettop
$panel.Controls.Add($hscrolledgeleftbottom)

$hscr_initvalue = ($imgwidth * 0.16)
$hscr_offsettop = $heightupper - ($vscr_handleheight * 5)
$hscrollsleevelefttop = makeHscrollBar $hscr_initvalue $hscr_width $hscr_height $hscr_offsetleft $hscr_offsettop
$panel.Controls.Add($hscrollsleevelefttop)

$hscr_offsettop = ($heightupper + $imgheight + 1 + $vscr_handleheight * 1)
$hscrollsleeveleftbottom = makeHscrollBar $hscr_initvalue $hscr_width $hscr_height $hscr_offsetleft $hscr_offsettop
$panel.Controls.Add($hscrollsleeveleftbottom)

$hscr_initvalue = ($imgwidth * 0.48)
$hscr_offsettop = $heightupper - ($vscr_handleheight * 4)
$hscrollcoverlefttop = makeHscrollBar $hscr_initvalue $hscr_width $hscr_height $hscr_offsetleft $hscr_offsettop
$panel.Controls.Add($hscrollcoverlefttop)

$hscr_offsettop = ($heightupper + $imgheight + 1 + $vscr_handleheight * 2)
$hscrollcoverleftbottom = makeHscrollBar $hscr_initvalue $hscr_width $hscr_height $hscr_offsetleft $hscr_offsettop
$panel.Controls.Add($hscrollcoverleftbottom)

$hscr_initvalue = ($imgwidth * 0.52)
$hscr_offsettop = $heightupper - ($vscr_handleheight * 3)
$hscrollbackfacetop = makeHscrollBar $hscr_initvalue $hscr_width $hscr_height $hscr_offsetleft $hscr_offsettop
$panel.Controls.Add($hscrollbackfacetop)

$hscr_offsettop = ($heightupper + $imgheight + 1 + $vscr_handleheight * 3)
$hscrollbackfacebottom = makeHscrollBar $hscr_initvalue $hscr_width $hscr_height $hscr_offsetleft $hscr_offsettop
$panel.Controls.Add($hscrollbackfacebottom)

$hscr_initvalue = ($imgwidth * 0.84)
$hscr_offsettop = $heightupper - ($vscr_handleheight * 2)
$hscrollcoverrighttop = makeHscrollBar $hscr_initvalue $hscr_width $hscr_height $hscr_offsetleft $hscr_offsettop
$panel.Controls.Add($hscrollcoverrighttop)

$hscr_offsettop = ($heightupper + $imgheight + 1 + $vscr_handleheight * 4)
$hscrollcoverrightbottom = makeHscrollBar $hscr_initvalue $hscr_width $hscr_height $hscr_offsetleft $hscr_offsettop
$panel.Controls.Add($hscrollcoverrightbottom)

$hscr_initvalue = $imgwidth
$hscr_offsettop = $heightupper - ($vscr_handleheight * 1)
$hscrollsleeverighttop = makeHscrollBar $hscr_initvalue $hscr_width $hscr_height $hscr_offsetleft $hscr_offsettop
$panel.Controls.Add($hscrollsleeverighttop)

$hscr_offsettop = ($heightupper + $imgheight + 1 + $vscr_handleheight * 5)
$hscrollsleeverightbottom = makeHscrollBar $hscr_initvalue $hscr_width $hscr_height $hscr_offsetleft $hscr_offsettop
$panel.Controls.Add($hscrollsleeverightbottom)

$event_vscrolltopchanged = $false
$event_vscrollbottomchanged = $false

$vscrolltopleft.Add_ValueChanged({
if($event_vscrolltopchanged) {
$event_vscrolltopchanged = $false
} else {
$event_vscrolltopchanged = $true
$vscrolltopright.Value = $vscrolltopleft.Value
$panel.Invalidate()
}
})
$vscrollbottomleft.Add_ValueChanged({
if($event_vscrolltopchanged) {
$event_vscrolltopchanged = $false
} else {
$event_vscrolltopchanged = $true
$vscrollbottomright.Value = $vscrollbottomleft.Value
$panel.Invalidate()
}
})
$vscrolltopright.Add_ValueChanged({
if($event_vscrolltopchanged) {
$event_vscrolltopchanged = $false
} else {
$event_vscrolltopchanged = $true
$vscrolltopleft.Value = $vscrolltopright.Value
$panel.Invalidate()
}
})
$vscrollbottomright.Add_ValueChanged({
if($event_vscrolltopchanged) {
$event_vscrolltopchanged = $false
} else {
$event_vscrolltopchanged = $true
$vscrollbottomleft.Value = $vscrollbottomright.Value
$panel.Invalidate()
}
})

$event_hscrolledgeleftchanged = $false
$event_hscrollsleeveleftchanged = $false
$event_hscrollcoverleftchanged = $false
$event_hscrollbackfacechanged = $false
$event_hscrollcoverrightchanged = $false
$event_hscrollsleeverightchanged = $false

$hscrolledgelefttop.Add_ValueChanged({
if($event_hscrolledgeleftchanged) {
$event_hscrolledgeleftchanged = $false
} else {
$event_hscrolledgeleftchanged = $true
$hscrolledgeleftbottom.Value = $hscrolledgelefttop.Value
$panel.Invalidate()
}
})

$hscrolledgeleftbottom.Add_ValueChanged({
if($event_hscrolledgeleftchanged) {
$event_hscrolledgeleftchanged = $false
} else {
$event_hscrolledgeleftchanged = $true
$hscrolledgelefttop.Value = $hscrolledgeleftbottom.Value
$panel.Invalidate()
}
})

$hscrollsleevelefttop.Add_ValueChanged({
if($event_hscrollsleeveleftchanged) {
$event_hscrollsleeveleftchanged = $false
} else {
$event_hscrollsleeveleftchanged = $true
$hscrollsleeveleftbottom.Value = $hscrollsleevelefttop.Value
$panel.Invalidate()
}
})

$hscrollsleeveleftbottom.Add_ValueChanged({
if($event_hscrollsleeveleftchanged) {
$event_hscrollsleeveleftchanged = $false
} else {
$event_hscrollsleeveleftchanged = $true
$hscrollsleevelefttop.Value = $hscrollsleeveleftbottom.Value
$panel.Invalidate()
}
})

$hscrollcoverlefttop.Add_ValueChanged({
if($event_hscrollcoverleftchanged) {
$event_hscrollcoverleftchanged = $false
} else {
$event_hscrollcoverleftchanged = $true
$hscrollcoverleftbottom.Value = $hscrollcoverlefttop.Value
$panel.Invalidate()
}
})

$hscrollcoverleftbottom.Add_ValueChanged({
if($event_hscrollcoverleftchanged) {
$event_hscrollcoverleftchanged = $false
} else {
$event_hscrollcoverleftchanged = $true
$hscrollcoverlefttop.Value = $hscrollcoverleftbottom.Value
$panel.Invalidate()
}
})

$hscrollbackfacetop.Add_ValueChanged({
if($event_hscrollbackfacechanged) {
$event_hscrollbackfacechanged = $false
} else {
$event_hscrollbackfacechanged = $true
$hscrollbackfacebottom.Value = $hscrollbackfacetop.Value
$panel.Invalidate()
}
})

$hscrollbackfacebottom.Add_ValueChanged({
if($event_hscrollbackfacechanged) {
$event_hscrollbackfacechanged = $false
} else {
$event_hscrollbackfacechanged = $true
$hscrollbackfacetop.Value = $hscrollbackfacebottom.Value
$panel.Invalidate()
}
})

$hscrollcoverrighttop.Add_ValueChanged({
if($event_hscrollcoverrightchanged) {
$event_hscrollcoverrightchanged = $false
} else {
$event_hscrollcoverrightchanged = $true
$hscrollcoverrightbottom.Value = $hscrollcoverrighttop.Value
$panel.Invalidate()
}
})

$hscrollcoverrightbottom.Add_ValueChanged({
if($event_hscrollcoverrightchanged) {
$event_hscrollcoverrightchanged = $false
} else {
$event_hscrollcoverrightchanged = $true
$hscrollcoverrighttop.Value = $hscrollcoverrightbottom.Value
$panel.Invalidate()
}
})

$hscrollsleeverighttop.Add_ValueChanged({
if($event_hscrollsleeverightchanged) {
$event_hscrollsleeverightchanged = $false
} else {
$event_hscrollsleeverightchanged = $true
$hscrollsleeverightbottom.Value = $hscrollsleeverighttop.Value
$panel.Invalidate()
}
})

$hscrollsleeverightbottom.Add_ValueChanged({
if($event_hscrollsleeverightchanged) {
$event_hscrollsleeverightchanged = $false
} else {
$event_hscrollsleeverightchanged = $true
$hscrollsleeverighttop.Value = $hscrollsleeverightbottom.Value
$panel.Invalidate()
}
})

$panel.Add_Paint({
$_.Graphics.DrawImage($img, $widthleft, $heightupper, $imgwidth, $imgheight)
$curr_topleft = New-Object Drawing.Point($widthleft, ($heightupper + $vscrolltopleft.Value))
$curr_topright = New-Object Drawing.Point(($widthleft + $imgwidth), ($heightupper + $vscrolltopleft.Value))
$curr_bottomleft = New-Object Drawing.Point($widthleft, ($heightupper + $vscrollbottomleft.Value))
$curr_bottomright = New-Object Drawing.Point(($widthleft + $imgwidth), ($heightupper + $vscrollbottomleft.Value))
$_.Graphics.DrawLine([System.Drawing.Pens]::Red, $curr_topleft, $curr_topright)
$_.Graphics.DrawLine([System.Drawing.Pens]::Blue, $curr_bottomleft, $curr_bottomright)
$curr_edgelefttop = New-Object Drawing.Point(($widthleft + $hscrolledgelefttop.Value - 1), $heightupper)
$curr_edgeleftbottom = New-Object Drawing.Point(($widthleft + $hscrolledgelefttop.Value - 1), ($heightupper + $imgheight))
$_.Graphics.DrawLine([System.Drawing.Pens]::Red, $curr_edgelefttop, $curr_edgeleftbottom)
$curr_sleevelefttop = New-Object Drawing.Point(($widthleft + $hscrollsleevelefttop.Value - 1), $heightupper)
$curr_sleeveleftbottom = New-Object Drawing.Point(($widthleft + $hscrollsleevelefttop.Value - 1), ($heightupper + $imgheight))
$_.Graphics.DrawLine([System.Drawing.Pens]::Blue, $curr_sleevelefttop, $curr_sleeveleftbottom)
$curr_coverlefttop = New-Object Drawing.Point(($widthleft + $hscrollcoverlefttop.Value - 1), $heightupper)
$curr_coverleftbottom = New-Object Drawing.Point(($widthleft + $hscrollcoverlefttop.Value - 1), ($heightupper + $imgheight))
$_.Graphics.DrawLine([System.Drawing.Pens]::Green, $curr_coverlefttop, $curr_coverleftbottom)
$curr_backfacetop = New-Object Drawing.Point(($widthleft + $hscrollbackfacetop.Value - 1), $heightupper)
$curr_backfacebottom = New-Object Drawing.Point(($widthleft + $hscrollbackfacetop.Value - 1), ($heightupper + $imgheight))
$_.Graphics.DrawLine([System.Drawing.Pens]::Orange, $curr_backfacetop, $curr_backfacebottom)
$curr_coverrighttop = New-Object Drawing.Point(($widthleft + $hscrollcoverrighttop.Value - 1), $heightupper)
$curr_coverrightbottom = New-Object Drawing.Point(($widthleft + $hscrollcoverrighttop.Value - 1), ($heightupper + $imgheight))
$_.Graphics.DrawLine([System.Drawing.Pens]::Blue, $curr_coverrighttop, $curr_coverrightbottom)
$curr_sleeverighttop = New-Object Drawing.Point(($widthleft + $hscrollsleeverighttop.Value - 1), $heightupper)
$curr_sleeverightbottom = New-Object Drawing.Point(($widthleft + $hscrollsleeverighttop.Value - 1), ($heightupper + $imgheight))
$_.Graphics.DrawLine([System.Drawing.Pens]::Red, $curr_sleeverighttop, $curr_sleeverightbottom)
})

$buttonClose = New-Object Windows.Forms.Button
$buttonClose.Text = "Close"
$buttonClose.Add_Click({
$winform.Close()
})
$buttonClose.Location = New-Object Drawing.Point(10, ($heightupper + $imgheight))
$panel.Controls.Add($buttonClose)

function saveImage() {
param([int]$left, [int]$top, [int]$width, [int]$height, [String]$opath)
if($width -le 0) {
    return
}
$rect = New-Object Drawing.Rectangle( $left, $top, $width, $height)
$imgBmp = New-Object Drawing.Bitmap($img)
$bmp = $imgBmp.Clone($rect, $imgBmp.PixelFormat)
if(Test-Path $opath) {
    Remove-Item $opath
}
$bmp.save($opath)
}

$buttonExec = New-Object Windows.Forms.Button
$buttonExec.Text = "Exec"
$buttonExec.Add_Click({
#Write-Host "Top:" $vscrolltopleft.Value ", Bottom:" $vscrollbottomleft.Value
#Write-Host "Edge:" $hscrolledgelefttop.Value ", SleeveLeft:" $hscrollsleevelefttop.Value ", CoverLeft:" $hscrollcoverlefttop.Value
#Write-Host "BackFace: " $hscrollbackfacetop.Value ", CoverRight:" $hscrollcoverrighttop.Value ", SleeveRight:" $hscrollsleeverighttop.Value
$top = $vscrolltopleft.Value
$height = ($vscrollbottomleft.Value - $top + 1)
$bpath = (Split-Path (Split-Path $imgPath -Parent) -Parent)

$left = $hscrolledgelefttop.Value + 1
$width = ($hscrollsleevelefttop.Value - $left + 1)
$opath = $bpath + "\C03.png"
saveImage $left $top $width $height $opath

$left = $hscrollsleevelefttop.Value + 1
$width = ($hscrollcoverlefttop.Value - $left + 1)
$opath = $bpath + "\C01.png"
saveImage $left $top $width $height $opath

$left = $hscrollcoverlefttop.Value + 1
$width = ($hscrollbackfacetop.Value - $left + 1)
$opath = $bpath + "\B00.png"
saveImage $left $top $width $height $opath

$left = $hscrollbackfacetop.Value + 1
$width = ($hscrollcoverrighttop.Value - $left + 1)
$opath = $bpath + "\C02.png"
saveImage $left $top $width $height $opath

$left = $hscrollcoverrighttop.Value + 1
$width = ($hscrollsleeverighttop.Value - $left + 1)
$opath = $bpath + "\C04.png"
saveImage $left $top $width $height $opath
})
$buttonExec.Location = New-Object Drawing.Point(10, ($heightupper + $imgheight + 30))
$panel.Controls.Add($buttonExec)

$winform.Add_Shown({$winform.Activate()})
$winform.ShowDialog() | Out-Null
