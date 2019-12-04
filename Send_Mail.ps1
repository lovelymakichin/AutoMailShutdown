[System.String]$fromAdrName = $FromFullName +  "<" + $FromAddress + ">"

# メールアドレス（受信者）
[System.String]$toAdrName = $ToFullName +  "<" + $ToAddress + ">"

# メールアドレス（BCC受信者）
[System.String]$bccAdrname = $fromAdrName

# メールサーバ
[System.String]$serverAddress = $smtpServer

# 日付
[System.String]$today = (Get-Date).ToString("M/dd")

# 件名
[System.String]$subject = $today + "   "+ $FromlastName + "出社状況"

#弁当有無
[String[]]$array = @("あり","なし")

$i = (Read-Host お弁当は？0:あり 1:なし)

#本文
[System.String]$body = $TolastName + "さん

お疲れ様です。
" + $today +"出社状況は、下記の通りです。


勤務状況：出勤
昼食　　："+ $array[$i] + "


宜しくお願い致します。

" + $FromlastName

Write-Host "--------------------------以下メール内容----------------------------------"
Write-Host $body
Write-Host "-------------------------以上メール内容----------------------------------"

Write-Host "上記内容でメールを送信します。キャンセルの場合はCtrl+Cを押して下さい。"

pause

# コマンド実行
send-mailmessage -from $fromAdrName -to $toAdrName -bcc $bccAdrname -SmtpServer $serverAddress -Encoding $encoding -subject $subject -Body $body

#タイムアウト
timeout 3

#シャットダウン
Stop-Computer

