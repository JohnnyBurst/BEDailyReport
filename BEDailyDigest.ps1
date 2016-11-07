#Email report of error jobs sorted by Time ended

#Importing Backup Exec Powershell awesomeness
$Module = Get-Module bemcli

If(!$Module){
    Import-Module bemcli
}

#last 24 hours
$LastDay = (Get-Date).AddDays(-1)

###########Lets make some variables##################################
$SmtpServer = 'smtp.mailserver.com'
$From = 'emailaddy@mailserver.com'
$To = 'emailaddy@mailserver.com'
$Subject = 'Backup Exec Job digest'

#Authentication# - Not mandatory
$Username = ''
$Password = ''

#Settings#
$SSL = $false
$Port = '25'
###########Thats enough variables####################################

Function SendEmailStatus($From, $To, $Subject, $SmtpServer, $Username, $Password, $BodyAsHtml, $Body)
{	$SmtpMessage = New-Object System.Net.Mail.MailMessage $From, $To, $Subject, $Body
	$SmtpMessage.IsBodyHTML = $BodyAsHtml
	$SmtpClient = New-Object System.Net.Mail.SmtpClient($SmtpServer, $Port)
    
    If($SSL){
        $SmtpClient.EnableSSL = $true
    }

    If($Username -and $Password){
        $SmtpClient.Credentials = New-Object System.Net.NetworkCredential($Username, $Password)
    }

	$SmtpClient.Send($SmtpMessage)
	$SmtpMessage.Dispose()
}

$Style = "<Style>BODY{font-size:12px;font-family:verdana,sans-serif;color:black;font-weight:normal;}" + `
"TABLE{width:100%;border-width:1px;cellpadding:0;cellspacing:0;border-style:solid;border-color:black;border-collapse:collapse;}" + `
"TH{background:#d3d3d3;font-size:12px;border-width:1px;padding:10px;border-style:solid;border-color:black;}" + `
"TR{font-size:12px;border-width:1px;padding:10px;border-style:solid;border-color:black;}" + `
"TD{width:15%;font-size:10px;border-width:1px;padding:4px;border-style:solid;border-color:black;}</Style>"

$Table1 = get-bejob | Get-BEJobHistory -FromStartTime $lastday -jobstatus Error | sort Name | select Name,JobType,JobStatus,@{name='Size (GB)';expression={$_.TotalDataSizebytes/1073741824}},ErrorMessage | Convertto-html -fragment
$Table2 = get-bejob | Get-BEJobHistory -FromStartTime $lastday -jobstatus Canceled | sort Name | select Name,JobType,JobStatus,@{name='Size (GB)';expression={$_.TotalDataSizebytes/1073741824}},ErrorMessage | Convertto-html -fragment
$Table3 = get-bejob -jobtype backup | Get-BEJobHistory -FromStartTime $lastday -jobstatus Succeeded | sort Name | select Name,JobType,JobStatus,@{name='Size (GB)';expression={$_.TotalDataSizebytes/1073741824 -as 'Int'}} | Convertto-html -fragment
$Table4 = get-bejob -jobtype backup | Get-BEJobHistory -FromStartTime $lastday -jobstatus SucceededWithExceptions | sort Name | select Name,JobType,JobStatus,@{name='Size (GB)';expression={$_.TotalDataSizebytes/1073741824 -as 'Int'}} | Convertto-html -fragment

$TablesHead = "<html><head>$Style</head>"
$TablesBody = "<body><table><TR><TD align=center bgcolor=RED><font color=WHITE><B>JOBS WITH ERRORS</B></font></TD></TR></table>$Table1$Table2 `n<table><TR><TD align=center bgcolor=Green><font color=WHITE><B>SUCCESSES</B></font></TD></TR></table>$Table3<table><TR><TD align=center bgcolor=Orange><font color=WHITE><B>SUCCESSES WITH EXCEPTIONS</B></font></TD></TR></table>$Table4</body>"
$TablesFoot = "</html>"
$email = $TablesHead + $TablesBody + $TablesFoot

SendEmailStatus -From $From -To $To -Subject $Subject -SmtpServer $SmtpServer -Username $Username -Password $Password -BodyAsHtml $True -Body ($email)
