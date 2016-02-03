#	Exchange 2010 no longer gives you the option to simply see the size of all mailboxes from long shot. 
# this script will show mailbox count, database size, server size, top ten largest mailboxes and 
# the size of all mailboxes. (to include disconnected mailboxes.) -Alix Hoover 
#
# 12/10/2015 	Added exchange snap in so you can right click run
# 01/07/2016 	Added mailbox count break down by database
# 01/07/2016 	Added status output to see what is being done. 
# 02/02/2016 	Added Prompt to keep the window from auto closing when using the right click function. 
# 02/02/2016 	Added notice that enter can skip email function
# 02/02/2016	Added Yes or No to Show Individual boxes. (can save time for those with 500+ boxes on one server)
# 02/03/2016	Added Mailbox Cleanup before report






# add snapping to be-able to right click run
add-pssnapin Microsoft.Exchange.Management.PowerShell.E2010
Write-host " Exchange 2010 MailBox Report" -foregroundcolor Green
Write-host "Email Address to send report to (Enter to skip email Function):" -foregroundcolor magenta -nonewline
			$ReportRecipient = read-host
			
# clean up the DB's before reports
Get-MailboxDatabase | Clean-MailboxDatabase


#Variables to configure
$MailServer = "SERVERNAME"
$fileName = "/MailBox Reports/exchange2010Report"+( get-date ).ToString('MM.dd.yyyy')+".html"
$ReportSender = "MicrosoftOutlook@DOMAIN.org"
$MailSubject = ("Exchange 2010 Mailbox Report for " + $MailServer + " - " + ( get-date ).ToString('MM/dd/yyyy'))






#SendEmailFunction
Function sendEmail
{ param($smtphost,$htmlFileName)
$smtp= New-Object System.Net.Mail.SmtpClient $smtphost
$msg = New-Object System.Net.Mail.MailMessage $ReportSender, $ReportRecipient, $MailSubject, (Get-Content $htmlFileName)
$msg.isBodyhtml = $true
$smtp.send($msg)
}



# TABLE count  command
$dbcount = Get-MailboxStatistics -server $MailServer | ?{!$_.DisconnectDate}| Group-Object -Property:database | Select-Object Name, Count |Sort-Object Name 


# TABLE Dbsize command
$dbsize = Get-MailboxDatabase -Status | Select Servername, Name, Databasesize | sort name

# TABLE top10 command
$top10 = Get-MailboxStatistics -Server $MailServer | Select-Object DisplayName, ItemCount, TotalItemSize, Database, StorageLimitStatus | sort -descending TotalItemSize |Select -First 10

#TABLE all users command
$exdata = Get-MailboxStatistics -Server $MailServer | Select-Object DisplayName, ItemCount, TotalItemSize, Database, StorageLimitStatus | sort Displayname







New-Item -ItemType file $fileName -Force

# HTML start
Add-Content $fileName "<html>"

# HEAD start
Add-Content $fileName "<head>"

add-content $fileName '<STYLE TYPE="text/css">'
add-content $fileName  "<!--"
add-content $fileName  "td {"
add-content $fileName  "font-family: Tahoma;"
add-content $fileName  "font-size: 11px;"
add-content $fileName  "border-top: 1px solid #999999;"
add-content $fileName  "border-right: 1px solid #999999;"
add-content $fileName  "border-bottom: 1px solid #999999;"
add-content $fileName  "border-left: 1px solid #999999;"
add-content $fileName  "padding-top: 0px;"
add-content $fileName  "padding-right: 0px;"
add-content $fileName  "padding-bottom: 0px;"
add-content $fileName  "padding-left: 0px;"
add-content $fileName  "}"
add-content $fileName  "body {"
add-content $fileName  "margin-left: 5px;"
add-content $fileName  "margin-top: 5px;"
add-content $fileName  "margin-right: 0px;"
add-content $fileName  "margin-bottom: 10px;"
add-content $fileName  ""
add-content $fileName  "table {"
add-content $fileName  "border: thin solid #000000;"
add-content $fileName  "}"
add-content $fileName  "-->"
add-content $fileName  "</style>"

# HEAD end
Add-Content $fileName "</head>"

# BODY start
Add-Content $fileName "<body>"





Add-Content $fileName "`n "
Add-Content $fileName "`n# of Mailboxes Per Database (Minus Disconnected)"
Write-host " Compiling Mailbox count (Minus Disconnected)….  " -foregroundcolor magenta
# TABLE count START
Add-Content $fileName "<table width='100%'>"

# TABLE count  Header
Add-Content $fileName "<tr bgcolor='#7C7C7C'>"
Add-Content $fileName "<td width='35%'>ServerName</td>"
Add-Content $fileName "<td width='25%'>DB Name</td>"
Add-Content $fileName "<td width='40%'># of Boxes</td>"
Add-Content $fileName "</tr>"
$totalsize = 0
$tempsize = 0
$alternateTableRowBackground = 0
while($alternateTableRowBackground -le $dbcount.length -1 )
{
if(($alternateTableRowBackground % 2) -eq 0)
{
Add-Content $fileName "<tr bgcolor='#CCCCCC'>"
}
else
{
Add-Content $fileName "<tr bgcolor='#FCFCFC'>"
}
Add-Content $fileName ("<td width='35%'>" + $MailServer + "</td>") 
Add-Content $fileName ("<td width='25%'>" + $dbcount[$alternateTableRowBackground].name + "</td>")
Add-Content $fileName ("<td width='40%'>" + $dbcount[$alternateTableRowBackground].count + "</td>")
$tempsize = $totalsize + $dbcount.Count; 
$alternateTableRowBackground = $alternateTableRowBackground + 1
}
Add-Content $fileName ("<tr bgcolor= '#CCCC00'><td width='35%'>" + "Total Count"+ "</td>") 
Add-Content $fileName ("<td width='25%'>" + " " + "</td>")
$tempdata = Get-MailboxStatistics -server $MailServer |  ?{!$_.DisconnectDate} | Group-Object -Property:database | Select-Object Name, Count | %{$_.Count} | Measure-Object -Sum | Select-Object -expand Sum
Add-Content $fileName ("<td width='40%'>" + $tempdata + "</td></tr>")
#TABLE count  end
Add-Content $fileName "</table>"


Add-Content $fileName "`n "
Add-Content $fileName "`nSize Per Database "
Write-host " Compiling Database Sizes….  " -foregroundcolor magenta
# TABLE Dbsize START
Add-Content $fileName "<table width='100%'>"

# TABLE Dbsize Header
Add-Content $fileName "<tr bgcolor='#7C7C7C'>"
Add-Content $fileName "<td width='35%'>ServerName</td>"
Add-Content $fileName "<td width='25%'>DB Name</td>"
Add-Content $fileName "<td width='40%'>DB Size</td>"
Add-Content $fileName "</tr>"
$totalsize = 0
$tempsize = 0
$alternateTableRowBackground = 0
while($alternateTableRowBackground -le $dbsize.length -1)
{
if(($alternateTableRowBackground % 2) -eq 0)
{
Add-Content $fileName "<tr bgcolor='#CCCCCC'>"
}
else
{
Add-Content $fileName "<tr bgcolor='#FCFCFC'>"
}
Add-Content $fileName ("<td width='35%'>" + $dbsize[$alternateTableRowBackground].ServerName + "</td>") 
Add-Content $fileName ("<td width='25%'>" + $dbsize[$alternateTableRowBackground].name + "</td>")
Add-Content $fileName ("<td width='40%'>" + $dbsize[$alternateTableRowBackground].Databasesize + "</td>")
$tempsize = $totalsize + $dbsize.Databasesize; 
$alternateTableRowBackground = $alternateTableRowBackground + 1
}
Add-Content $fileName ("<tr bgcolor= '#CCCC00'><td width='35%'>" + "Total Size"+ "</td>") 
Add-Content $fileName ("<td width='25%'>" + " " + "</td>")
$tempdata = get-mailboxdatabase -status |%{$_.databasesize} | Measure-Object -Sum | Select-Object -expand Sum
$temp1 =($tempdata /1024 )
$temp2 = ($temp1 /1024 )
$temp3 = ($temp2 / 1024) 
$GBanswer = [math]::Round($temp3,2)
Add-Content $fileName ("<td width='40%'>" + $GBanswer + " GB </td></tr>")

#TABLE Dbsize end
Add-Content $fileName "</table>"


Add-Content $fileName "`n "
Add-Content $fileName "`nTop 10 Largest MailBoxs"
Write-host " Compiling Top 10 Largest Mailboxes….  " -foregroundcolor magenta
# TABLE Top10 start
Add-Content $fileName "<table width='100%'>"

# TABLE Top10 Header
Add-Content $fileName "<tr bgcolor='#7C7C7C'>"
Add-Content $fileName "<td width='35%'>DisplayName</td>"
Add-Content $fileName "<td width='10%'>ItemCount</td>"
Add-Content $fileName "<td width='10%'>TotalItemSize</td>"
Add-Content $fileName "<td width='25%'>Database</td>"
Add-Content $fileName "<td width='20%'>StorageLimitStatus</td>"
Add-Content $fileName "</tr>"

$alternateTableRowBackground = 0

# TABLE Top10 Content
while($alternateTableRowBackground -le $top10.length -1)
{
if(($alternateTableRowBackground % 2) -eq 0)
{
Add-Content $fileName "<tr bgcolor='#CCCCCC'>"
}
else
{
Add-Content $fileName "<tr bgcolor='#FCFCFC'>"
}
Add-Content $fileName ("<td width='30%'>" + $top10[$alternateTableRowBackground].DisplayName + "</td>") 
Add-Content $fileName ("<td width='10%'>" + $top10[$alternateTableRowBackground].ItemCount + "</td>")
Add-Content $fileName ("<td width='15%'>" + $top10[$alternateTableRowBackground].TotalItemSize + "</td>")
Add-Content $fileName ("<td width='25%'>" + $top10[$alternateTableRowBackground].Database + "</td>")


#BelowLimit or NoChecking
if(($top10[$alternateTableRowBackground].StorageLimitStatus -eq "BelowLimit") -or ($top10[$alternateTableRowBackground].StorageLimitStatus -eq "NoChecking"))
{
Add-Content $fileName ("<td bgcolor='#007F00' width='20%'>" + $top10[$alternateTableRowBackground].StorageLimitStatus + "</td>")
}
#IssueWarning
if($top10[$alternateTableRowBackground].StorageLimitStatus -eq "IssueWarning")
{
Add-Content $fileName ("<td bgcolor='#7F7F00' width='20%'>" + $top10[$alternateTableRowBackground].StorageLimitStatus + "</td>")
}
#ProhibitSend or MailboxDisabled
if(($top10[$alternateTableRowBackground].StorageLimitStatus -eq "ProhibitSend") -or ($top10[$alternateTableRowBackground].StorageLimitStatus -eq "MailboxDisabled"))
{
Add-Content $fileName ("<td bgcolor='#7F0000' width='20%'>" + $top10[$alternateTableRowBackground].StorageLimitStatus + "</td>")
}
Add-Content $fileName "</tr>"

$alternateTableRowBackground = $alternateTableRowBackground + 1
}


#TABLE Top10 end
Add-Content $fileName "</table>"

#Popup yes or no to Show Individual boxes
			#Button Types  
			# 
			#Value  Description   
			#0 Show OK button. 
			#1 Show OK and Cancel buttons. 
			#2 Show Abort, Retry, and Ignore buttons. 
			#3 Show Yes, No, and Cancel buttons. 
			#4 Show Yes and No buttons. 
			#5 Show Retry and Cancel buttons. 

			$a = new-object -comobject wscript.shell 
			$intAnswer = $a.popup("Do you want to Show Individual boxes?",0,"Show Individual boxes",4) 
			If ($intAnswer -eq 6) 
			{ # open IF to yes or no
			$a.popup("Individual boxes are being queried.")
			
			
Add-Content $fileName "`n "
Add-Content $fileName "`nIndividual MailBox (Including Disconnected)"



# TABLE all users start
Add-Content $fileName "<table width='100%'>"
Write-host " Compiling Everyone's Mailbox size (Including Disconnected)….  " -foregroundcolor magenta
# TABLE all users Header
Add-Content $fileName "<tr bgcolor='#7C7C7C'>"
Add-Content $fileName "<td width='35%'>DisplayName</td>"
Add-Content $fileName "<td width='10%'>ItemCount</td>"
Add-Content $fileName "<td width='10%'>TotalItemSize</td>"
Add-Content $fileName "<td width='25%'>Database</td>"
Add-Content $fileName "<td width='20%'>StorageLimitStatus</td>"
Add-Content $fileName "</tr>"

$alternateTableRowBackground = 0

# TABLE all users Content
while($alternateTableRowBackground -le $exdata.length -1)
{
if(($alternateTableRowBackground % 2) -eq 0)
{
Add-Content $fileName "<tr bgcolor='#CCCCCC'>"
}
else
{
Add-Content $fileName "<tr bgcolor='#FCFCFC'>"
}
Add-Content $fileName ("<td width='30%'>" + $exdata[$alternateTableRowBackground].DisplayName + "</td>") 
Add-Content $fileName ("<td width='10%'>" + $exdata[$alternateTableRowBackground].ItemCount + "</td>")
Add-Content $fileName ("<td width='15%'>" + $exdata[$alternateTableRowBackground].TotalItemSize + "</td>")
Add-Content $fileName ("<td width='25%'>" + $exdata[$alternateTableRowBackground].Database + "</td>")
#BelowLimit or NoChecking
if(($exdata[$alternateTableRowBackground].StorageLimitStatus -eq "BelowLimit") -or ($exdata[$alternateTableRowBackground].StorageLimitStatus -eq "NoChecking"))
{
Add-Content $fileName ("<td bgcolor='#007F00' width='20%'>" + $exdata[$alternateTableRowBackground].StorageLimitStatus + "</td>")
}
#IssueWarning
if($exdata[$alternateTableRowBackground].StorageLimitStatus -eq "IssueWarning")
{
Add-Content $fileName ("<td bgcolor='#7F7F00' width='20%'>" + $exdata[$alternateTableRowBackground].StorageLimitStatus + "</td>")
}
#ProhibitSend or MailboxDisabled
if(($exdata[$alternateTableRowBackground].StorageLimitStatus -eq "ProhibitSend") -or ($exdata[$alternateTableRowBackground].StorageLimitStatus -eq "MailboxDisabled"))
{
Add-Content $fileName ("<td bgcolor='#7F0000' width='20%'>" + $exdata[$alternateTableRowBackground].StorageLimitStatus + "</td>")
}
Add-Content $fileName "</tr>"

$alternateTableRowBackground = $alternateTableRowBackground + 1
}



#TABLE all users end
Add-Content $fileName "</table>"


} # Close IF to yes or no
				else { # open Else to yes or no
						$a.popup("Individual list wont be shown.") 
						
					} # Close Else to yes or no
					
# BODY end
Add-Content $fileName "</body>"

# HTML end
Add-Content $fileName "</html>"
Write-host " Creating File and Sending Email….  " -foregroundcolor magenta

if ($ReportRecipient)
{
sendEmail $MailServer $fileName
}

else {


Write-host " ---Closing script--- " -foregroundcolor magenta
}

Read-Host -Prompt "Press Enter to exit"
