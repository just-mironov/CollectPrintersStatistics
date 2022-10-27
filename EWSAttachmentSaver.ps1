<#
  .SYNOPSIS
	This script looks for specific emails in an exchange users mailbox, downloads the attachments, 
	then marks those emails as read and moves the messages to a processed folder for archiving.
           Name: EWS Email Attachment Saver
         Author: Spencer Alessi (@techspence)
        License: MIT License
    Assumptions: The 'processed folder' is a subfolder of the root of the users mailbox (e.g. \\email@company.com\ProcessedFolder)
   Requirements: Exchange 2007 or newer
			     			 Exchange Web Services (EWS) Managed API 2.2
  .DESCRIPTION
	In general this script:
		1. Determines the Folder ID of the $processedfolderpath
		2. Finds the correct email messages based on defined search filters (e.g. unread, subject, has attachments)
		3. Copy's the attachments to the appropriate download location(s)
		4. Mark emails as read and move to the processed folder
  
  .NOTES
	The 'processed folder' is a subfolder of the root of the users mailbox (e.g. \\email@company.com\ProcessedFolder). 
	The root of a users mailbox is called the Top Information Store. If your 'processed folder' is a subfolder under any other 
	folder you must change $processedfolderpath and $tftargetidroot appropriately. 
	In this example, the processed folder is a subfolder of the root mailbox: Location: \\\email@company.com\ProcessedFolder
		$processedfolderpath = "/ProcessedFolder"
		$tftargetidroot = New-Object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::MsgFolderRoot,$mailbox)
	In this example, the processed folder is a subfolder of Inbox: Location: \\\email@company.com\Inbox\ProcessedFolder
		
		$processedfolderpath = "/Inbox/ProcessedFolder"
		$tftargetidroot = New-Object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Inbox,$processedfolderpath)

    Важно! ip принтера = должен быть одинаковым для 1) тема письма 2) в пути до папки
#>
   

    [CmdletBinding()]
    Param
    (
        [Parameter(Mandatory=$true, Position=0, HelpMessage="Адрес почтового ящики в формате email@host.com")]
        [string]$mailbox,      
        [Parameter(Mandatory=$true, Position=1, HelpMessage="Пароль от ящика")]
        [string]$MailPassword,
        [Parameter(Mandatory=$true, Position=2, HelpMessage="Путь до скрипта")] 
        [string]$ScriptPath,
        [Parameter(Mandatory=$true, Position=3, HelpMessage="Тема письма")] 
        [string]$subjectfilter
    )
Begin {
    Write-Host "Start Saver" -ForegroundColor Yellow
    $serverreportfolder = $ScriptPath + "\" + $subjectfilter + "\"
    $logname = "EWSAttachmentSaver-$(get-date -f yyyy-MM-dd).log"
    $logfile = $serverreportfolder + $logname
    $processedfolderpath = "/Kyocera_Log"
    $datestamp = (Get-Date).toString("dd/MM/YYYY HH:mm:ss")
}
Process {
    Function LogWrite
    {
	    Param ([string]$logstring)
	
	    if (!(Test-Path $serverreportfolder)) {
		    New-Item -ItemType Directory $serverreportfolder | Out-Null
	    } 
	    else { 
		    Add-content $logfile -value $logstring
	    }
    }

    Function FindTargetFolder($folderpath){
	    $tftargetidroot = New-Object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::MsgFolderRoot,$mailbox)
	    $tftargetfolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($exchangeservice,$tftargetidroot)
        $pfarray = $folderpath.Split("/")
	
	    # Loop processed folders path until target folder is found
	    for ($i = 1; $i -lt $pfarray.Length; $i++){
		    $fvfolderview = New-Object Microsoft.Exchange.WebServices.Data.FolderView(1)
		    $sfsearchfilter = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.FolderSchema]::DisplayName,$pfarray[$i])
        $findfolderresults = $exchangeservice.FindFolders($tftargetfolder.Id,$sfsearchfilter,$fvfolderview)
		
		    if ($findfolderresults.TotalCount -gt 0){
			    foreach ($folder in $findfolderresults.Folders){
				    $tftargetfolder = $folder				
			    }
		    }
		    else {
			    LogWrite "### Error ###"
			    Logwrite $datestamp " : Folder Not Found"
			    $tftargetfolder = $null
			    break
		    }	
	    }
	    $Global:findFolder = $tfTargetFolder
    }

    Function FindTargetEmail($subject){
	    foreach ($email in $foundemails.Items){
		    $email.Load()
		    $attachments = $email.Attachments
		    foreach ($attachment in $attachments){
			    $attachment.Load()
			    $attachmentname = $attachment.Name.ToString()
					$ExportName = $email.DateTimeReceived.ToString("dd.MM.yyyy_HH.mm.ss_") + $attachmentname
				    LogWrite "$ExportName saved to $serverreportfolder"
                    $ExportFullName = $serverreportfolder + $ExportName
				    $file = New-Object System.IO.FileStream($ExportFullName, [System.IO.FileMode]::Create)	
				    $file.Write($attachment.Content, 0, $attachment.Content.Length)
				    $file.Close()
			    }
	    # Mark email as read & move to processed folder
	    $email.IsRead = $true
	    $email.Update([Microsoft.Exchange.WebServices.Data.ConflictResolutionMode]::AlwaysOverwrite)
	    [VOID]$email.Move($Global:findFolder.Id)
        }
    }

    LogWrite "DATETIME START: $datestamp"
    LogWrite "Mailbox: $mailbox"
    LogWrite "Processed Folder: $processedfolderpath"
    LogWrite "Subject Filter: $subjectfilter"

    # Load the EWS Managed API
    $dllpath = $ScriptPath + "\Microsoft.Exchange.WebServices.dll"
    [void][Reflection.Assembly]::LoadFile($dllpath)

    # Create EWS Service object for the target mailbox name
    # Note, ExchangeVersion does not need to match the version of your Exchange server
    # You set the version to indicate the lowest level of service you support
    $exchangeservice = new-object Microsoft.Exchange.WebServices.Data.ExchangeService
    $exchangeservice.Credentials = New-Object Microsoft.Exchange.WebServices.Data.WebCredentials($mailbox,$mailpassword)
    $exchangeservice.AutodiscoverUrl($mailbox)

    # Bind to the Inbox folder of the target mailbox
    $inboxfolderid = New-Object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Inbox,$mailbox)
    $inboxfolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($exchangeservice,$inboxfolderid)

    # Search the Inbox for messages that are: unread, has specific subject AND has attachment(s)
    $sfunread = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.EmailMessageSchema]::IsRead, $false)
    $sfsubject = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+ContainsSubstring ([Microsoft.Exchange.WebServices.Data.EmailMessageSchema]::Subject, $subjectfilter)
    $sfattachment = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.EmailMessageSchema]::HasAttachments, $true)
    $sfcollection = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+SearchFilterCollection([Microsoft.Exchange.WebServices.Data.LogicalOperator]::And);
    $sfcollection.add($sfunread)
    $sfcollection.add($sfsubject)
    $sfcollection.add($sfattachment)

    # Use -ArgumentList 30 to reduce query overhead by viewing the Inbox 10 items at a time
    $view = New-Object -TypeName Microsoft.Exchange.WebServices.Data.ItemView -ArgumentList 30
    $foundemails = $inboxfolder.FindItems($sfcollection,$view)

    # Find $processedfolderpath Folder ID
    FindTargetFolder($processedfolderpath)

    # Process found emails
    FindTargetEmail($subject)

    $message = "Найдено писем по теме """+$subjectfilter+""": "+$foundemails.Items.count
    Write-Host $message -ForegroundColor Yellow
	
	LogWrite "DATETIME END: $datestamp"
}