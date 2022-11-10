<#
  .SYNOPSIS
	Этот скрипт по имени принтера собирает информацию по количеству работ за определенный период и отправляет на почту.
	Период выставлен с 07:00 вчерашнего дня до 07:00 текущего дня.
  .DESCRIPTION
	Умеет собирать данные с принтеров моделей: Kyocera FS-4200DN и Xerox VersaLink.
    Логирует события в файл .log.
	Агрегирует данные из: 1) База SQL SKDR
						  2) Папки PdfBuilder\worker\Temp на принт-серверах
					      3) Логи печати с принт-серверов (Kibana)
					      4) Из журналов принтеров
            Имя: Collect Printers Statistics
          Автор: Mironov Aleksandr (@just_slev1n)
		 Версия: 1.2
       Лицензия: Open Software License
     Требования: 1) SNMP Module (https://www.powershellgallery.com/packages/SNMP)
				 2) Доступ до ElasticSearch (Kibana)
				 3) Скрипт EWSAttachmentSaver.ps1
				 4) Доступ до принт-серверов и принтеров
				 5) Доступ к почтовому ящику mfu_skan@domen.ru
	  Допущения: 1) На всех принтерах стоит одинаковый пароль/логин (принятый стандарт в инвитро)
				 2) На всех принтерах включено snmp и одинаковое community name = public
				 3) Настроен на принтерах Kyocera FS-4200DN автоотчет на почту с темой ip принтера
				 4) Настроен на принтерах Xerox VersaLink https и журнал аудита
                 5) На Xerox VersaLink включено SSL и TLS 1.1 или старше
  .NOTES
	Так как принтер "Kyocera FS-9130DN NO SORTER" в базе SKDR только один, а физически это два устройства и два адреса, то значения их заданий складываются.
	Так как принтеры "Kyocera FS-C5250DN KX corp 5", "Kyocera FS-C5250DN KX_Cito" на самом деле физически одно устройство и один адрес, значения их складываются.
	Так-же из этого скрипта можно получить расширенный отчёт, который будет включать в себя количество отпечатанных страниц на принтере и в логах серверов печати.
#>

$StartTime = ((Get-Date).Date).AddDays(-1).AddHours(7)
$EndTime = ((Get-Date).Date).AddDays(0).AddHours(7)

# Входные параметры
$Printers = "Kyocera FS-9130DN NO SORTER", "Kyocera_P6026cdn_MEDKVADRAT_SD", "Kyocera FS-C5250DN KX corp 5", "Kyocera FS-C5250DN KX_Cito", "Kyocera FS-4000DN KX ClinRes", "korp5_od_Dializ_Frezen_Bibraun_FS4200"
$PrintServers = "sv-sd05", "sv-sd02"
$ScriptPath = "C:\Users\amironov\Desktop\PrinterLogs\CollectPrintersStatistics"
# Для отправки email
$MailBox = "mfu_skan@domen.ru"
$MailPassword = "password"
$password = ConvertTo-SecureString $MailPassword -AsPlainText -Force
$Cred = New-Object System.Management.Automation.PSCredential ($MailBox, $password)

# Логируем события скрипта
Function LogWrite {
    Param ([string]$logstring)
    $logfile =  $PSCommandPath -replace ".ps1", ".log"
    if (!(Test-Path $logfile)) {
        New-Item -ItemType File $logfile | Out-Null
    } else {
        $time = Get-Date -Format "dd.MM.yyyy HH:mm:ss: "
        Add-content $logfile -value ($time + $logstring)
	}
}

LogWrite "Запуск $($MyInvocation.MyCommand.Name)"
LogWrite "C $StartTime по $EndTime"

# Проверка, что время указано корректно
if ([datetime]$EndTime -le [datetime]$StartTime) {
    LogWrite -logstring "Дата окончания не может быть меньше или равна даты начала!"
    exit
}
if ([datetime]$EndTime -gt (Get-Date)){
    LogWrite -logstring "Дата окончания не может быть больше чем текущая дата!"
    exit
}

# Проверяем доступность по snmp и ждем до 30 секунд
Function GetName-FromSNMP ([ipaddress]$ip) {
    [String]$realname = ""
	while (!($realname) -and ($i -lt 30)) {
		$RealName = (Get-SnmpData -ip $IP -OID .1.3.6.1.2.1.25.3.2.1.3.1).data
		Start-Sleep -Seconds 1
		$i++
	}
    if (!($realname)) { LogWrite "Не удалось подключиться по SNMP до $ip" }
   	return $realname
}

# Собираем ip адреса с принтсерверов
foreach ($server in $PrintServers[0]) {

    $AllPrintersFromServer += Get-Printer -ComputerName $server
}
# Добавляем данные IP и реальное имя
$PrinterList = @()
foreach ($Printer in $Printers) {
    $ipRange = ""
    $ipRange = ($AllPrintersFromServer | where Name -eq $Printer | select -First 1).PortName
    if (!($ipRange)) {
        LogWrite -logstring "Не удалось получить IP для $Printer с $PrintServers"
    }
    foreach ($ip in $ipRange.Split(",")) {
	    $RealName = GetName-FromSNMP -ip $ip
	    if ($RealName -eq "FS-4200DN") { $RealName = "Kyocera " + $RealName }
        $prop = [ordered]@{
            Name=$Printer
            ipAddress=[ipaddress]$ip
            RealName = $RealName
        }
        $PrinterList += New-Object –TypeName PSObject –Property $prop
    }
}

# Считаем по количеству файлов в папке Temp на ПринтСерверах
Function GetCounts-FromTemp () {
    $RemoteFiles = Invoke-Command -ComputerName $PrintServers -ScriptBlock {
        Param($StartTime,$EndTime)
        Get-ChildItem -Path "D:\PdfBuilder\worker\Temp" | select Name, CreationTime, Length | Where {
        ($_.CreationTime -gt $StartTime -and $_.CreationTime -le $EndTime) }
    } -ArgumentList ($StartTime, $EndTime) -ErrorVariable errmsg
    if ($errmsg) { LogWrite $errmsg.exception }

    foreach ($file in $RemoteFiles) {
        $Match = [regex]::Match($File.Name, "_\w{8}-")
        if ($Match.Success) {
            $file.Name = $file.Name.Substring(0,$Match.Index)
        }
    }

    if ($RemoteFiles) {
        LogWrite "Получено $($RemoteFiles.count) записей из папок Temp"
    } else {
        LogWrite "Ошибка получения количества файлов из папки Temp с принтсерверов!"
    }

    return $RemoteFiles | where {($Printers -contains $_.Name)} | Group-Object -Property Name | select Name, Count
}

# Берём данные из логов печати через Кибану
Function GetCounts-FromKibana () {

    # Проверяем, что сервер SR-ADM02 доступен
    Invoke-Command -ComputerName SR-ADM02 -ScriptBlock { echo Ping } -AsJob -JobName TestConn
    do {
        Start-Sleep -Seconds 1
        $Connection = (Receive-Job -name TestConn) -eq "Ping"
        $i++
    } while (($i -le 5) -and !$Connection)
    if (!$Connection) {
        LogWrite "Не удалось подключиться к SR-ADM02"
        Remove-Job TestConn
        return 0
    }

    # -3 часа смещение часового пояса
    $StartTimeUnix = $StartTime.AddHours(-3).Subtract((Get-Date 1/1/1970)).TotalMilliseconds
    $EndTimeUnix = $EndTime.AddHours(-3).Subtract((Get-Date 1/1/1970)).TotalMilliseconds
    $obj = @()
    foreach ($Printer in $Printers) {
 $body = '{
    "from": 0,
    "size": 10000,
  "query": {
    "bool": {
      "must": [
        {
          "match_all": {}
        },
        {
          "match_phrase": {
            "log_name": {
              "query": "Microsoft-Windows-PrintService/Operational"
            }
          }
        },
        {
          "bool": {
            "minimum_should_match": 1,
            "should": [
              {
                "match_phrase": {
                  "host.name": "SV-SD02"
                }
              },
              {
                "match_phrase": {
                  "host.name": "SV-SD05"
                }
              }
            ]
          }
        },
        {
          "match_phrase": {
            "event_id": {
              "query": 307
            }
          }
        },
        {
          "bool": {
            "minimum_should_match": 1,
            "should": [
              {
                "match_phrase": {
                  "user_data.Param5": "' + $Printer + '"
                }
              }
            ]
          }
        },
                {
          "range": {
            "@timestamp": {
              "gte": ' + $StartTimeUnix + ',
              "lte": ' + $EndTimeUnix + ',
              "format": "epoch_millis"
            }
          }
        }
      ]
    }
  }
  }'
    $req = Invoke-Command -ComputerName sr-adm02 -ScriptBlock {
    param ($body)
    Invoke-RestMethod -URI "http://sv-kb01:9200/winlogbeat_*/_search?pretty" -Method 'POST' -ContentType 'application/json' -Body $body
    } -ArgumentList $body

    if ($req) {
        LogWrite "Получено строк из кибаны $($req.hits.total) для $Printer"
    } else {
        LogWrite "Ошибка получения данных из кибаны $Printer"
        }

	Start-Sleep -Seconds 2     # ждем между запросами 2 секунды

    if ($req.hits.total -eq 0) { continue } # если не найдено записей, то пропускаем

	# считаем количество страниц по принтеру
    $TotalPages = 0
    foreach ($item in $req.hits.hits._source.user_data) {
        $TotalPages += $item.Param8
    }
    $prop = New-Object psobject -Property @{
        Name = $Printer
        Jobs = $req.hits.total
        Pages = $TotalPages
        }
    $obj += $prop
    }
    return $obj
}

# Берём данные из SQL SKDR
Function GetCounts-FromSKDR () {
    $dataSource = "fcsql07"
    $database = "SKDR"
    $sql = "select
    cd.Destination printerName,
    count(cd.Id) rows
    from SKDR.dbo.CoolDelivery_Log as cd (nolock)
    where  cd.[TimeStamp] >= dateadd(hour, -3,'$($StartTime.ToString("yyyyMMdd HH:mm:ss"))') and cd.[TimeStamp] < dateadd(hour, -3, '$($EndTime.toString("yyyyMMdd HH:mm:ss"))')
    and cd.DeliveryType = 0
    and cd.Message is null
    group by cd.Destination"
    $auth = "User ID = PrintReport; Password = "
    $connectionString = "Provider=sqloledb; " +
    "Data Source=$dataSource; " +
    "Initial Catalog=$database; " +
    "$auth; "
    $connection = New-Object System.Data.OleDb.OleDbConnection $connectionString
    $command = New-Object System.Data.OleDb.OleDbCommand $sql,$connection
    $connection.Open()
    $adapter = New-Object System.Data.OleDb.OleDbDataAdapter $command
    $dataset = New-Object System.Data.DataSet
    [void] $adapter.Fill($dataSet)
    $connection.Close()
    $rows = ($dataset.Tables | Select-Object -Expand Rows)
    $rows = $rows | where {$Printers -contains $_.printerName}
    if ($rows) {
        LogWrite -logstring "Получено строк из СКДР $(($rows | Measure-Object -Property rows -sum).sum)"
    } else {
        LogWrite -logstring "Ошибка получения данных из СКДР"
    }
    return $rows
}

# Берём данные с принтеров Ксерокс
Function GetReport-FromXerox ($ipString, $LogFolder) {
    #поддержка SSL
add-type @"
    using System.Net;
    using System.Security.Cryptography.X509Certificates;
    public class TrustAllCertsPolicy : ICertificatePolicy {
        public bool CheckValidationResult(
            ServicePoint srvPoint, X509Certificate certificate,
            WebRequest request, int certificateProblem) {
            return true;
        }
    }
"@
    [System.Net.ServicePointManager]::CertificatePolicy = New-Object TrustAllCertsPolicy
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    $Fields = @{"name"="name";"PSW"="password"}

    # что-бы быстрей скачивалось
    $ProgressPreference = 'SilentlyContinue'

    # сначала проверяем есть ли нужный лог в папке
    $CheckLocalFile = Get-ChildItem $LogFolder -Filter *.txt |`
    where {($_.Name -match $ipString) -and ($_.LastWriteTime -gt $EndTime) -and ($_.Length -gt 140)} | select -Last 1

    if ($CheckLocalFile) {
        # если есть берем оттуда
        $auditFile = Get-Content -path $CheckLocalFile.fullname
        LogWrite "Найден подходящий файл $($CheckLocalFile.Name)"
    } else {
        # Проходим авторизацию
        $rest = Invoke-WebRequest ("https://"+$ipString+"/LOGIN.cmd") -ContentType "application/x-www-form-urlencoded; charset=UTF-8" -Body $Fields -Method Post -SessionVariable session
        LogWrite "$($ipString): ответ $($rest.StatusCode)"

        # скачиваем файл аудита
        $MeasureDownload = Measure-Command -Expression {
            $Response = Invoke-WebRequest ("https://"+$ipString+"/ALOGEXPT.CMD") -Method Post -WebSession $session
        }
        if (($Response.StatusCode -eq 200) -and ($Response.RawContent -gt 140)) {
            $ExportFileName = "\auditfile_"+$ipString+(Get-Date).ToString("_dd_MM_yyyy_HH_mm_ss")+".txt"
            $FilePath = $LogFolder + $ExportFileName
            $Response.Content | Out-File -FilePath $FilePath -Encoding utf8
            $auditFile = $Response.Content
            LogWrite "Скачен файл $($ExportFileName) за $([int]$MeasureDownload.TotalSeconds) секунд"
        } else { return 0 }

    }

    $result = Parse-AuditFileXerox -FileContent $auditFile
    LogWrite -logstring "$($ipString): Получено строк из файла $($result)"

    return $result
}

# Парсим данные из txt для Xerox, возвращает количество работ
Function Parse-AuditFileXerox($FileContent) {
    $MatchResult = $FileContent.Split([Environment]::NewLine) | Select-String -Pattern "0x0401" | where {
    ($_ -notmatch "PS Initialize|Print Reports") -and ($_ -match "Completed") -and ($_ -match "(\d{2}/\d{2}/\d{4})\s(\d{2}:\d{2}:\d{2})") }` | where {
                $DateTime = [DateTime]::Parse($Matches[0])
                (($DateTime -ge $StartTime) -and ($DateTime -lt $EndTime))
                }
    return ($MatchResult | Measure-Object).Count
}

# Берём данные из Kyocera FS-4200DN
Function GetReport-FromKyocera ($ipString, $LogFolder) {
    # сначала проверяем файл с последней выгрузкой
    $CheckLocalFile = Get-ChildItem $LogFolder -Filter *.xml | where {($_.Name -match "export_job_logResponse") -and ($_.CreationTime -gt $EndTime)} | select -Last 1

    # если файл не нашли, то запрашиваем с принтера
    if (!($CheckLocalFile)) {
        $ContentType = "application/x-www-form-urlencoded; charset=UTF-8"

        # Wake UP Neo
        do {
            $Rest = Invoke-WebRequest ("http://" + $ipString) -ContentType $ContentType -Method Post
            $CheckSleep = $Rest.AllElements.src -match "DeepSleep"
            if ($CheckSleep) {
                $Body = "submit001=%D0%9F%D1%83%D1%81%D0%BA&okhtmfile=DeepSleepApply.htm&func=wakeup"
                $WakeUp = Invoke-WebRequest ("http://" + $ipString + "/esu/set.cgi") -ContentType $ContentType -Body $Body -Method Post
            }
            Start-Sleep -Seconds 5
        } while ($CheckSleep)

        # Проходим авторизацию
        $Body = "failhtmfile=%2Fstartwlm%2FStart_Wlm.htm&okhtmfile=%2Fstartwlm%2FStart_Wlm.htm&func=authLogin&arg03_LoginType=_mode_off&arg04_LoginFrom=_wlm_login&language=..%2Fwlmpor%2Findex.htm&hiddRefreshDevice=..%2Fstartwlm%2FHme_DvcSts.htm&hiddRefreshPanelUsed=..%2Fstartwlm%2FHme_PnlUsg.htm&hiddRefreshPaperid=..%2Fstartwlm%2FHme_Paper.htm&hiddRefreshTonerid=..%2Fstartwlm%2FHme_StplPnch.htm&hiddRefreshStapleid=..%2Fstartwlm%2FHme_Toner.htm&hiddnBackNavIndx=1&hiddRefreshDeviceBack=&hiddRefreshPanelUsedBack=&hiddRefreshPaperidBack=&hiddRefreshToneridBack=&hiddRefreshStapleidBack=&hiddCompatibility=&hiddPasswordToOpenChk=&hiddPasswordToOpen=&hiddRePasswordToOpen=&hiddPasswordToEditChk=&hiddPasswordToEdit=&hiddRePasswordToEdit=&hiddPrinting=&hiddChanges=&hiddCopyingOfText=&hiddEmaiSID=&hiddEmaiName=&hiddECM=&hiddDocID=&privid=&publicid=&attrtype=&attrname=&hiddFaxType=&hiddSMBNumber1=&hiddSMBNumber2=&hiddSMBNumber3=&hiddSMBNumber4=&hiddSMBNumber5=&hiddSMBNumber6=&hiddSMBNumber7=&hiddFTPNumber1=&hiddFTPNumber2=&hiddFTPNumber3=&hiddFTPNumber4=&hiddFTPNumber5=&hiddFTPNumber6=&hiddFTPNumber7=&hiddFAXAddress1=&hiddFAXAddress2=&hiddFAXAddress3=&hiddFAXAddress4=&hiddFAXAddress5=&hiddFAXAddress6=&hiddFAXAddress7=&hiddFAXAddress8=&hiddFAXAddress9=&hiddFAXAddress10=&hiddIFaxAdd=&hiddIFaxConnMode=&hiddIFaxResolution=&hiddIFaxComplession=&hiddIFaxPaperSize=&hiddImage=&hiddTest=&hiddDoc_Num=&hiddCopy=&hiddDocument=&hiddDocRec=&AddressNumberPersonal=&AddressNumberGroup=&hiddPersonAddressID=&hiddGroupAddressID=&hiddPrnBasic=&hiddPageName=&hiddDwnLoadType=&hiddPrintType=&hiddSend1=&hiddSend2=&hiddSend3=&hiddSend4=&hiddSend5=&hiddAddrBokBackUrl=&hiddAddrBokName=&hiddAddrBokFname=&hiddSendFileName=&hiddenAddressbook=&hiddenAddressbook1=&hiddSendDoc_Num=&hiddSendColor=&hiddSendAddInfo=&hiddSendFileFormat=&hiddRefreshDevice=..%2Fstartwlm%2FHme_DvcSts.htm&hiddRefreshPanelUsed=..%2Fstartwlm%2FHme_PnlUsg.htm&hiddRefreshPaperid=..%2Fstartwlm%2FHme_Paper.htm&hiddRefreshTonerid=..%2Fstartwlm%2FHme_StplPnch.htm&hiddRefreshStapleid=..%2Fstartwlm%2FHme_Toner.htm&hiddnBackNavIndx=0&hiddRefreshDeviceBack=&hiddRefreshPanelUsedBack=&hiddRefreshPaperidBack=&hiddRefreshToneridBack=&hiddRefreshStapleidBack=&hiddValue=&arg01_UserName=admin&arg02_Password=Inviprint34&arg03_LoginType=&arg05_AccountId=&Login=%D0%92%D1%85%D0%BE%D0%B4+%D0%B2+%D1%81%D0%B8%D1%81%D1%82%D0%B5%D0%BC%D1%83&hndHeight=0"
        $Rest = Invoke-WebRequest ("http://" + $ipString + "/startwlm/login.cgi") -ContentType $ContentType -Body $Body -Method Post -SessionVariable Session

        # Запрашиваем файл лога печати
        $Body = "okhtmfile=%2Fadv%2FJoblognoteSendResult.htm&failhtmfile=%2Fadv%2FAdvError.htm&func=sendSMTPJoblognote"
        $Rest = Invoke-WebRequest ("http://" + $ipString + "/adv/set.cgi") -ContentType $ContentType -Body $Body -Method Post -WebSession $Session

        LogWrite -logstring "$($ipString): ответ $($Rest.StatusCode)"

        # Ждём 10 секунд пока придёт письмо
        Start-Sleep -Seconds 10

        #Запускаем скрипт сбора вложений из письма
        #-mailbox $mailbox -MailPassword $MailPassword -ScriptPath $ScriptPath -subjectfilter $ipString
        $ArgumentList = $ScriptPath + "\EWSAttachmentSaver.ps1 " + $mailbox + " " + $MailPassword + " " + $ScriptPath + " " + $ipString
		LogWrite "Запуск скрипта EWSAttachmentSaver.ps1"
        Start-Process -FilePath powershell.exe -ArgumentList $ArgumentList -Wait
        LogWrite "Окончание работы скрипта EWSAttachmentSaver.ps1"
    }

    #Запускаем парсер
    $ParserResult = Parse-XML -FolderPath ($ScriptPath + "\" +$ipString)
    LogWrite "$($ipString): Получено строк из файла $($ParserResult.TotalJobs)"
    return $ParserResult.TotalJobs
}

# Парсим данные из XML для Kyocera, возвращает количество работ
Function Parse-XML($FolderPath) {
    $ExportJobs = Get-ChildItem $FolderPath -Filter *.xml | where { ($_.LastWriteTime -gt ((Get-Date).AddDays(-4)))}

    if ($ExportJobs.Count -eq 0) {
        LogWrite "Xml not found in $FolderPath!"
        return 0
    }

    Function ConvertXmlTime ($XMLTime) {
        $year = [int]$XMLTime.year.replace("12","202")
        $month = [int]$XMLTime.month + 1
        $date = Get-Date -Day $XMLTime.day -Year $year -Month $month -Hour $XMLTime.hour -Minute $XMLTime.minute -Second $XMLTime.Second
        return $date
    }

    $obj = @()
    foreach ($ExportJob in $ExportJobs) {
        [XML]$XmlLog = Get-Content $ExportJob.FullName
        $PrintJobs = $XmlLog.export_job_logResponse.export_job_log.print_job_log
        foreach ($PrintJob in $PrintJobs) {
            $prop = [ordered]@{
                job_number = $PrintJob.common.job_number
                job_name = $PrintJob.common.job_name
                user_name = $PrintJob.common.user_name
                job_result = $PrintJob.common.job_result
                job_DateTime = ConvertXmlTime -XMLTime $PrintJob.common.start_time
                complete_pages = $PrintJob.detail.complete_copies
                complete_copies = $PrintJob.detail.complete_pages
                total_pages = [int]$PrintJob.detail.complete_copies*[int]$PrintJob.detail.complete_pages
                }
            $obj += New-Object –TypeName PSObject –Property $prop
        }
    }
    $MeasureObj = $obj | where {($_.job_DateTime -ge $StartTime) -and ($_.job_DateTime -lt $EndTime) -and ($_.total_pages -ne 0)} | Measure-Object -Property total_pages -Sum
    $prop = [ordered]@{
        TotalJobs = $MeasureObj.Count
        TotalPages = $MeasureObj.Sum
        }
    $ReturnObj = New-Object –TypeName PSObject –Property $prop
    return $ReturnObj
}

# Объединяем данные со всех принтеров
Function GetCounters-FromPrinters {
    $obj = @()

    # берём уникальные IP
    $PrinterListunique = $PrinterList | Group-Object ipAddress | %{ $_.Group | Select RealName, ipAddress, Name -First 1}

    foreach ($Printer in $PrinterListunique) {
        [int]$counter = 0
        $ipString = $Printer.ipAddress.IPAddressToString
        $FolderPath = $ScriptPath+"\"+$ipString

        # проверка связи
        if (!(Test-Connection $ipString -Count 5 -Delay 2 -Quiet)) {
            LogWrite "Не удалось установить связь с $ipString!"
            Write-Host "Не удалось установить связь с" $ipString -ForegroundColor Red
            continue
        }

        # создаём папку лога
        if (!(Test-Path -Path $FolderPath -ErrorAction SilentlyContinue)) {
            New-Item -Path $FolderPath -ItemType Directory
        }

        if ($Printer.RealName -match "Xerox VersaLink") {
            $counter = GetReport-FromXerox -ip $ipString -LogFolder $FolderPath
        }

        if ($Printer.RealName -match "FS-4200DN") {
            $counter = GetReport-FromKyocera -ip $ipString -LogFolder $FolderPath
        }

        $prop = New-Object psobject -Property @{
            IP = $Printer.ipAddress
            Jobs = $counter
            RealName = $Printer.RealName
            Name = $Printer.Name
            }
        $obj += $prop
    }
    return $obj
}

function Send-EmailHtml($SendObj, $EmailTO, $CopyTO) {
    $ExportDate = (Get-Date -Format "dd_MM_yyyy")
    $Message = "Задания печати с " + $StartTime.ToString("dd.MM.yyyy HH:mm:ss") + " по " + $EndTime.ToString("dd.MM.yyyy HH:mm:ss")

    # Отправка по почте
    $style = "<style>BODY{font-family: Open Sans; font-size: 12pt;}"
    $style = $style + "TABLE{border: 1px solid black; border-collapse: collapse;}"
    $style = $style + "TH{border: 1px solid black; background: #dddddd; padding: 5px; }"
    $style = $style + "TD{border: 1px solid black; padding: 5px; }"
    $style = $style + "</style>"
$Header = @"
$style
"@
    Send-MailMessage -From $MailBox -To $EmailTO -Subject 'Аудит печати' -Body ($Message + [string]($SendObj | ConvertTo-Html -Head $Header))`
     -BodyAsHtml -Encoding UTF8 -SmtpServer 'cas.invitro.ru' -Credential $cred
    LogWrite "Отправлено письмо $EmailTO"
}

$checkTemp = GetCounts-FromTemp
$checkKibana = GetCounts-FromKibana
$checkSKDR = GetCounts-FromSKDR
$checkPrinters = GetCounters-FromPrinters

$TotalObj = @()
foreach ($Printer in $Printers) {
    $property = [ordered]@{
        Name = [string]$Printer
        RealName = [String]($PrinterList | where Name -eq $Printer | select -First 1).RealName
        IP = [string]($PrinterList | where Name -eq $Printer).ipaddress.IPAddressToString
        Temp = [string]($checkTemp | where Name -eq $Printer).Count
		SKDR = [int]($checkSKDR | where printerName -eq $Printer).Rows
        Kibana = [int]($checkKibana | where Name -eq $Printer | Measure-Object -Property Jobs -Sum).Sum
        Printer = [int]($checkPrinters | where Name -eq $Printer | Measure-Object -Property Jobs -Sum).Sum
        }
    $TotalObj += New-Object –TypeName PSObject –Property $property
}

# Костыль объединения двух одинаковых принтеров
if (($TotalObj.name -contains "Kyocera FS-C5250DN KX corp 5") -and ($TotalObj.name -contains "Kyocera FS-C5250DN KX_Cito")) {
    $counter = 0
    foreach ($tobj in $TotalObj) {
        if ($tobj.name -eq "Kyocera FS-C5250DN KX corp 5") { $i = $counter }
        if ($tobj.name -eq "Kyocera FS-C5250DN KX_Cito")   { $j = $counter }
    $counter++
    }
    $TotalObj[$i].Name = "Kyocera FS-C5250DN KX*"
    $TotalObj[$i].SKDR += $TotalObj[$j].SKDR
    $TotalObj[$i].Kibana += $TotalObj[$j].Kibana
    $TotalObj = $TotalObj | where Name -NE 'Kyocera FS-C5250DN KX_Cito'
}


#Посылаем обрезанный и полные отчёты
Send-EmailHtml -SendObj ($TotalObj | select RealName, IP, SKDR, Printer) -EmailTO ("otchet_po_zadaniyam_na_pechat@domen.ru", "amironov@domen.ru", "akolosov@domen.ru")
Send-EmailHtml -SendObj $TotalObj -EmailTO "amironov@domen.ru"

LogWrite "Окончание $($MyInvocation.MyCommand.Name)"
