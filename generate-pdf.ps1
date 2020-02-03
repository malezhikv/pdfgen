#
#  функции для работы с документами PDF:
#
#  Generate-MyCloudPDF   - генерация по шаблону
#  Convert-NumberToText  - вспомогательные функции (число-в-текст,...)

#  Generate-MyCloudPDF - генерация документа по шаблону. 
#  На вход структура данных JSON -> ConvertFrom-Json включающая данные в свойствах объекта, а также массивы в свойствах объекта. Также шаблон HTML и имя выходного файла.
#  Пример структуры данных:  
#    { Name: "TEST",
#      Phone: "0001234567",
#      Emails: [ { Email: "mail@server.com", IsPrimary: false }, { Email: "mail2@server.com", IsPrimary: true } ]
#    ... }
#
#  В шаблоне при подстановке значений данными происходит поиск свойств в структуре данных по названию переменных из шаблона. Таким образом на вход нужно подавать структуру данных с названием свойств, которые используются в шаблоне
#  Пример (текст из шаблона):
#  ..<p>$Name$</p>
#    <p>$Phone$</p>
#    <table>
#      <tr><td>$Email$</td><td>$IsPrimary$</td></tr>
#    </table>..
#
#  На выходе - PDF файл (с учетом пути).

function Generate-MyCloudPDF {
    param($ClientInfo,               # объект PSObject с информацией о клиенте и его заказах, см. ниже структуру
        [string]$Template,           # путь к файлу-шаблону HTML
        [string]$Footer,             # путь к файлу-шаблону HTML footer - для многостраничных документов
        [string]$Header,             # путь к файлу-шаблону HTML header - для многостраничных документов
        [string]$DocumentDate="",
		    [string]$OutputFileName,     # опционально имя файла c расширением pdf (пример 129_10.01.2019_invoice.pdf)
        [bool]$Debug=$false)         # true = не удалять промежуточные файлы: html и отдельные pdf

  $ClientEDPROU = $ClientInfo.Id

	if(!$OutputFileName) {
		$OutputFileName = ('.\{0}\{0}_{1}.pdf' -f $ClientEDPROU,$docName)
	}
	
	$OutputFileNameHTML = $OutputFileName.Replace("pdf","html")
	
  # генерим хтмл из шаблона заменяя параметры значениями из входной структуры данных 

  $docName = (Get-ChildItem $template).Name.Replace(".html","")
  $docPath = (Get-ChildItem $template).FullName.Replace(".html","")
  $i=1
  $j=1
  $s = $ClientInfo.SellerBankRequisites # example "р/р 26005052634833 в АТ КБ ПРИВАТБАНК МФО 300711"
  $SellerBankAcc = ""
  $SellerBankMFO = ""
  $SellerBankDep = ""
  $r = select-string -allmatches -inputobject $s -pattern "\b\d{14}\b" # Account Number - 14 digits
  if($r) { $SellerBankAcc = $r.Matches[0].Value }
  $r = select-string -allmatches -inputobject $s -pattern "\b\d{6}\b" # MFO - 6 digits
  if($r) { $SellerBankMFO = $r.Matches[0].Value }
  # bank department  
  $x=$s.Substring(21,$s.Length-21) 
  $y=$x.substring(0,$x.IndexOf("МФО")) # y = OHO
  if($y) { $SellerBankDep = $y }
  # ToPay
  $ToPay = 0
  if((($ClientInfo.Order.TotalPrice | Measure-Object -Sum).Sum - $ClientInfo.Balance) -gt 0){ $ToPay = ($ClientInfo.Order.TotalPrice | Measure-Object -Sum).Sum - $ClientInfo.Balance }

  $document = (Get-Content $template).
      Replace('$DogovirNumber$',$ClientInfo.ContractNumber).
      Replace('$DogovirDate$',$ClientInfo.ContractDate).
      Replace('$ClientName$',$ClientInfo.Name).            
      Replace('$VOsobi$',$ClientInfo.Represented).            
      Replace('$EDRPOU$',$ClientInfo.EDRPOU). 
      Replace('$ClientBank$',$ClientInfo.BankRequisites).            
      Replace('$LegalAddress$',$ClientInfo.LegalAddress).            
      Replace('$PostAddress$',$ClientInfo.PostAddress).            
      Replace('$Phone$',$ClientInfo.Phone).            
      Replace('$ClientSignedBy$',$ClientInfo.SignedName).
      Replace('$ClientSignedPost$',$ClientInfo.SignedPost).
      Replace('$ClientOrder$',($ClientInfo.Order | %{ '<tr><td class="tb-bordered">'+$_.Name+'</td><td class="tb-bordered" style="text-align: center;">'+$_.Amount+'</td><td class="tb-bordered" style="text-align:center;">'+$_.Price+',00</td><td class="tb-bordered" style="text-align: center;">'+$_.TotalPrice+',00</td></tr>'}) -join "`r`n").
      Replace('$Total$',($ClientInfo.Order.TotalPrice | Measure-Object -Sum).Sum).
      Replace('$EURTotal$',(($ClientInfo.Order.TotalPrice | Measure-Object -Sum).Sum/$EURRate).ToString("#.#")).
      Replace('$EURRate$', $EURRate).
      Replace('$ActNumber$', $ClientInfo.AccrualNumber).
      Replace('$InvoiceBody$',($ClientInfo.Order | %{'<tr><td class="tb-bordered" style="text-align: center;">'+($i++)+'</td><td class="tb-bordered">'+$_.Name+'</td><td class="tb-bordered" style="text-align: right;" >'+$_.Amount+'</td><td class="tb-bordered" >послуга</td><td class="tb-bordered" style="text-align: right; white-space: nowrap;">'+$_.Price+',00</td><td class="tb-bordered" style="text-align: right; white-space: nowrap;">'+$_.TotalPrice+',00</td></tr>'}) -join "`r`n").
      Replace('$ActBody$',($ClientInfo.Accrual | %{'<tr><td class="tb-bordered" style="text-align: center;">'+($j++)+'</td><td class="tb-bordered">'+$_.Name+' (з '+('{0:dd.MM.yyyy}' -f $_.BilledFrom)+' по '+('{0:dd.MM.yyyy}' -f $_.BilledTo)+')</td><td class="tb-bordered" style="text-align: right;" >'+$_.Amount+'</td><td class="tb-bordered" >послуга</td><td class="tb-bordered" style="text-align: right; white-space: nowrap;">'+$_.Price+',00</td><td class="tb-bordered" style="text-align: right; white-space: nowrap;">'+$_.Rental+',00</td></tr>'}) -join "`r`n").
      Replace('$ActTotal$',($ClientInfo.Accrual.Rental | Measure-Object -Sum).Sum).
      Replace('$ActTotalInWords$',(Convert-NumberToText ($ClientInfo.Accrual.Rental | Measure-Object -Sum).Sum.ToString().Replace(".00","").Replace(",00",""))).
      Replace('$OrderItemsCount$',($ClientInfo.Order.TotalPrice | Measure-Object).Count).
      Replace('$InvoiceNumber$',$ClientInfo.InvoiceNumber).
      Replace('$SellerName$',$ClientInfo.SellerName).            
      Replace('$SellerVOsobi$',$ClientInfo.SellerRepresented).            
      Replace('$SellerEDRPOU$',$ClientInfo.SellerEDRPOU). 
      Replace('$SellerBank$',$ClientInfo.SellerBankRequisites).            
      Replace('$SellerBankAcc$',$SellerBankAcc).            
      Replace('$SellerBankDep$',$SellerBankDep).            
      Replace('$SellerBankMFO$',$SellerBankMFO).            
      Replace('$SellerLegalAddress$',$ClientInfo.SellerLegalAddress).            
      Replace('$SellerPostAddress$',$ClientInfo.SellerPostAddress).            
      Replace('$SellerPhone$',$ClientInfo.SellerPhone).            
      Replace('$SellerSignedBy$',$ClientInfo.SellerSignedName).
      Replace('$SellerSignedPost$',$ClientInfo.SellerSignedPost).
      Replace('$DocumentDate$',$DocumentDate).
      Replace('$Balance$', $ClientInfo.Balance.ToString().Replace(".00","").Replace(",00","")).
      Replace('$TotalInWords$',(Convert-NumberToText $ToPay.ToString().Replace(".00","").Replace(",00",""))).
      Replace('$ToPay$',"{0}" -f $ToPay.ToString().Replace(".00","").Replace(",00",""))

    if(Test-Path $OutputFileName) { Remove-Item $OutputFileName }
    if(Test-Path $OutputFileNameHTML) { Remove-Item $OutputFileNameHTML }
    $document | Out-File -Encoding utf8 -FilePath $OutputFileNameHTML
    $cmdFooter = ""
    $cmdHeader = ""
    if($Footer) { $cmdFooter = " --footer-html $('{0}_footer.html' -f $docPath)"}
    if($Header) { $cmdHeader = " --header-html $('{0}_header.html' -f $docPath)"}
    
    # вызов wkhtmltopdf для генерации ПДФ из хтмл
    $cmdStr = "wkhtmltopdf.exe{0}{1} {2} {3}" -f $cmdFooter,$cmdHeader,$OutputFileNameHTML,$OutputFileName
    cmd /c $cmdStr

    if(!$Debug){
        rm $OutputFileNameHTML -Force
    }
}

#
#  Convert-NumberToText - преобразить число в текст ( для денежных сумм). Пример 45 - сорок п'ять гривень
#  BUG:  не корректно работает на суммах с копейками!
#

function Convert-NumberToText ($x) {
  $e=@("одна","дві","три","чотири","п'ять","шість","сім","вісім","дев'ять")
  $d = @("десять","двадцать","тридцять","сорок","п'ятдесят","шістдесят","сімдесят","вісімдесят","дев'яносто")
  $t=@("одинадцять","дванадцять","тринадцять","чотирнадцять","п'ятнадцять","шістнадцять","сімнадцять","вісімнадцять","дев'ятнадцять")
  $h=@("сто","двісті","триста","чотириста","п'ятсот","шістсот","сімсот","вісімсот","дев'ятсот")

  # разбиваем на разряды 999
  $c = [Math]::floor(([string]$x).length/3)
  $a=@()
  if ($c -ne 0) {
      $a += 1..$c | %{ ([string]$x).Substring(([string]$x).Length-$_*3,3)} #массив чисел 0 - 0..999, 1 - 1000..999000, 2 - 1000000..999000000
  }
  if(([string]$x).length%3 -ne 0 ) { $a += ([string]$x).Substring(0,([string]$x).Length-$c*3) }

  $n2t = ""

  ($a.Length-1)..0 | % {
      $number = [string]([int]$a[$_])
      switch($number){
          {$number.Length -eq 1}                { $n2t += "$($e[[int]$number-1]) "}
          {$number.Length -eq 2 -and [int]$number -lt 20} { 
                                                         if([int]$number -ne 10) { $n2t += "$($t[[int]$number-11]) "} else { $n2t +="десять " }
                                                      }
          {$number.Length -eq 2 -and [int]$number -gt 19} { 
                                                        $n2t +="$($d[[math]::Floor([int]$number/10)-1]) " 
                                                       if([int]$number%10 -ne 0) { $n2t +="$($e[([int]$number%10)-1]) " }
                                                     }
          {$number.Length -eq 3}                { 
                                                       $n2t +="$($h[[math]::Floor([int]$number/100)-1]) " # сотни
                                                       $z = [int]$number%100                      # десятки 
                                                       if($z -ne 0) {               
                                                          if(([string]$z).Length -eq 2 -and $z -lt 20) { # 
                                                              if($z -ne 10) { $n2t +="$($t[$z-11]) " } else { $n2t +="десять " }
                                                          } else {
                                                              if (([string]$z).Length -eq 2 -and $z -gt 19) {
                                                                  $n2t +="$($d[[math]::floor($z/10)-1]) " 
                                                              }
                                                              if($z%10 -ne 0) { $n2t +="$($e[($z%10)-1]) " }       
                                                          }
                                                       }

                                                     }
      }
      switch($_){
          1 { 
              $last2digits = 0
              if(([string]$number).Length -eq 3) { $last2digits = [int]([string]$number).Substring(([string]$number).Length-2,2) } else { $last2digits=$number }
              switch($last2digits){
                  { $last2digits -ge 11 -and $last2digits -le 19 } { $n2t += "тисяч "; break; }
                  { $last2digits%10 -eq 1 } { $n2t += "тисяча " }
                  { $last2digits%10 -eq 2 } { $n2t += "тисячі " }
                  { $last2digits%10 -eq 3 } { $n2t += "тисячі " }
                  { $last2digits%10 -eq 4 } { $n2t += "тисячі " }
                  default {$n2t += "тисяч "}
              }          
            }
          2 { $n2t +="мільон " }
      }
  }
  if(([string]$x).Length -ge 3) { $last2digits = [int]([string]$x).Substring(([string]$x).Length-2,2) } else { $last2digits=$x }
  switch($last2digits){
      { $last2digits -ge 11 -and $last2digits -le 19 } { $n2t += "гривень"; break; }
      { $last2digits%10 -eq 1 } { $n2t += "гривня" }
      { $last2digits%10 -eq 2 } { $n2t += "гривні" }
      { $last2digits%10 -eq 3 } { $n2t += "гривні" }
      { $last2digits%10 -eq 4 } { $n2t += "гривні" }
      default {$n2t += "гривень"}
  }
  if([int]$x -eq 0) { return "нуль гривень"} else { return $n2t }
}

function Generate-MonthlyDocs ($CompanyId=$null,          # если нет ИД компании, получаем список всех активных. может принимать значение массива ИД. пример: Generate-MonthlyDocs 1,2,3 2019 5
                                $Year=$null,              # год и месяц - период за который сгенерировать доки (акт + счет), если пусто берем предыдущий период(месяц) - год и месяц
                                $Month=$null,
                                $SendTo=$null,            # для дебага - сгенерить и отправить документ на указанній имейл вместо клиентского из майклауда
                                $ExcludeIds=$null,        # исключить из рассілки следующие CompanyId (список, через запятую)    
                                [switch]$NoEmails=$false) # сгенерировать доки и не отправлять по почте 
{
    Import-Module SQLPS -DisableNameChecking
    . C:\Scripts\mycloud-db.ps1
    cd C:\gendoc\uploads # путь где будут сгенерированы ПДФки

    if(!$Year -or !$Month){
        $Year = (get-date).AddMonths(-1).Year
        $Month = (get-date).AddMonths(-1).Month
    }

    # на первое число след. месяца
    $d = (get-date(("01.{0:d2}.{1}" -f $Month,$Year))).AddMonths(1)
    $docdate="01.{0:d2}.{1}" -f $d.Month,$d.Year

    # если нет ИД компании, получаем список всех активных
    if(!$companyId){ 
        $companyId = (Invoke-Sqlcmd -ServerInstance .\SQLEXPRESS -Database cloud2_1_db -Query "select Id from Companies where IsActive=1").Id
    }
    
    $companyId | % {
        # get company Info 
        $id = $_

        $clientsJson = (Get-MyCloudCompany $Id -OrderDate (get-date($docdate)) -AccrualDate (get-date($docdate)).AddDays(-1))

        # gen documents 
        $attachments=@()
        $actname=$invname=""

        # INVOICE
        if($clientsJson[0].Order -or $clientsJson[0].Balance -lt 0){
            $clientsJson[0].InvoiceNumber = Get-MyCloudInvoiceId
            #if no AccrualNumber (new client) - set it to 0
            if(!$($clientsJson[0].AccrualNumber)) { $clientsJson[0].AccrualNumber=0 }
            $invname = $docname = "$($Id)_$($clientsJson[0].AccrualNumber)_$($docdate)_invoice.pdf"  # $clientsJson[0].AccrualNumber вместо $clientsJson[0].InvoiceNumber в названии файла, т.к. ссылки формируются в JS и там нет данных о номере счета (в БД)
            Generate-MyCloudPDF -ClientInfo  $clientsJson[0] -Template ..\templates\invoice.html -DocumentDate $docdate -OutputFilename $docname
            Invoke-Sqlcmd -ServerInstance .\SQLEXPRESS -Database cloud2_1_db -Query "insert into EventsLog (Date,EventLog,CreaterId,CompanyId) values(getdate(),'Выставлен счет № $($clientsJson[0].InvoiceNumber) от $docdate','5ccac4ca-0806-4789-8579-e7eeebc08407',$Id)" 
            $attachments+=$invname
        }

        # ACT
        if($clientsJson[0].Accrual){
            $actname = $docname = "$($Id)_$($clientsJson[0].AccrualNumber)_$($docdate)_act.pdf"
            $actdate = ('{0:dd.MM.yyyy}' -f (get-date($docdate)).AddDays(-1))
            Generate-MyCloudPDF -ClientInfo  $clientsJson[0] -Template ..\templates\act.html -DocumentDate $actdate -OutputFilename $docname
            Invoke-Sqlcmd -ServerInstance .\SQLEXPRESS -Database cloud2_1_db -Query "insert into EventsLog (Date,EventLog,CreaterId,CompanyId) values(getdate(),'Сформирован Акт № $($clientsJson[0].AccrualNumber) за $Month-$Year','5ccac4ca-0806-4789-8579-e7eeebc08407',$Id)" 
            if($clientsJson[0].InvoiceNumber) {
                Invoke-Sqlcmd -ServerInstance .\SQLEXPRESS -Database cloud2_1_db -Query "update Accruals set BillingId=$($clientsJson[0].InvoiceNumber),BillingSended='$("{0:yyyy-MM-dd}" -f (get-date($docdate)))' where id=$($clientsJson[0].AccrualNumber)" 
            } else { # if no InvoiceNumber - invoice wasn't sent, only Act (closed client)
                Invoke-Sqlcmd -ServerInstance .\SQLEXPRESS -Database cloud2_1_db -Query "update Accruals set BillingId=0,BillingSended='$("{0:yyyy-MM-dd}" -f (get-date($docdate)))' where id=$($clientsJson[0].AccrualNumber)" 
            }
            $attachments+=$actname
        }

        # exclude Company
        if ($ExcludeIds -notcontains $id -and $NoEmails -eq $false) { 
        
            # SEND E-MAIL
            $smtp_user = "sender@domain"
            $smtpServer = "smtp.server"  
            $smtp_pass = ("YourSecurePath" | ConvertTo-SecureString -AsPlainText -Force)
            $smtpCred = New-Object –TypeName System.Management.Automation.PSCredential –ArgumentList $smtp_user, $smtp_pass

            $from = "sender@domain"  
            $to = ([string]$clientsJson[0].Emails).Split(",")
            if($SendTo){
                $to = $SendTo # override with another email(s)
            }
            $cc = $from

            $subject = "ОБЛАКО - Cчет на оплату услуг № $($clientsJson[0].InvoiceNumber) от $docdate для $($clientsJson[0].Name)"

            $clientEmailBlock = ""

            if($clientsJson[0].Tags.Name -contains "оферта" -and $clientsJson[0].Balance -le 0) {
                $clientEmailBlock = "<p>Напоминаем, что, согласно п.4.19 <a href='https://drive.zimbra.in.ua/s/9g2goN4FHPgTetw#pdfviewer' target='_blank'>Договора публичной оферты</a>, предоплату за текущий месяц необходимо внести по 7 число включительно.</p>"
            }
            if($clientsJson[0].Tags.Name -contains "договор2" -and $clientsJson[0].Balance -le 0) {
                $clientEmailBlock = "<p>Напоминаем, что, согласно п.3.19 Договора на оказание услуг, предоплату за текущий месяц необходимо внести по 7 число включительно.</p>"
            }
            if($clientsJson[0].Tags.Name -contains "договор1" -and $clientsJson[0].Balance -lt 0) {
                $clientEmailBlock = "<p>Напоминаем, что, согласно условиям взаиморасчётов по Вашему Договору, оплату за потребленные услуги необходимо внести по 20 число включительно.</p>"
            }

            $invoicetext = ""
            $accrualtext = ""
            
            if($clientsJson[0].InvoiceNumber) {
                $invoicetext = "- актуальный Cчет на оплату услуг № $($clientsJson[0].InvoiceNumber) от $docdate<br/>"
            }
            if($clientsJson[0].Accrual) {
                $accrualtext = "- информацию по списанию за предыдущий месяц в Акте № $($clientsJson[0].AccrualNumber) за $Year-$("{0:d2}" -f $Month)"
            }
                   
            $to | % { 

                $messageId = [System.Guid]::NewGuid()
                $body = "<p>Здравствуйте, Уважаемый клиент!</p>

                        <p>В приложении к письму Вы найдете:<br/> 
                        $invoicetext
                        $accrualtext</p>

                        <p>Посмотреть Ваш баланс и оплатить онлайн с помощью банковской карты можно из личного кабинета https://my.cloud.net.ua</p>
                        
                        $clientEmailBlock

                        <p>С уважением, <br/>
                        Сервис провайдер Облако <br/>
                        +38 (044) 363-93-94<br/>
                        +38 (050) 383-93-94<br/>
                        +38 (068) 383-93-94<br/>
                        <a href='https://oblako.zendesk.com' target='_blank'>Вопросы и ответы</a> | <a href='mailto:support@cloud.net.ua' target='_blank'>Написать в поддержку</a> | <a href='https://cloud.net.ua' target='_blank'>Сайт</a> | <a href='https://facebook.com/cloudnetua' target='_blank'>Фейсбук</a></p>
                        <img src='https://www.cloud.net.ua/img/logo.png' alt='[OBLAKO_LOGO]' />"

                if($attachments.Count -gt 0){
                    Send-MailMessage -SmtpServer $smtpServer -From $from -To $_ -BCc $cc -Subject $subject -Body $body -BodyAsHtml -UseSsl -Credential $smtpCred -Port 587 -Encoding UTF8 -Attachments $attachments 
                }
            }

        } else { Write-Host "No email sent to $($clientsJson[0].Name)"}
    }
}
