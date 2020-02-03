# pdfgen - генерация PDF счета и акта из HTML шаблона с помощью PowerShell и wkhtmltopdf

## How it works
Получив из своей БД/СРМ/учетной системы данные о заказах и суммах, а также реквизиты клиентов, в виде структуры JSON, передаем данные в функцию Generate-PDF предварительно сконвертировав JSON в PowerShell PSObject. 
Функция Generate-PDF читает указанный при вызове шаблон и заменяет в нем параметры на значения из структуры данных и сохраняет в промежуточном HTML, который потом с помощью утилиты wkhtmltopdf конвертирует в PDF файл
```JSON
var json = [{
        // данные клиента
        Id : "",
        EDRPOU : "",
        Name : "",
        LegalAddress : "",
        PostAddress : "",
        Phone : "",
        BankRequisites : "р/р 26005052634833 в АТ КБ ПРИВАТБАНК МФО 300711",
        ContractNumber : "",
        ContractDate : "",
        SignedName : "ФИО",                       // подписант клиента
        SignedPost : "Директор",                  // должность подписанта клиента
        Represented  :  "директора Иванова П.С.", // в лице  для акта

        InvoiceNumber : "",
        
        // массив строк заказа, для счета на предоплату на текущий месяц
        Order : [{
          Name : "Оренда серверу.." ,    // название позиции в заказе 
          Amount : 0,                    // кол-во
          Price : 100,                   // цена за ед
          TotalPrice : 100              // общая сумма (с учетом скидок и тп)
        },
        {
          ...
        }],
        
        ActNumber = "",

        // массив строк для Акта - за что начислено абонплату за предідущий период
        Accrual : [{
          Name : " Оренда серверу.."     // название позиции в заказе,
          Amount : 0                     // кол-во
          Price : 100                    // цена за ед
          Rental : 100                   // общая сумма (с учетом скидок и тп)
          BilledFrom : "01.01.2020"
          BilledTo : "31.01.2020"
        },
        {
            ...
        }],
        
        // данные продавца
        SellerName : "Продавець",
        SellerVOsobi : "директора Продавця Ф.О.",
        SellerEDRPOU : "" ,
        SellerBank : "",
        SellerBankAcc : "",
        SellerBankDep : "",
        SellerBankMFO : "",
        SellerLegalAddress : "",
        SellerPostAddress : "",
        SellerPhone : "",
        SellerSignedBy : "Продавець Ф.О.",
        SellerSignedPost : "директор"
    },
    {
       ...
    }]
```
```PowerShell
$json = get-your-clients-from-db-and-return-json-array
$x = ConvertFrom-Json $json -Depth 3

# OR create PSObject directly

$x = New-Object PSObject -Property @{
        # данные клиента
        Id = $Id
        EDRPOU = $EDRPOU
        Name = $Name
        LegalAddress = $RegisteredAddress
        PostAddress = $AddressForCorrespondence
        Phone = $ContactPhoneNumber
        BankRequisites = $BankRequisites
        ContractNumber = $ContractNumber
        ContractDate = $ContractDate
        SignedName = $SignedName   # подписал  - "ФИО"
        SignedPost = $SignedPost   # должность - "Директор"
        Represented  = $Represented  # в лице "директора Иванова П.С." - для акта

        InvoiceNumber = $InvoiceNumber
        
        # массив строк заказа, для счета на предоплату на текущий месяц
        Order = new-object PSObject -Property @{
          Name = $OrderItemName     # название позиции в заказе, Оренда серверу..
          Amount = $Amount          # кол-во
          Price = $Price            # цена за ед
          TotalPrice = $TotalPrice  # общая сумма (с учетом скидок и тп)
        } 
        
        ActNumber = $ActNumber

        # массив строк для Акта - за что начислено абонплату за предідущий период
        Accrual = new-object PSObject -Property @{
          Name = $OrderItemName     # название позиции в заказе, Оренда серверу..
          Amount = $Amount          # кол-во
          Price = $Price            # цена за ед
          TotalPrice = $TotalPrice  # общая сумма (с учетом скидок и тп)
        }
        # данные продавца
        $SellerName = $SellerName
        $SellerVOsobi = $SellerRepresented
        $SellerEDRPOU = $SellerEDRPOU 
        $SellerBank = $SellerBankRequisites
        $SellerBankAcc = $SellerBankAcc
        $SellerBankDep = $SellerBankDep
        $SellerBankMFO = $SellerBankMFO
        $SellerLegalAddress = $SellerLegalAddress
        $SellerPostAddress = $SellerPostAddress
        $SellerPhone = $SellerPhone
        $SellerSignedBy = $SellerSignedName
        $SellerSignedPost = $SellerSignedPost
    }
```
## Setup
- установить wkhtmltopdf ( загрузка аткуальной версии отсюда https://wkhtmltopdf.org/ )
- скопировать в папку, например c:\gendoc\templates, файлы шаблона act.html invoice.html и файл скрипта generate-pdf.ps1
- изменить файл generate-pdf.ps1 - установить рабочий каталог куда складываются файлы PDF и обновить параметры подключения к SMTP серверу
