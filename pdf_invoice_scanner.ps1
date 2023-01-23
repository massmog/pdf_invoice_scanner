$folder = "D:\Documents\Verträge\Handy Österreich HoT\Rechnungen"
$filetype = "*.pdf"
$files = Join-Path -Path $folder -ChildPath $filetype
$output_folder = "D:\Documents\Verträge\Handy Österreich HoT\Rechnungen\Sortiert"
function convert-date {
    Param ($raw_date)
    $converted_date = Get-Date -Year $raw_date[3] -Month $raw_date[2] -Day $raw_date[1] -Hour 0 -Minute 0 -Second 0
    return $converted_date
}

Get-ChildItem $files | ForEach-Object {
    $invoice = $null
    $matches = $null

    $invoice_date = $null
    $doc_class = $null
    $doc_description = $null
    $doc_name = $null

    $isVorsteuerbescheinigung = $false
    $isTicket = $false
    $isInvoice = $false
    $isDeliveryNote = $false
    $isEinzelgespraechsnachweis = $false

    pdftotext.exe -enc UTF-8 $_.FullName temp.txt
    write-host "File: `t`t$($_.Name)"
    $invoice = (Get-Content temp.txt -Encoding UTF8 -Raw) -replace '\s{2,}', ' '
   
    # Find dates
    $invoice_date = [DateTime] "01/01/1970"
    if ($invoice -match '(\d{1,2}) (Januar|Februar|März|April|Mai|Juni|Juli|August|September|Oktober|November|Dezember) (20\d{2})') {
        $raw_invoice_date = $matches
        $date_dict = @{Januar = 1; Februar = 2; März = 3; April = 4; Mai = 5; Juni = 6; Juli = 7; August = 8; September = 9; Oktober = 10; November = 11; Dezember = 12 }
        $raw_invoice_date[2] = $date_dict[$raw_invoice_date[2]]
        $temp_date = convert-date($raw_invoice_date)
        if($temp_date -gt $invoice_date){
            $invoice_date = $temp_date
        }
    }
    if ($invoice -match '(\d{2})\.(01|02|03|04|05|06|07|08|09|10|11|12)\.(20\d{2})') {
        $temp_date = convert-date($matches)
        if($temp_date -gt $invoice_date){
            $invoice_date = $temp_date
        }
    }
    if ($invoice -match '(\d{2})\. (Jan|Feb|Mär|Apr|Mai|Jun|Jul|Aug|Sep|Okt|Nov|Dez) (20\d{2})') {
        $raw_invoice_date = $matches
        $date_dict = @{Jan = 1; Feb = 2; Mär = 3; Apr = 4; Mai = 5; Jun = 6; Jul = 7; Aug = 8; Sep = 9; Okt = 10; Nov = 11; Dez = 12 }
        $raw_invoice_date[2] = $date_dict[$raw_invoice_date[2]]
        $temp_date = convert-date($raw_invoice_date)
        if($temp_date -gt $invoice_date){
            $invoice_date = $temp_date
        }
    }
    if ($invoice -match '(\d{2})\.(01|02|03|04|05|06|07|08|09|10|11|12)\.(\d{2})') {
        $raw_invoice_date = $matches
        $raw_invoice_date[3] = [int]($raw_invoice_date[3]) + 2000
        $temp_date = convert-date($raw_invoice_date)
        if($temp_date -gt $invoice_date){
            $invoice_date = $temp_date
        }
    }

    # Find vendors
    if ($invoice -cmatch '(?:Amazon|Conrad|Vettore|ÖBB|DB|Hotel|Jusline|ACP|HoT)') {
        $vendor = $matches[0]

        $doc_description = $vendor 

        if ($vendor -eq 'Amazon') {
            # Find Amazon Rechnungsnummer
            if ($invoice -cmatch 'AT[A-Z0-9]{10}') {
                $amazon_order_no = $matches[0]
                $doc_description = $doc_description+"_"+$amazon_order_no
             }
        }
        elseif ($vendor -eq 'ÖBB') {
            # Find ÖBB Buchungsnr.
            if ($invoice -cmatch '(?:\d{4} \d{4} \d{4} \d{4})') {
                $oebb_ticket_id = ($matches[0] -replace ' ', '')
                $doc_description = $doc_description+"_"+$oebb_ticket_id
            }
            $isVorsteuerbescheinigung = $invoice -match 'Vorsteuerbescheinigung'
        }
    }

    # Classify document
    $isTicket = $invoice -match 'Ticket'
    $isInvoice = $invoice -match '(?:Rechnung|Invoice)'
    $isDeliveryNote = $invoice -match 'Lieferschein'
    $isEinzelgespraechsnachweis = $invoice -match 'Einzelgesprächsnachweis'
    $isKontoauszug = $invoice -match 'Kontoauszug'

    # Determine doc_class
    if ($isTicket -and $isVorsteuerbescheinigung) {
        $doc_class = "VSt_Bescheinigung"
    }
    elseif ($isTicket) {
        $doc_class = "Ticket"
    }    
    elseif ($isKontoauszug){
        $doc_class = "Kontoauszug"
    }
    elseif ($isInvoice) {
        $doc_class = "Rechnung"
    }
    elseif ($isDeliveryNote){
        $doc_class = "Lieferschein"
    }
    elseif ($isEinzelgespraechsnachweis){
        $doc_class = "Einzelgesprächsnachweis"
    }

    # Create document name
    $first_string = $true 
    
    if($invoice_date -ne $null)
    {
        if($first_string){
            $doc_name = $invoice_date.ToString('yyMMdd')
            $first_string = $false
        }
        else{
            $doc_name = $doc_name+"_"+$invoice_date.ToString('yyMMdd')
        }
    }

    if($doc_class -ne $null)
    {
        if($first_string){
            $doc_name = $doc_class
            $first_string = $false
        }
        else{
            $doc_name = $doc_name+"_"+$doc_class
        }
    }

    if($doc_description -ne $null)
    {
        if($first_string){
            $doc_name = $doc_description
            $first_string = $false
        }
        else{
            $doc_name = $doc_name+"_"+$doc_description
        }
    }


    # # Find UIDs
    # $tax_id = $null
    # if ($invoice -cmatch 'DE\d{9}') {
    #     $tax_id = $matches[0]
    # }
    # elseif ($invoice -cmatch 'ATU\d{8}') {
    #     $tax_id = $matches[0]
    # }
    # elseif ($invoice -cmatch 'LU\d{8}') {
    #     $tax_id = $matches[0]
    # }
    # elseif ($invoice -cmatch 'FR[a-zA-Z0-9]{2}\d{9}') {
    #     $tax_id = $matches[0]
    # }
    # elseif ($invoice -cmatch 'EL\d{9}') {
    #     $tax_id = $matches[0]
    # }

    # Write documents to output folder
    # Create output folder
    if (-not (Test-Path -Path $output_folder)){
        $new_folder_info = New-Item -ItemType Directory -Path $output_folder
        Write-Host "Output folder created."
    }
    # Find free filename
    $doc_name_ext = $doc_name+$_.Extension
    $new_file_path = Join-Path -Path $output_folder -ChildPath $doc_name_ext
    
    for ($file_no = 2;Test-Path -Path $new_file_path;$file_no++) {
        $doc_name_ext = $doc_name+"_"+$file_no+$_.Extension
        $new_file_path = Join-Path -Path $output_folder -ChildPath $doc_name_ext
    }
    write-host "New filename: `t$doc_name_ext"
    
    Copy-Item $_.FullName -Destination $new_file_path
    write-host "Copied to: `t$new_file_path"

    Write-Host "------------------------"
    # Clean up
    Remove-Item .\temp.txt    
}


