# Anthony Rizzo 2023
# Last Modified: 11/28/2023

Add-Type -AssemblyName System.Web

$printServer = Read-Host "Provide the name of a print server"

if ((Test-Connection $printServer -Count 1 -Quiet) -eq $FALSE) {
    Write-Output "Failed connecting to print server."
    return
}

try {
    $printers = Get-Printer -ComputerName $printServer
    
    $object = @()
    $ipRegex = "\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}"
    foreach ($printer in $printers) {
        Write-Output "Processing $($printer.Name)..."
        $ipAddr = $printer.PortName | Select-String -Pattern $ipRegex -AllMatches | ForEach-Object {$_.Matches} | ForEach-Object {$_.Value}
        
        # Test Connection
        $online = $FALSE
        If (Test-Connection $printer.PortName -Count 1 -Quiet) {
            $online = $TRUE
        }
        
        $object += [PSCustomObject]@{
            Name = $printer.Name;
            Address = "<a href='http://$($ipAddr)' target='_blank' rel='noopener'>$($ipAddr)</a>";
            Online = (&{if($online) { "Online" } else { "Offline" }});
            Driver = $printer.DriverName
        }
    }
} catch {
    Write-Output "Something went wrong..."
    return
}
    
$documentGeneratedOn = Get-Date

$header = @"
<style>
h1, h5, th { text-align: center; color:#0046c3; }
table { margin: auto; font-family: Segoe UI; box-shadow: 10px 10px 5px #888; border: thin ridge grey; }
th { background: #0046c3; color: #fff; max-width: 400px; padding: 5px 10px; }
td { font-size: 1rem; padding: 5px 20px; color: #000; }
tr { background: #b8d1f3; }
tr:nth-child(even) { background: #dae5f4; }
tr:nth-child(odd) { background: #b8d1f3; }
</style>
<title>
Print Server Report
</title>
"@

$body = "<p>Document Generated On: $($documentGeneratedOn)</p>" + ($object | ConvertTo-HTML -Fragment)

$html = $object | ConvertTo-HTML -Head $header -Body $body | ForEach-Object {
    $PSItem -replace "<td>Offline</td>", "<td style='background-color:#FF0000;'>Offline</td>"
}

$outFile = "C:\PrinterReport\$($printServer)_printerreport.html"
New-Item $outFile -Force | Out-Null
[System.Web.HttpUtility]::HtmlDecode($html) | Out-File $outFile
Invoke-Item $outFile