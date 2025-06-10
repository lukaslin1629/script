# Description: This script processes a xlsx file containing serial numbers, retrieves product information from Dell (via Playwright), HP (via WebRequest), and Lenovo (via Warranty API)

# First time use:
# Install-Module -Name ImportExcel -Scope CurrentUser
# Install Playwright (global):
# npm install -g playwright
# playwright install

param (
    [string]$excelname = "assets"
)

Import-Module ImportExcel

$data = Import-Excel -Path "$PSScriptRoot\$excelname.xlsx"

$total = $data.Count
$count = 0

foreach ($row in $data) {
    $count++
    $serial = $row.'Serial'
    Write-Output "[$count/$total] Processing serial: $serial"

    if ([string]::IsNullOrWhiteSpace($serial)) {
        Write-Output "Skipping empty serial on row $count"
        continue
    }

    $foundModel = $false

    try {
        $modelName = & node "$PSScriptRoot\playwright-dell.js" $serial | Out-String
        $modelName = $modelName.Trim()

        if ($modelName -match "Latitude\s+\d+|Precision\s+\d+|OptiPlex\s+\d+|XPS\s+\d+|Inspiron\s+\d+") {
            Write-Output "Matched Dell model: $modelName"

            $row.Make = "Dell"
            $row.Model = $modelName

            if ($row.Model -match "Latitude|Precision|XPS") {
                $row.Type = "Laptop"
            }
            elseif ($row.Model -match "OptiPlex") {
                $row.Type = "Desktop"
            }
            else {
                $row.Type = "Unknown"
            }

            $foundModel = $true
        } else {
            Write-Output "Dell model not found for serial: $serial"
        }
    } catch {
        Write-Output "Dell lookup failed for serial: $serial"
    }

    if (-not $foundModel) {
        Write-Output "Trying HP fallback for serial: $serial"

        try {
            $hpUrl = "https://support.hp.com/us-en/product-warranty-results?serialnumber=$serial"
            $hpResponse = Invoke-WebRequest -Uri $hpUrl
            Start-Sleep -Seconds 1

            $hpH1s = $hpResponse.ParsedHtml.getElementsByTagName("h1")

            foreach ($hpH1 in $hpH1s) {
                if ($hpH1.innerText -match "HP\s+\w+\s+\d+|HP\s+\w+\s+\d+\s+Notebook") {
                    $hpModelName = $hpH1.innerText
                    Write-Output "Matched HP model: $hpModelName"

                    $row.Make = "HP"
                    $row.Model = $hpModelName

                    if ($row.Model -match "EliteBook|ProBook|Notebook") {
                        $row.Type = "Laptop"
                    }
                    elseif ($row.Model -match "Desktop|Tower") {
                        $row.Type = "Desktop"
                    }
                    else {
                        $row.Type = "Unknown"
                    }

                    $foundModel = $true
                    break
                }
            }
        } catch {
            Write-Output "HP lookup failed for serial: $serial"
        }
    }

    if (-not $foundModel) {
        Write-Output "Trying Lenovo fallback (Warranty API) for serial: $serial"

        try {
            $cleanSerial = ($serial -replace '-', '').ToUpper()

            $body = @{
                CountryCode = "US"
                SerialNumber = $cleanSerial
                MachineTypeModel = ""
                ProductId = ""
                ShipToCountryCode = ""
            } | ConvertTo-Json -Depth 10

            $response = Invoke-WebRequest -Uri "https://pcsupport.lenovo.com/services/rest/api/v2/products?serialNumber=$cleanSerial" `
                -Method POST `
                -Headers @{
                    "User-Agent" = "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"
                }

            $json = $response.Content | ConvertFrom-Json
            $modelName = $json.ProductName

            if ($modelName -match "ThinkPad|ThinkCentre|IdeaPad|Legion|Yoga") {
                Write-Output "Matched Lenovo model (API): $modelName"

                $row.Make = "Lenovo"
                $row.Model = $modelName

                if ($row.Model -match "ThinkPad|IdeaPad|Yoga") {
                    $row.Type = "Laptop"
                }
                elseif ($row.Model -match "ThinkCentre|Legion|Desktop") {
                    $row.Type = "Desktop"
                }
                else {
                    $row.Type = "Unknown"
                }

                $foundModel = $true
            } else {
                Write-Output "Lenovo model not found (API) for serial: $serial (cleaned: $cleanSerial)"
            }
        } catch {
            Write-Output "Lenovo API lookup failed for serial: $serial"
        }
    }
}

$data | Export-Excel -Path "$PSScriptRoot\updated.xlsx"

Write-Output "Serials processed"