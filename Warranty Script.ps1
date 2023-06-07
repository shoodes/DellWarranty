# Load assembly
Add-Type -AssemblyName System.Windows.Forms

# Create a new form
$form = New-Object System.Windows.Forms.Form
$form.Text = 'Dell Warranty API'

# Create a label for the service tag field
$label1 = New-Object System.Windows.Forms.Label
$label1.Location = New-Object System.Drawing.Point(10, 20)
$label1.Size = New-Object System.Drawing.Size(280, 20)
$label1.Text = 'Enter ServiceTag: (XXXXX, XXXXX, XXXXX)'
$form.Controls.Add($label1)

# Create a text box for the service tag field
$serviceTagBox = New-Object System.Windows.Forms.TextBox
$serviceTagBox.Location = New-Object System.Drawing.Point(10, 40)
$serviceTagBox.Size = New-Object System.Drawing.Size(260, 20)
$form.Controls.Add($serviceTagBox)

# Create a label for the API Key field
$label2 = New-Object System.Windows.Forms.Label
$label2.Location = New-Object System.Drawing.Point(10, 70)
$label2.Size = New-Object System.Drawing.Size(280, 20)
$label2.Text = 'Enter API Key: (N/A)'
$form.Controls.Add($label2)

# Create a text box for the API Key field
$apiKeyBox = New-Object System.Windows.Forms.TextBox
$apiKeyBox.Location = New-Object System.Drawing.Point(10, 90)
$apiKeyBox.Size = New-Object System.Drawing.Size(260, 20)
$form.Controls.Add($apiKeyBox)

# Create a label for the Key Secret field
$label3 = New-Object System.Windows.Forms.Label
$label3.Location = New-Object System.Drawing.Point(10, 120)
$label3.Size = New-Object System.Drawing.Size(280, 20)
$label3.Text = 'Enter Key Secret: (N/A)'
$form.Controls.Add($label3)

# Create a text box for the Key Secret field
$keySecretBox = New-Object System.Windows.Forms.TextBox
$keySecretBox.Location = New-Object System.Drawing.Point(10, 140)
$keySecretBox.Size = New-Object System.Drawing.Size(260, 20)
$form.Controls.Add($keySecretBox)

# Create an OK button
$okButton = New-Object System.Windows.Forms.Button
$okButton.Location = New-Object System.Drawing.Point(10, 170)
$okButton.Size = New-Object System.Drawing.Size(75, 23)
$okButton.Text = 'OK'
$okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
$form.AcceptButton = $okButton
$form.Controls.Add($okButton)

$form.Topmost = $true

$result = $form.ShowDialog()

if ($result -eq [System.Windows.Forms.DialogResult]::OK)
{
    $ServiceTags = $serviceTagBox.Text
    $ApiKey = $apiKeyBox.Text
    $KeySecret = $keySecretBox.Text
}


$ORG_API_Key="YOUR_API_KEY_HERE"
$ORG_API_Secret="YOUR_API_SECRET_HERE"

$bios = gwmi win32_bios
$st = $bios.SerialNumber


[String]$servicetags = $ServiceTags -join ", "

$headers = @{"Accept" = "application/json" }
$headers.Add("Authorization", "Bearer $token")

$params = @{ }
$params = @{servicetags = $servicetags; Method = "GET" }

# Create an array to store the outputs for exporting
$OutputArray = @()

Try {
	$Global:response = Invoke-RestMethod -Uri "https://apigtwb2c.us.dell.com/PROD/sbil/eapi/v5/asset-entitlements" -Headers $headers -Body $params -Method Get -ContentType "application/json"
}
Catch {
	$AuthURI = "https://apigtwb2c.us.dell.com/auth/oauth/v2/token"
	$OAuth = "$ApiKey`:$KeySecret"
	$Bytes = [System.Text.Encoding]::ASCII.GetBytes($OAuth)
	$EncodedOAuth = [Convert]::ToBase64String($Bytes)
	$Headers = @{ }
	$Headers.Add("authorization", "Basic $EncodedOAuth")
	$Authbody = 'grant_type=client_credentials'
	[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
	$AuthResult = Invoke-RESTMethod -Method Post -Uri $AuthURI -Body $AuthBody -Headers $Headers
	$Global:token = $AuthResult.access_token
	$Global:response = Invoke-RestMethod -Uri "https://apigtwb2c.us.dell.com/PROD/sbil/eapi/v5/asset-entitlements" -Headers $headers -Body $params -Method Get -ContentType "application/json"
}
Finally {
	foreach ($Record in $response) {
		$servicetag = $Record.servicetag
		$Json = $Record | ConvertTo-Json
		$Record = $Json | ConvertFrom-Json 
		$Device = $Record.productLineDescription
		$ShipDate = $Record.shipDate
		$EndDate = ($Record.entitlements | Select -Last 1).endDate
		$Support = ($Record.entitlements | Select -Last  1).serviceLevelDescription
		$ShipDate = $ShipDate | Get-Date -f "MM-dd-y" #Invalid ST
		$EndDate = $EndDate | Get-Date -f "MM-dd-y" #Invalid ST
		$today = get-date
		$type = $Record.ProductID

		if ($type -Like '*desktop') { 
			$type = 'desktop'        
		}
		elseif ($type -Like '*laptop') { 
			$type = 'laptop'
		}

		Write-Host -ForegroundColor White -BackgroundColor "DarkRed" $Computer
		Write-Host "Service Tag   : $servicetag"
		Write-Host "Model Name    : $Device"
		Write-Host "Model Type    : $type"
		Write-Host "Shipped Date  : $ShipDate"
		if ($today -ge $EndDate) {
			Write-Host -NoNewLine "Warranty Exp. : $EndDate  "; 
			Write-Host -ForegroundColor "RED" "[WARRANTY HAS EXPIRED]"
		}
		else {
			Write-Host "Warranty Exp. : $EndDate" 
		} 

		$ServiceLevel = ""
		foreach ($Item in ($($Record.entitlements.serviceLevelDescription | select -Unique | Sort-Object -Descending))) {
			$ServiceLevel += "$Item, "
		}

		$ServiceLevel = $ServiceLevel.TrimEnd(', ')

		# Create an object and add it to the output array
		$OutputObject = New-Object -TypeName PSObject
		$OutputObject | Add-Member -MemberType NoteProperty -Name "Service Tag" -Value $servicetag
		$OutputObject | Add-Member -MemberType NoteProperty -Name "Model" -Value $Device
		$OutputObject | Add-Member -MemberType NoteProperty -Name "Type" -Value $type
		$OutputObject | Add-Member -MemberType NoteProperty -Name "Ship Date" -Value $ShipDate
		$OutputObject | Add-Member -MemberType NoteProperty -Name "Warranty Exp." -Value $EndDate
		$OutputObject | Add-Member -MemberType NoteProperty -Name "Service Level" -Value $ServiceLevel
		$OutputArray += $OutputObject
	}

	# Ask user if they want to export to a .csv file
	$exportToCsv = Read-Host -Prompt 'Would you like to export the data to a .csv file? (y/n) '
	if ($exportToCsv -eq 'y') {
		$csvPath = Read-Host -Prompt 'Please provide the full path where you would like to save the .csv file. EX: C:\temp\output.csv '
		$OutputArray | Export-Csv -Path $csvPath -NoTypeInformation
		Write-Host "Data exported to $csvPath"
	}
}
