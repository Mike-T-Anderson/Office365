
$DefaultApplicationName = "LogicMonitor Office365 monitoring"
$DefaultOffice365Domain = ""
$DefaulLMTenantName = ""
$DefaultLMApiId = ''
$DefaultLMApiKey = ''
$DefaultLMDeviceName = ""

$LogFile = "$(Get-Location)" + "\Log.txt"
$DataSourceO365File   = "$(Get-Location)" + "\Office365_app_status.xml"
$DataSourceGraphFile1 = "$(Get-Location)" + "\Exchange_email_clients.xml"
$DataSourceGraphFile2 = "$(Get-Location)" + "\Exchange_email_stats.xml"
$DataSourceGraphFile3 = "$(Get-Location)" + "\Exchange_mailbox_stats.xml"

$LogObjectReference = [Ref]""

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
Add-Type -AssemblyName System.Windows.Forms

Function Write-Log ($Message)
{
    $CurrentTime = Get-Date
    #Write-Host "[$($CurrentTime)] $Message"
    $LogObjectReference.Value.Text += "[$($CurrentTime)] $Message `r`n"
    "[$($CurrentTime)] $Message" | Out-File -FilePath $LogFile -Append
}

Function Show-OAuthWindow
{
    param(
        [System.Uri]$Url
    )

    Write-Log "Generating Oauth Form"
 
    $Form = New-Object -TypeName System.Windows.Forms.Form -Property @{Width=440;Height=640}
    $Web  = New-Object -TypeName System.Windows.Forms.WebBrowser -Property @{Width=420;Height=600;Url=($url)}
    $DocComp  = {
        $Global:uri = $Web.Url.AbsoluteUri
        if ($Global:Uri -match "error=[^&]*|code=[^&]*") {$Form.Close() }
    }
    $Web.ScriptErrorsSuppressed = $true
    $Web.Add_DocumentCompleted($DocComp)

    $Form.Controls.Add($Web)
    $Form.Add_Shown({$Form.Activate()})
    $Form.ShowDialog() | Out-Null

    $QueryOutput = [System.Web.HttpUtility]::ParseQueryString($Web.Url.Query)
    $Output = @{}
    foreach($Key in $QueryOutput.Keys)
    {
        $Output["$Key"] = $QueryOutput[$Key]
    }
    
    return $Output
}

Function Get-LmDevice ($Tenant, $AccessID, $AccessKey, $DeviceName)
{
    $ResourcePath = '/device/devices'
    $ReturnFilter = "?filter=displayName:$DeviceName"

    $Url = 'https://' + $Tenant + '.logicmonitor.com/santaba/rest' + $ResourcePath + $ReturnFilter

    $epoch = [Math]::Round((New-TimeSpan -start (Get-Date -Date "1/1/1970") -end (Get-Date).ToUniversalTime()).TotalMilliseconds)
    $hmac = New-Object System.Security.Cryptography.HMACSHA256
    $hmac.Key = [Text.Encoding]::UTF8.GetBytes($AccessKey)

    $httpVerb = 'GET'
    $requestVars = $httpVerb + $epoch + $ResourcePath
    $signatureBytes = $hmac.ComputeHash([Text.Encoding]::UTF8.GetBytes($requestVars))
    $signatureHex = [System.BitConverter]::ToString($signatureBytes) -replace '-'
    $signature = [System.Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($signatureHex.ToLower()))
    $auth = 'LMv1 ' + $AccessId + ':' + $signature + ':' + $epoch
    $headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
    $headers.Add("Authorization", $auth)
    $headers.Add("Content-Type", 'application/json')

    Write-Log -Message "Obtaining list of devices"
    Write-Log -Message "API call: $Url"
    $Response = Invoke-RestMethod -Uri $Url -Method Get -Header $headers
    Write-Log -Message "Response code is $($Response.status)"

    return $Response
}

Function Update-LMDeviceProperties ($Tenant, $AccessID, $AccessKey, $DeviceId, $PropertiesObject)
{
    $ResourcePath = "/device/devices/$DeviceId"
    $QueryParams = '?patchFields=customProperties&opType=replace'

    $WebDataObject = New-Object System.Object
    $WebDataObject | Add-Member -MemberType NoteProperty -Name "customProperties" -Value $PropertiesObject
    $Data = $WebDataObject | ConvertTo-Json

    $Url = 'https://' + $Tenant + '.logicmonitor.com/santaba/rest' + $ResourcePath + $QueryParams

    $epoch = [Math]::Round((New-TimeSpan -start (Get-Date -Date "1/1/1970") -end (Get-Date).ToUniversalTime()).TotalMilliseconds)
    $hmac = New-Object System.Security.Cryptography.HMACSHA256
    $hmac.Key = [Text.Encoding]::UTF8.GetBytes($AccessKey)
    $httpVerb = 'PATCH'
    $requestVars = $httpVerb + $epoch + $Data + $ResourcePath
    $signatureBytes = $hmac.ComputeHash([Text.Encoding]::UTF8.GetBytes($requestVars))
    $signatureHex = [System.BitConverter]::ToString($signatureBytes) -replace '-'
    $signature = [System.Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($signatureHex.ToLower()))
    $auth = 'LMv1 ' + $AccessId + ':' + $signature + ':' + $epoch
    $headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
    $headers.Add("Authorization", $auth)
    $headers.Add("Content-Type", 'application/json')

    Write-Log -Message "Updating device properties, device id is $DeviceId"
    Write-Log -Message "API call: $Url"
    $Response = Invoke-RestMethod -Uri $Url -Method Patch -Header $headers -Body $Data
    Write-Log -Message "Response code is $($Response.status)"
}

Function Upload-DataSource ($Tenant, $AccessID, $AccessKey, $DataSourceXML)
{
    $ResourcePath = "/setting/datasources/importxml"
    $Url = 'https://' + $Tenant + '.logicmonitor.com/santaba/rest' + $ResourcePath

    $boundary = [System.Guid]::NewGuid().ToString()

    $data = "--" + $boundary + "`r`n" + "Content-Disposition: form-data; name=""file""; filename=""Datasource.xml""" + "`r`n`r`n" + "$DataSourceXML" + "`r`n" + "--" + $boundary + "--"

    $epoch = [Math]::Round((New-TimeSpan -start (Get-Date -Date "1/1/1970") -end (Get-Date).ToUniversalTime()).TotalMilliseconds)
    $hmac = New-Object System.Security.Cryptography.HMACSHA256
    $hmac.Key = [Text.Encoding]::UTF8.GetBytes($AccessKey)
    $httpVerb = 'POST'
    $requestVars = $httpVerb + $epoch + $data + $ResourcePath
    $signatureBytes = $hmac.ComputeHash([Text.Encoding]::UTF8.GetBytes($requestVars))
    $signatureHex = [System.BitConverter]::ToString($signatureBytes) -replace '-'
    $signature = [System.Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($signatureHex.ToLower()))
    $auth = 'LMv1 ' + $AccessId + ':' + $signature + ':' + $epoch
    $headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
    $headers.Add("Authorization", $auth)
    $headers.Add("Content-Type","multipart/form-data; boundary=$boundary")

    Write-Log -Message "Importing DataSource."
    Write-Log -Message "API call: $Url"
    $Response = Invoke-RestMethod -Uri $Url -Method Post -Header $headers -Body $Data
    Write-Log -Message "Response code is $($Response.status)"
}

Function Show-MainForm
{
    $Form = New-Object System.Windows.Forms.Form
    $Label1 = New-Object System.Windows.Forms.Label
    $Label2 = New-Object System.Windows.Forms.Label
    $Label3 = New-Object System.Windows.Forms.Label
    $Label4 = New-Object System.Windows.Forms.Label
    $Label5 = New-Object System.Windows.Forms.Label
    $Label6 = New-Object System.Windows.Forms.Label
    $Label7 = New-Object System.Windows.Forms.Label
    $Label8 = New-Object System.Windows.Forms.Label
    $Label9 = New-Object System.Windows.Forms.Label
    $Label10 = New-Object System.Windows.Forms.Label
    $Label11 = New-Object System.Windows.Forms.Label
    $Label12 = New-Object System.Windows.Forms.Label
    $Label13 = New-Object System.Windows.Forms.Label
    $Button2 = New-Object System.Windows.Forms.Button
    $Button3 = New-Object System.Windows.Forms.Button
    $Button4 = New-Object System.Windows.Forms.Button
    $CheckBox1 = New-Object System.Windows.Forms.CheckBox
    $CheckBox2 = New-Object System.Windows.Forms.CheckBox
    $TextBox1 = New-Object System.Windows.Forms.TextBox
    $TextBox2 = New-Object System.Windows.Forms.TextBox
    $TextBox3 = New-Object System.Windows.Forms.TextBox
    $TextBox4 = New-Object System.Windows.Forms.TextBox
    $TextBox5 = New-Object System.Windows.Forms.TextBox
    $TextBox6 = New-Object System.Windows.Forms.TextBox
    $TextBox7 = New-Object System.Windows.Forms.TextBox
    $TextBox8 = New-Object System.Windows.Forms.TextBox
    $GroupBox1 = New-Object System.Windows.Forms.GroupBox
    $GroupBox2 = New-Object System.Windows.Forms.GroupBox

    $Label1.Text = "Application Name:"
    $Label1.Location = "20,20"
    $Label1.Size = "100,15"
    $Label1.TextAlign = "MiddleLeft"

    $Label2.Text = "Office365 Domain:"
    $Label2.Location = "20,45"
    $Label2.Size = "100,15"
    $Label2.TextAlign = "MiddleLeft"

    $Label3.Text = "Application ID:"
    $Label3.Location = "20,70"
    $Label3.Size = "100,15"
    $Label3.TextAlign = "MiddleLeft"

    $Label4.Text = "LM company:"
    $Label4.Location = "20,20"
    $Label4.Size = "100,15"
    $Label4.TextAlign = "MiddleLeft"

    $Label5.Text = "API ID:"
    $Label5.Location = "20,45"
    $Label5.Size = "100,15"
    $Label5.TextAlign = "MiddleLeft"

    $Label6.Text = "API Key:"
    $Label6.Location = "20,70"
    $Label6.Size = "100,15"
    $Label6.TextAlign = "MiddleLeft"

    $Label7.Text = ".logicmonitor.com"
    $Label7.Location = "280,20"
    $Label7.Size = "100,15"
    $Label7.TextAlign = "MiddleLeft"

    $Label8.Text = "You must have Powershell for Azure`n installed for one-time setup."
    $Label8.Location = "500,20"
    $Label8.Size = "200,25"
    $Label8.TextAlign = "MiddleCenter"

    $Label9.Text = "Creates an app in your `n Office365 account and gets info from it."
    $Label9.Location = "500,90"
    $Label9.Size = "200,25"
    $Label9.TextAlign = "MiddleCenter"

    $Label10.Text = "Applies the DataSource`nto the device you specify and sets`nrequired properties on the device."
    $Label10.Location = "500,185"
    $Label10.Size = "200,40"
    $Label10.TextAlign = "MiddleCenter"

    $Label11.Text = "Status: Ready"
    $Label11.Location = "500,355"
    $Label11.Size = "200,25"
    $Label11.TextAlign = "MiddleLeft"

    $Label12.Text = "Apply to device:"
    $Label12.Location = "20,95"
    $Label12.Size = "100,15"
    $Label12.TextAlign = "MiddleLeft"

    $Label13.Text = ""
    $Label13.Location = "500,325"
    $Label13.Size = "250,40"
    $Label13.TextAlign = "MiddleLeft"

    $TextBox1.Location = "130,20"
    $TextBox1.Size = "300,15"
    if ($DefaultApplicationName)
    {
        $TextBox1.Text = $DefaultApplicationName
    }
    else
    {
        $TextBox1.ForeColor = "GrayText"
        $TextBox1.Text = "Application Name"
    }
    $TextBox1.add_Leave({
        
        if ($TextBox1.Text.Length -eq 0)
        {
            $TextBox1.ForeColor = "GrayText"
            $TextBox1.Text = "Application Name"
        }
    })
    $TextBox1.add_Enter({
        
        if ($TextBox1.Text -eq "Application Name")
        {
            $TextBox1.Text = ""
            $TextBox1.ForeColor = "WindowText"
        }
    })

    $TextBox2.Location = "130,45"
    $TextBox2.Size = "300,15"
    if ($DefaultOffice365Domain)
    {
        $TextBox2.Text = $DefaultOffice365Domain
    }
    else
    {
        $TextBox2.ForeColor = "GrayText"
        $TextBox2.Text = "yourcompany.com"
    }
    $TextBox2.add_Leave({
        
        if ($TextBox2.Text.Length -eq 0)
        {
            $TextBox2.ForeColor = "GrayText"
            $TextBox2.Text = "yourcompany.com"
        }
    })
    $TextBox2.add_Enter({
        
        if ($TextBox2.Text -eq "yourcompany.com")
        {
            $TextBox2.Text = ""
            $TextBox2.ForeColor = "WindowText"
        }
    })

    $TextBox3.Location = "130,70"
    $TextBox3.Size = "300,15"
    $TextBox3.ReadOnly = $true

    $TextBox4.Location = "130,20"
    $TextBox4.Size = "150,15"
    if ($DefaulLMTenantName)
    {
        $TextBox4.Text = $DefaulLMTenantName
    }
    else
    {
        $TextBox4.ForeColor = "GrayText"
        $TextBox4.Text = "yourcompany"
    }
    $TextBox4.add_Leave({
        
        if ($TextBox4.Text.Length -eq 0)
        {
            $TextBox4.ForeColor = "GrayText"
            $TextBox4.Text = "yourcompany"
        }
    })
    $TextBox4.add_Enter({
        
        if ($TextBox4.Text -eq "yourcompany")
        {
            $TextBox4.Text = ""
            $TextBox4.ForeColor = "WindowText"
        }
    })

    $TextBox5.Location = "130,45"
    $TextBox5.Size = "300,15"
    if ($DefaultLMApiId)
    {
        $TextBox5.Text = $DefaultLMApiId
    }
    else 
    {
        $TextBox5.ForeColor = "GrayText"
        $TextBox5.Text = "Get From Settings > Users > API tokens"
    }
    $TextBox5.add_Leave({
        
        if ($TextBox5.Text.Length -eq 0)
        {
            $TextBox5.ForeColor = "GrayText"
            $TextBox5.Text = "Get From Settings > Users > API tokens"
        }
    })
    $TextBox5.add_Enter({
        
        if ($TextBox5.Text -eq "Get From Settings > Users > API tokens")
        {
            $TextBox5.Text = ""
            $TextBox5.ForeColor = "WindowText"
        }
    })

    $TextBox6.Location = "130,70"
    $TextBox6.Size = "300,15"
    $TextBox6.PasswordChar = '*'
    if ($DefaultLMApiKey)
    {
        $TextBox6.Text = $DefaultLMApiKey
    }
    else
    {
        $TextBox6.ForeColor = "GrayText"
        $TextBox6.Text = "Get From Settings > Users > API tokens (Shows as dots)"
        $TextBox6.PasswordChar = $null
    }
    $TextBox6.add_Leave({
        
        if ($TextBox6.Text.Length -eq 0)
        {
            $TextBox6.ForeColor = "GrayText"
            $TextBox6.Text = "Get From Settings > Users > API tokens (Shows as dots)"
            $TextBox6.PasswordChar = $null
        }
    })
    $TextBox6.add_Enter({
        
        if ($TextBox6.Text -eq "Get From Settings > Users > API tokens (Shows as dots)")
        {
            $TextBox6.Text = ""
            $TextBox6.ForeColor = "WindowText"
            $TextBox6.PasswordChar = '*'
        }
    })

    $TextBox7.Location = "20,400"
    $TextBox7.Size = "940,370"
    $TextBox7.Multiline = $true
    $TextBox7.ScrollBars = "Vertical"
    $LogObjectReference = [Ref]$TextBox7

    $TextBox8.Location = "130,95"
    $TextBox8.Size = "300,15"
    if ($DefaultLMDeviceName)
    {
        $TextBox8.Text = $DefaultLMDeviceName
    }
    else
    {
        $TextBox8.ForeColor = "GrayText"
        $TextBox8.Text = "The display name as it shows in the device tree"
    }
    $TextBox8.add_Leave({
        
        if ($TextBox8.Text.Length -eq 0)
        {
            $TextBox8.ForeColor = "GrayText"
            $TextBox8.Text = "The display name as it shows in the device tree"
        }
    })
    $TextBox8.add_Enter({
        
        if ($TextBox8.Text -eq "The display nameas it shows in the device tree")
        {
            $TextBox8.Text = ""
            $TextBox8.ForeColor = "WindowText"
        }
    })

    $Button2.Text = "1) Register app"
    $Button2.Location = "500,60"
    $Button2.Size = "200,25"
    $Button2.add_Click({

        $Label11.Text = "Status: Busy"
        $ExecutionStartTime = Get-Date

        $TargetApplicationName = $TextBox1.Text

        Write-Log -Message "Connecting to AzureAD"
        $AzureADConnection = Connect-AzureAD

        Write-Log -Message "Getting list of registered applications"
        $TargetApplication = Get-AzureADApplication -SearchString $TargetApplicationName | Where-Object {$_.DisplayName -eq $TargetApplicationName}

        
        if ($TargetApplication)
        {
            Write-Log -Message "Discovered Application $TargetApplicationName, ID is $($TargetApplication.AppId)"
            $TextBox3.Text = $TargetApplication.AppId
        }
        else
        {
            Write-Log -Message "Creating new App"
            $ManagedAPI = Get-AzureADServicePrincipal -SearchString "Office 365 Management APIs"
            $ManagedAPIPermission = $ManagedAPI.AppRoles | Where-Object {$_.Value -eq "ServiceHealth.Read"}
            $ResourceAccessObject = New-Object -TypeName "Microsoft.Open.AzureAD.Model.RequiredResourceAccess"
            $ResourceAccessList = New-Object -TypeName "Microsoft.Open.AzureAD.Model.ResourceAccess" -ArgumentList $ManagedAPIPermission.Id,"Scope"
            $ResourceAccessObject.ResourceAccess = $ResourceAccessList
            $ResourceAccessObject.ResourceAppId = $ManagedAPI.AppId
            $Office365AccessObject = $ResourceAccessObject

            $ManagedAPI = Get-AzureADServicePrincipal -SearchString "Microsoft Graph"
            $ManagedAPIPermission = $ManagedAPI.AppRoles | Where-Object {$_.Value -eq "Reports.Read.All"}
            $ResourceAccessObject = New-Object -TypeName "Microsoft.Open.AzureAD.Model.RequiredResourceAccess"
            #$ResourceAccessList = New-Object -TypeName "Microsoft.Open.AzureAD.Model.ResourceAccess" -ArgumentList $ManagedAPIPermission.Id,"Scope"
            $ResourceAccessList = New-Object -TypeName "Microsoft.Open.AzureAD.Model.ResourceAccess" -ArgumentList "02e97553-ed7b-43d0-ab3c-f8bace0d040c","Scope"
            $ResourceAccessObject.ResourceAccess = $ResourceAccessList
            $ResourceAccessObject.ResourceAppId = $ManagedAPI.AppId
            $GraphAPIAccessObject = $ResourceAccessObject

            $CreatedApp = New-AzureADApplication -DisplayName $TargetApplicationName -ReplyUrls "urn:ietf:wg:oauth:2.0:oob" -RequiredResourceAccess $Office365AccessObject,$GraphAPIAccessObject -PublicClient $true
            Write-Log -Message "App Id is $($CreatedApp.AppId)"
            $TextBox3.Text = $CreatedApp.AppId
        }

        while ($true)
        {
            Start-Sleep -Seconds 15            
            $TargetApplication = Get-AzureADApplication -SearchString $TargetApplicationName
            if ($TargetApplication)
            {
                Write-Log -Message "Application verified"
                break
            }
        }

        Write-Log -Message "Disconnecting from AzureAD"
        Disconnect-AzureAD

        Write-Log -Message "Step 1 complete. Continue with Step2."

        $ExecutionEndTime = Get-Date
        $ExecutionTime = New-TimeSpan $ExecutionStartTime $ExecutionEndTime
        $Label13.Text = "App registered in $($ExecutionTime.TotalSeconds.ToString("N2")) seconds."

        $Label11.Text = "Status: Ready"
    })

    $Button3.Text = "2) Setup LogicMonitor"
    $Button3.Location = "500,155"
    $Button3.Size = "200,25"
    $Button3.add_Click({

        $Label11.Text = "Status: Busy"
        $ExecutionStartTime = Get-Date
        
        $Office365Domain = $TextBox2.Text
        $AzureApplicationID = $TextBox3.Text
        $Tenant = $TextBox4.Text
        $AccessID = $TextBox5.Text
        $AccessKey = $TextBox6.Text
        $DeviceName = $TextBox8.Text

        $ObtainedDevice = Get-LmDevice -Tenant $Tenant -AccessID $AccessID -AccessKey $AccessKey -DeviceName $DeviceName

        if ($ObtainedDevice.data.items.Count -ne 1)
        {
            Write-Log -Message "Discovered $($ObtainedDevice.data.items.Count) devices, aborting operation."
            return
        }

        $DeviceID = $ObtainedDevice.data.items[0].id

        Write-Log -Message "Deploying tokens of target app"
        $RedirectUri = "urn:ietf:wg:oauth:2.0:oob"
        $URLEncodedRedirectUri = [System.Web.HttpUtility]::UrlEncode($RedirectUri)
        $Resource = "https://manage.office.com"
        $URLEncodedResource = [System.Web.HttpUtility]::UrlEncode($Resource)
        $ResourceGraph = "https://graph.microsoft.com"
        $URLEncodedResourceGraph = [System.Web.HttpUtility]::UrlEncode($ResourceGraph)
        $Authority = "https://login.microsoftonline.com/$Office365Domain/oauth2/authorize?"

        $Uri = "$Authority" + "client_id=$AzureApplicationID" + "&response_type=code" + "&redirect_uri=$URLEncodedRedirectUri" + "&response_mode=query" + "&prompt=admin_consent"
        $QueryOutput = Show-OAuthWindow -Url $Uri
        $Code = $QueryOutput.Code

        Write-Log -Message "Requesting Office 365 tokens"
        $Authority = "https://login.windows.net/$Office365Domain/oauth2/token"
        $Body = "resource=$URLEncodedResource" + "&client_id=$AzureApplicationID" + "&redirect_uri=$URLEncodedRedirectUri" + "&grant_type=authorization_code" + "&code=$Code"
        $Result = Invoke-RestMethod -Method POST -uri $Authority -Body $Body
        Write-Log -Message "Response: Token type: $($Result.token_type); Scope: $($Result.scope); Expires in: $($Result.expires_in); Resource: $($Result.resource)"


        Write-Log -Message "Requesting Graph API tokens"
        ##auth code may be used only once
        #$Authority = "https://login.windows.net/$Office365Domain/oauth2/token"
        #$Body = "resource=$URLEncodedResourceGraph" + "&client_id=$AzureApplicationID" + "&redirect_uri=$URLEncodedRedirectUri" + "&grant_type=authorization_code" + "&code=$Code"
        #$ResultGraph = Invoke-RestMethod -Method POST -uri $Authority -Body $Body
        #$ResultGraph
        $Authority = "https://login.windows.net/$Office365Domain/oauth2/token"
        $Body = "resource=$URLEncodedResourceGraph" + "&client_id=$AzureApplicationID" + "&redirect_uri=$URLEncodedRedirectUri" + "&grant_type=refresh_token" + "&refresh_token=$($Result.refresh_token)"
        $ResultGraph = Invoke-RestMethod -Method POST -uri $Authority -Body $Body
        Write-Log -Message "Response: Token type: $($ResultGraph.token_type); Scope: $($ResultGraph.scope); Expires in: $($ResultGraph.expires_in); Resource: $($ResultGraph.resource)"

        $PropertiesArray = @()
        $PropertyObject = New-Object System.Object
        $PropertyObject | Add-Member -MemberType NoteProperty -Name "name" -Value "Office365.TokenExpires"
        $PropertyObject | Add-Member -MemberType NoteProperty -Name "value" -Value $Result.expires_on
        $PropertiesArray += $PropertyObject

        $PropertyObject = New-Object System.Object
        $PropertyObject | Add-Member -MemberType NoteProperty -Name "name" -Value "Office365.AccessToken"
        $PropertyObject | Add-Member -MemberType NoteProperty -Name "value" -Value $Result.access_token
        $PropertiesArray += $PropertyObject

        $PropertyObject = New-Object System.Object
        $PropertyObject | Add-Member -MemberType NoteProperty -Name "name" -Value "Office365.RefreshToken.key"
        $PropertyObject | Add-Member -MemberType NoteProperty -Name "value" -Value $Result.refresh_token
        $PropertiesArray += $PropertyObject

        $PropertyObject = New-Object System.Object
        $PropertyObject | Add-Member -MemberType NoteProperty -Name "name" -Value "Office365.GraphAPI.TokenExpires"
        $PropertyObject | Add-Member -MemberType NoteProperty -Name "value" -Value $ResultGraph.expires_on
        $PropertiesArray += $PropertyObject

        $PropertyObject = New-Object System.Object
        $PropertyObject | Add-Member -MemberType NoteProperty -Name "name" -Value "Office365.GraphAPI.AccessToken"
        $PropertyObject | Add-Member -MemberType NoteProperty -Name "value" -Value $ResultGraph.access_token
        $PropertiesArray += $PropertyObject

        $PropertyObject = New-Object System.Object
        $PropertyObject | Add-Member -MemberType NoteProperty -Name "name" -Value "Office365.GraphAPI.RefreshToken.key"
        $PropertyObject | Add-Member -MemberType NoteProperty -Name "value" -Value $ResultGraph.refresh_token
        $PropertiesArray += $PropertyObject

        $PropertyObject = New-Object System.Object
        $PropertyObject | Add-Member -MemberType NoteProperty -Name "name" -Value "Office365.Tenant"
        $PropertyObject | Add-Member -MemberType NoteProperty -Name "value" -Value $Office365Domain
        $PropertiesArray += $PropertyObject

        $PropertyObject = New-Object System.Object
        $PropertyObject | Add-Member -MemberType NoteProperty -Name "name" -Value "Office365.AppID.key"
        $PropertyObject | Add-Member -MemberType NoteProperty -Name "value" -Value $AzureApplicationID
        $PropertiesArray += $PropertyObject

        $PropertyObject = New-Object System.Object
        $PropertyObject | Add-Member -MemberType NoteProperty -Name "name" -Value "Office365Monitoring"
        $PropertyObject | Add-Member -MemberType NoteProperty -Name "value" -Value "This field Enables O365 Monitoring"
        $PropertiesArray += $PropertyObject

        $PropertyObject = New-Object System.Object
        $PropertyObject | Add-Member -MemberType NoteProperty -Name "name" -Value "Office365.GraphAPIMonitoringO365"
        $PropertyObject | Add-Member -MemberType NoteProperty -Name "value" -Value "This field Enables O365 Graph API Monitoring"
        $PropertiesArray += $PropertyObject

        $PropertyObject = New-Object System.Object
        $PropertyObject | Add-Member -MemberType NoteProperty -Name "name" -Value "Office365.LM.Id"
        $PropertyObject | Add-Member -MemberType NoteProperty -Name "value" -Value $AccessID
        $PropertiesArray += $PropertyObject

        $PropertyObject = New-Object System.Object
        $PropertyObject | Add-Member -MemberType NoteProperty -Name "name" -Value "Office365.LM.Key"
        $PropertyObject | Add-Member -MemberType NoteProperty -Name "value" -Value $AccessKey
        $PropertiesArray += $PropertyObject

        $PropertyObject = New-Object System.Object
        $PropertyObject | Add-Member -MemberType NoteProperty -Name "name" -Value "lm.account"
        $PropertyObject | Add-Member -MemberType NoteProperty -Name "value" -Value $Tenant
        $PropertiesArray += $PropertyObject

        Update-LMDeviceProperties -Tenant $Tenant -AccessID $AccessID -AccessKey $AccessKey -DeviceId $DeviceID -PropertiesObject $PropertiesArray

        if ($CheckBox1.Checked)
        {
            Write-Log -Message "Importing Office 365 datasource template."
            if (Test-Path $DataSourceO365File)
            {
                $DataSource = Get-Content $DataSourceO365File -Raw
                Upload-DataSource -Tenant $Tenant -AccessID $AccessID -AccessKey $AccessKey -DataSourceXML $DataSource
            }
            else 
            {
                Write-Log -Message "DataSource file not found!"
                [System.Windows.Forms.MessageBox]::Show("Datasource file template not found, this step would be ignored.")
            }
        }

        if ($CheckBox2.Checked)
        {
            Write-Log -Message "Importing Graph API datasource template."
            if ((Test-Path $DataSourceGraphFile1) -and (Test-Path $DataSourceGraphFile2) -and (Test-Path $DataSourceGraphFile3))
            {
                $DataSource = Get-Content $DataSourceGraphFile1 -Raw
                Upload-DataSource -Tenant $Tenant -AccessID $AccessID -AccessKey $AccessKey -DataSourceXML $DataSource

                $DataSource = Get-Content $DataSourceGraphFile2 -Raw
                Upload-DataSource -Tenant $Tenant -AccessID $AccessID -AccessKey $AccessKey -DataSourceXML $DataSource

                $DataSource = Get-Content $DataSourceGraphFile3 -Raw
                Upload-DataSource -Tenant $Tenant -AccessID $AccessID -AccessKey $AccessKey -DataSourceXML $DataSource
            }
            else 
            {
                Write-Log -Message "DataSource file not found!"
                [System.Windows.Forms.MessageBox]::Show("Datasource file template not found, this step would be ignored.")
            }
        }

        $ExecutionEndTime = Get-Date
        $ExecutionTime = New-TimeSpan $ExecutionStartTime $ExecutionEndTime
        $Label13.Text = "LogicMonitors properties applied in $($ExecutionTime.TotalSeconds.ToString("N2")) seconds."

        $Label11.Text = "Status: Ready"
    })

    $Button4.Text = "Log file"
    $Button4.Location = "600,290"
    $Button4.Size = "100,25"
    $Button4.add_Click({
        &($LogFile)
    })

    $CheckBox1.Text = "Office 365"
    $CheckBox1.Location = "500,245"
    $CheckBox1.Size = "210,25"
    $CheckBox1.Checked = $true

    $CheckBox2.Text = "Exchange email stats"
    $CheckBox2.Location = "500,265"
    $CheckBox2.Size = "250,25"
    $CheckBox2.Checked = $true

    $GroupBox1.Size = "450,120"
    $GroupBox1.Location = "15,15"
    $GroupBox1.Text = "Office 365 Info"
    $GroupBox1.Controls.Add($Label1)
    $GroupBox1.Controls.Add($Label2)
    $GroupBox1.Controls.Add($Label3)
    $GroupBox1.Controls.Add($TextBox1)
    $GroupBox1.Controls.Add($TextBox2)
    $GroupBox1.Controls.Add($TextBox3)

    $GroupBox2.Size = "450,140"
    $GroupBox2.Location = "15,150"
    $GroupBox2.Text = "LogicMonitor Info"
    $GroupBox2.Controls.Add($Label4)
    $GroupBox2.Controls.Add($Label5)
    $GroupBox2.Controls.Add($Label6)
    $GroupBox2.Controls.Add($Label7)
    $GroupBox2.Controls.Add($TextBox4)
    $GroupBox2.Controls.Add($TextBox5)
    $GroupBox2.Controls.Add($TextBox6)
    $GroupBox2.Controls.Add($TextBox8)
    $GroupBox2.Controls.Add($Button1)
    $GroupBox2.Controls.Add($Label12)

    $Form.Controls.Add($GroupBox1)
    $Form.Controls.Add($GroupBox2)
    $Form.Controls.Add($Label8)
    $Form.Controls.Add($Label9)
    $Form.Controls.Add($Label10)
    $Form.Controls.Add($Label11)
    $Form.Controls.Add($Label13)
    $Form.Controls.Add($Button2)
    $Form.Controls.Add($Button3)
    $Form.Controls.Add($Button4)
    $Form.Controls.Add($TextBox7)
    $Form.Controls.Add($CheckBox1)
    $Form.Controls.Add($CheckBox2)
    $Form.add_Load({
        if (-not (get-module -listavailable AzureAD))
        {
            Write-Log -Message "AzureAD powershell module not installed"
            $Resp = [System.Windows.Forms.MessageBox]::Show("Would you like to install Azure AD Powershell Module? It's required for one-time setup.", "AzureAD required", "YesNo", "Info")
            if ($Resp -eq "Yes")
            {
                Write-Log -Message "Installation AzureAD powershell module"
                Start-Process -FilePath "powershell.exe" -Verb RunAs -ArgumentList '-command "Install-Module AzureAD -Force"'
            }
        }
    })

	$Form.Text = "Setup Office365 monitoring"
	$Form.Size = "1000,820"
	$Form.MaximizeBox = $false
    #$Form.TopMost = $true
	$Form.FormBorderStyle = "FixedDialog"
    $Form.ShowDialog() | Out-Null 

}

$Script:showWindowAsync = Add-Type -MemberDefinition @"
[DllImport("user32.dll")]
public static extern bool ShowWindowAsync(IntPtr hWnd, int nCmdShow);
"@ -Name "Win32ShowWindowAsync" -Namespace Win32Functions -PassThru
$showWindowAsync::ShowWindowAsync((Get-Process -Id $pid).MainWindowHandle, 2)

Show-MainForm