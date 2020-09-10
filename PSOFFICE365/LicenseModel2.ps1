

# Connect to Microsoft Online
Import-Module MSOnline
Connect-MsolService -Credential $Office365credentials

write-host "Connecting to Office 365..."

# Get a list of all licences that exist within the tenant


$licensetype = Get-MsolAccountSku | Where {$_.ConsumedUnits -ge 1}

        

# Loop through all licence types found in the tenant
Get-MsolUser -all | foreach-object { 

$headerstring = New-Object PSObject -Property @{"DisplayName" = "";"UserPrincipalName" = "";"Licenses" = "";"ServiceName" = ""}
$headerstring.DisplayName = [String] $_.DisplayName
$headerstring.UserPrincipalName = [String] $_.UserPrincipalName 
$headerstring.Licenses = [String] $_.Licenses.AccountSku.SkuPartNumber
	
$data = @()
foreach ($license in $licensetype){
         foreach($row in $($license.ServiceStatus)){
          #$headerstring.ServiceName = [String] $row.ServicePlan.ServiceName -replace " ",","
              $data = $data + $row.ServicePlan.ServiceName  
          }
         }	
	
	

	
	
    
	
   
        foreach ($service in $data){
          if($_.Licenses -ne (Out-Null)) {
            $headerstring.ServiceName = $headerstring.ServiceName + "," + $service
            }
          }

        
         
	

        #$thislicense = $user.licenses | Where-Object {$_.accountskuid -eq $license.accountskuid}

		
        #$headerstring.AccountSku = $license.SkuPartNumber

        #foreach ($row in $($license.ServiceStatus)){
		#$headerstring.ServiceName = ($row.ServicePlan.servicename)
		
		#foreach ($row in $($thislicense.servicestatus)) {
			
			# Build data string
			#$datastring = ($datastring + "," + $($row.provisioningstatus))
		#}
Return $headerstring 
} | Export-Excel $env:USERPROFILE\Documents\O365_licencas.xlsx -AutoSize -StartRow 1 -TableName Adelina_Licencas 

