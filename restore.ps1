
#Setting up AD unit
Try {

New-ADOrganizationalUnit -Name "finance" -Path "DC=ucertify,DC=com" -ProtectedFromAccidentalDeletion $false

$newAD = import-csv $PSScriptRoot\financePersonnel.csv
$path = "OU=finance,DC=ucertify,DC=com"


foreach ($ADUSer in $NewAD)
{
$First =$ADUSer.First_Name
$last =$ADUSer.Last_Name
$DisplayName =$ADUSer.First_Name + ' ' + $ADUSer.Last_Name
$PostalCode = $ADUSer.PostalCode
$Office = $ADUSer.OfficePhone
$Mobile =$ADUSer.MobilePhone

New-ADUser -DisplayName $DisplayName -GivenName $First -Surname $last -PostalCode $PostalCode -Officephone $office -MobilePhone $Mobile -Path $path

}
}
Catch [system.outofmemoryexception] {
write-host "A system out of memory exception has occured."
}

Clear-host

Try{

import-module sqlps -DisableNameChecking -force

##Create object for local sql connection

$srv = New-Object -TypeName Microsoft.sqlServer.Management.Smo.Server -ArgumentList .\ucertify3
$db = New-Object -TypeName Microsoft.SqlServer.Management.Smo.Database -ArgumentList $srv, ClientDB
$db.Create()

Invoke-Sqlcmd -ServerInstance .\ucertify3 -Database ClientDB -InputFile $PSScriptRoot\Client_A_Contacts.sql

$table ='dbo.Client_A_Contacts'
$db = 'ClientDB'
#Importing the NewClientData CSV file and placing it into the table that was just created
Import-Csv $PSScriptRoot\NewClientData.csv | ForEach-Object {Invoke-Sqlcmd `
-Database ClientDB -ServerInstance .\ucertify3 -Query "insert into $table (first_name, last_name, city, county, zip, officePhone, mobilePhone) VALUES `
('$($_.first_name)','$($_.last_name)','$($_.city)','$($_.county)','$($_.zip)','$($_.officePhone)','$($_.mobilePhone)')"
}
}
Catch [system.outofmemoryexception] {
write-host "A system out of memory exception has occured."
} 

Get-ADUser -filter * -searchbase "ou=finance,dc=ucertify,dc=com" -properties DisplayName, PostalCode, MobilePhone, OfficePhone | Export-Csv $PSScriptRoot\AdResults -NoTypeInformation 

Invoke-Sqlcmd -database ClientDB -ServerInstance .\Ucertify3 -Query 'select * from dbo.Client_A_Contacts' > $PSScriptRoot\SqlResults.txt
