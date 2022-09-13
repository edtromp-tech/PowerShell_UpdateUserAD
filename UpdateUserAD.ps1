Import-module ActiveDirectory
$DIR = Get-Location 
$InFileName = "\Phase_II.csv"
$OutFileName = "\Outfile.csv"

Import-CSV "$DIR$InFileName" | 
ForEach-Object {
	#Get the csv file data
	#------------------------------------
	$User = $_.User
	$Manager = $_.Manager
	$Department = $_.Department
	$Title = $_.Title
	$Office = $_.Office
	
	#Displays the csv file data
	#------------------------------------
	#write-host($User)
	#write-host($Manager)
	#write-host($Department)
	#write-host($Title)
	#write-host($Office)
	
	
	#Check that the user and manager are 
	#in the Active Directory
	#------------------------------------
	$ADUser = Get-ADUser -Identity $User
	$ADManager = Get-ADUser -Identity $Manager
	#$ADUser = $User
	#$ADManager = $Manager
	write-host("----------------------"
	if ($ADUser){
		write-host("User found in AD")
		write-host("Lets update Department, Title, and Office")
		$UserAnswer = Read-Host -Prompt 'Shall we update? (y or n)'
		if ($ADManager){
			write-host("Manager found in AD")
			write-host("Lets update Manager for User")
			$ManagerAnswer = Read-Host -Prompt 'Shall we update? (y or n)'
		}	
		else { 
			write-host("Manager not found in AD")
			write-host("Logging to Output.csv")
			$csvWrite = [PSCustomObject]@{
				User = $User
				Manager = $Manager
				Department = $Department 
				Title = $Title 
				Office = $Office 
				Error = "Manager not found in AD"
			}
			$csvWrite | Export-Csv ("$DIR$OutFileName") -Append
		}
	}
	else { 
		write-host("User not found in AD")
		write-host("Logging to Output.csv")
		$csvWrite = [PSCustomObject]@{
			User = $User
			Manager = $Manager
			Department = $Department 
			Title = $Title 
			Office = $Office 
			Error = "User not found in AD"
		}
		$csvWrite | Export-Csv ("$DIR$OutFileName") -Append
	}
	
	
	#Update the active directory or log 
	#the data based on user input
	#------------------------------------
	if ($UserAnswer -eq "y") {
		Try {
			write-host("Updating...")
			#At least the first line works
			#-----------------------------------
			Get-ADUser -Identity $User | Set-ADUser -Department $Department -Title $Title -Office $Office
			#$ADUser | Set-ADUser -Department $Department -Title $Title -Office $Office
		}
		Catch {
			write-host("Failed to update Department, Title, and Office. Exception handled (Set-ADUser)")
			write-host("Logging to Output.csv")
			$csvWrite = [PSCustomObject]@{
				User = $User
				Manager = $Manager
				Department = $Department 
				Title = $Title 
				Office = $Office 
				Error = "Failed to update Department, Title, and Office. Exception handled (Set-ADUser)"
			}
			$csvWrite | Export-Csv ("$DIR$OutFileName") -Append
		}
	}
	else {
		write-host("Did not update Department, Title, and Office")
		write-host("Logging to Output.csv")
		$csvWrite = [PSCustomObject]@{
			User = $User
			Manager = $Manager
			Department = $Department 
			Title = $Title 
			Office = $Office 
			Error = "Did not update Department, Title, and Office"
		}
		$csvWrite | Export-Csv ("$DIR$OutFileName") -Append
		
	}
	
	if ($ManagerAnswer -eq "y") {
		Try {
			write-host("Updating...")
			#At least the first line works
			#-----------------------------------
			Get-ADUser -Identity $User | Set-ADUser -Manager $ADManager
			#$ADUser | Set-ADUser -Manager $ADManager
		}
		Catch {
			write-host("Failed to update Manager. Exception handled (Set-ADUser)")
			write-host("Logging to Output.csv")
			$csvWrite = [PSCustomObject]@{
				User = $User
				Manager = $Manager
				Department = $Department 
				Title = $Title 
				Office = $Office 
				Error = "Failed to update Manager. Exception handled (Set-ADUser)"
			}
			$csvWrite | Export-Csv ("$DIR$OutFileName") -Append
		}
		
	}
	else {
		write-host("Did not update Manager")
		write-host("Logging to Output.csv")
		$csvWrite = [PSCustomObject]@{
			User = $User
			Manager = $Manager
			Department = $Department 
			Title = $Title 
			Office = $Office 
			Error = "Did not update Manager"
		}
		$csvWrite | Export-Csv ("$DIR$OutFileName") -Append
	}
}
