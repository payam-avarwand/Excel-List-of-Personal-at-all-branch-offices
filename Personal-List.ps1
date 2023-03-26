# The script creates an Excel list of foreign personal accounts for each individual Branch Office of the
# Company with 5 parameters per record: First Name + Last Name + Job Position + Email address + Expiration_Date of the account

# The list will be emailed from our Info-Email-address to the Boss.

# ©Payam Avarwand

############################################################
#Create: an Excel Object, a WorkBook and a WorkSheet 

$excel_Obj = New-Object -ComObject Excel.Application
$excel_Obj.visible = $false
$workbook = $excel_Obj.Workbooks.Add()
$Sheet= $workbook.Worksheets.Item(1)

$D = Get-Date -Format 'dd-MM-yyyy' 
$Sheet.Name = "Branch-List-$D"

############################################################
#Create the Title Cells

$Sheet.Cells.Item(1,2) = "First name"
$Sheet.Cells.Item(1,3) = "Last name"
$Sheet.Cells.Item(1,4) = "Position"
$Sheet.Cells.Item(1,5) = "Email address"
$Sheet.Cells.Item(1,6) = "Expiration Date"


############################################################
#Get the Office Names

$OFFICES = Get-ADOrganizationalUnit -LDAPFilter '(name=*)' -SearchBase 'OU=Users,DC=Payam,DC=de' -SearchScope OneLevel | Where-Object {$_.Name -notlike "*no*"} | select Name -ExpandProperty name

############################################################
#Check the Offices

$i1=2 ; $j1=1 ; $i2=2 ; $i3=2 ; $i4=2

foreach ($Office in $OFFICES){

$Sheet.Cells.Item($i1,2) = $Office
$ADD = "OU=Users,OU=$Office,OU=Users,DC=Payam,DC=de"


############################################################
#First_Name + #Last_Name

$i1=$i1+1
$All1 = Get-ADUser -Filter 'enabled -eq "true"' -SearchBase $ADD -Properties * | Select Name -ExpandProperty name | Where-Object {$_.Name -notlike "*test*"} | Sort-Object -Property Name
	foreach($Name1 in $All1){
	$SecondName, $Firstname = $Name1.Split()		
	$SecondName = $SecondName.TrimEnd(",")
	$Sheet.Cells.Item($i1,1) = $j1 #Numbering on the first Column
$Sheet.Cells.Item($i1,2) = $Firstname
$Sheet.Cells.Item($i1,3) = $SecondName
$i1++
$j1++
}

############################################################
#Position
$i2=$i2+1
$All2 = Get-ADUser -Filter 'enabled -eq "true"' -SearchBase $ADD -Properties * | Select Name, title -ExpandProperty title | Where-Object {$_.Name -notlike "*test*"} | Sort-Object -Property Name
	foreach($title in $All2){	
	$Sheet.Cells.Item($i2,4) = $title		
	$i2++
	}


############################################################
#Email_address

$i3=$i3+1
$All3 = Get-ADUser -Filter 'enabled -eq "true"' -SearchBase $ADD -Properties * | Select Name, mail -ExpandProperty mail | Where-Object {$_.Name -notlike "*test*"} | Sort-Object -Property Name
	foreach($mail in $All3){	
	$Sheet.Cells.Item($i3,5) = $mail		
	$i3++
	}

	
############################################################
#Expiration_Date

$i4=$i4+1
$sam = Get-ADUser -Filter 'enabled -eq "true"' -SearchBase $ADD -Properties * | Where-Object {$_.Name -notlike "*test*"} | Select Name, SamAccountName -ExpandProperty SamAccountName | Sort-Object -Property Name
    foreach ($s1 in $sam) {
    $exp = Get-ADUser -Identity $s1 -Properties * | select AccountExpirationDate -ExpandProperty AccountExpirationDate
    $Sheet.Cells.Item($i4,6) = $exp
    $i4++
    }

}

############################################################
# Format and Save the List + Close the File 

$usedRange = $Sheet.UsedRange                                                                                              
$usedRange.EntireColumn.AutoFit()
$workbook.SaveAs("C:\Temp\Branch-Personal-List_$D.xlsx")
$excel_Obj.Quit()

############################################################
# A short delay to make sure the Excel Object is closed.

Start-Sleep -Seconds 5
$Temp_File = "C:\Temp\Branch-Personal-List_$D.xlsx"

############################################################
# Forward the list file as an attached file

$From = "info@payam.de"
$To = "Boss@payam.de"
$Subject = "[AB-ALL] Personal List of branches - $D"
$cc = "payam.avar@payam.de"
$SMTP_Server = "mailint.payam.de"
$body = "Attached you will find the current list of colleagues in all
branches who have active accounts in the IT department.
With best regards"

Send-MailMessage -To $To -Subject $Subject -Body $body -SmtpServer $SMTP_Server -Attachments $Temp_File -Cc $cc -DNO OnSuccess, OnFailure -From $From

Start-Sleep -Seconds 5
Remove-Item -Path $Temp_File


# Created_by Payam.Avarwand
# Payam_avar@yahoo.com
