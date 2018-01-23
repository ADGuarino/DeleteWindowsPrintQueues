############################################################################################
#Specifying Windows Print Servers
$PrintServers = (<#"Enter", "Print", "Servers", "Here"#>)

############################################################################################
#Load the .Net Assembly for my PopUp boxes
[void][System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')

############################################################################################
#Prompt for the Printer Name
$PrinterName = [Microsoft.VisualBasic.Interaction]::InputBox('Enter the Host Name of the Printer', 'Delete Print Queues')
Write-Host "Gathering configuration of $PrinterName..." -ForegroundColor Yellow
############################################################################################
#Get IP Address of Printer for Port Deletion
$IPAddress = get-printer -Name $PrinterName -ComputerName <#EnterPrintServerHere#> -ErrorAction SilentlyContinue | select -ExpandProperty PortName -First 1 -ErrorAction SilentlyContinue


############################################################################################
#Prompt to Verify Printer Info was entered correctly
$a = new-object -comobject wscript.shell 
$intAnswer = $a.popup("Click Yes to delete Windows Print Queues`n`n Printer Name: $PrinterName`n`n Port/IP: $IPAddress`n`nClick No to cancel`n", 
0,"Verify Printer Settings",4)
If ($intAnswer -eq 6) {

############################################################################################
#Actual script that deletes the print queues
    
    foreach ($PrintServer in $PrintServers) {
        Write-Host "Performing task: Delete $PrinterName on $PrintServer"
        $Printer = Get-Printer -ComputerName $PrintServer -Name $PrinterName -ErrorAction SilentlyContinue
        If ($Printer.Name -eq $Null) {
        Write-Host "$PrinterName already does not exist on $Printserver" -ForegroundColor Red
        }
        Else {
        Remove-Printer -ComputerName $PrintServer -Name $PrinterName -ErrorAction SilentlyContinue -Verbose
        }
        $Port = Get-Printerport -ComputerName $PrintServer -Name $IPAddress -ErrorAction SilentlyContinue
        If ($Port.Name -eq $Null) {
        Write-Host "The port $IPAddress already does not exist on $Printserver" -ForegroundColor Red
        }
        Else {
        Remove-PrinterPort -ComputerName $PrintServer -Name $IPAddress -ErrorAction SilentlyContinue -Verbose
        }
        Write-Host "The print queue has been successfully deleted" -ForegroundColor Green                
}

############################################################################################
#Notification that the script has completed. 
$a.popup("The print queues for $PrinterName have been successfully deleted on all Windows Print Servers.")
Exit
}

 else { 
    $a.popup("Delete Windows Print Queues has been cancelled")
    Exit 
} 