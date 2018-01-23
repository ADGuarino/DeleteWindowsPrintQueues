# Get the ID and security principal of the current user account
 $myWindowsID=[System.Security.Principal.WindowsIdentity]::GetCurrent()
 $myWindowsPrincipal=new-object System.Security.Principal.WindowsPrincipal($myWindowsID)
  
 # Get the security principal for the Administrator role
 $adminRole=[System.Security.Principal.WindowsBuiltInRole]::Administrator
  
 # Check to see I am currently running "as Administrator"
 if ($myWindowsPrincipal.IsInRole($adminRole))
    {
    # I am running "as Administrator" - so change the title and background color to indicate this
    $Host.UI.RawUI.WindowTitle = $myInvocation.MyCommand.Definition + "(Elevated)"
    $Host.UI.RawUI.BackgroundColor = "Black"
    clear-host
    }
 else
    {
    # I am not running "as Administrator" - so relaunch as administrator
    
    # Create a new process object that starts PowerShell
    $newProcess = new-object System.Diagnostics.ProcessStartInfo "PowerShell";
    
    # Specify the current script path and name as a parameter
    $newProcess.Arguments = $myInvocation.MyCommand.Definition;
    
    # Indicate that the process should be elevated
    $newProcess.Verb = "runas";
    
    # Start the new process
    [System.Diagnostics.Process]::Start($newProcess);
        
    
    # Exit from the current, unelevated, process
    exit
    }

############################################################################################  
#Adjust window size
$h = get-host
$win = $h.ui.rawui.windowsize
$win.width  = 100 
$win.Height = 30
$h.ui.rawui.set_windowsize($win)

############################################################################################
#Specifying Windows Print Servers
$PrintServers = ("tec-v-prntsrv01", "tec-v-prntsrv02", "tec-v-prntsrv03", "tec-v-prntsrv04")

############################################################################################
#Load the .Net Assembly for my PopUp boxes
[void][System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')

############################################################################################
#Prompt for the Printer Name
$PrinterName = [Microsoft.VisualBasic.Interaction]::InputBox('Enter the Host Name of the Printer', 'Delete Print Queues')
Write-Host "Gathering configuration of $PrinterName..." -ForegroundColor Yellow
############################################################################################
#Get IP Address of Printer for Port Deletion
$IPAddress = get-printer -Name $PrinterName -ComputerName tec-v-prntsrv01 -ErrorAction SilentlyContinue | select -ExpandProperty PortName -First 1 -ErrorAction SilentlyContinue


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
        Remove-Printer -AsJob -ComputerName $PrintServer -Name $PrinterName -ErrorAction SilentlyContinue -Verbose
        }
        $Port = Get-Printerport -ComputerName $PrintServer -Name $IPAddress -ErrorAction SilentlyContinue
        If ($Port.Name -eq $Null) {
        Write-Host "The port $IPAddress already does not exist on $Printserver" -ForegroundColor Red
        }
        Else {
        Remove-PrinterPort -AsJob -ComputerName $PrintServer -Name $IPAddress -ErrorAction SilentlyContinue -Verbose
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