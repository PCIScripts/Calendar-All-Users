
Function LoginMFA{
    #This imports the ExchangeOnlineManagement Module. This is assumed that the module is already installed but can be easily obtained or an auto install function can be created if needed
    Import-Module -Name ExchangeOnlineManagement
    #initiates a connection with no details so the technician can login as whoever they need
    Connect-ExchangeOnline -ShowBanner:$False  
    Get-Users  
}

Function Get-Users{
    Clear-Host
    Write-Host "Getting a list of all users"
    $Script:Users = Get-EXOMailbox | Select -ExpandProperty Name
    Write-Host ('Found {0} Users' -f $Script:Users.count)
    Read-Host = 'Press Enter to continue'
    AdjustPermissions
}

Function AdjustPermissions{
    $email = Read-Host "Enter the email address for the calendar"
    $calendar = $email + ":\Calendar"
    Foreach($user in $Script:Users){
        Try{
        Write-Host "Attempting to add the permission for $user"
        Add-MailboxFolderPermission -Identity $calendar -User $user -AccessRights Editor -ErrorAction Stop
            }Catch{
                $message = $_
                Write-Warning $message
                Write-Host "Adding permission for $user failed."
                Write-Host "Would you like to try edit permissions instead?"
                $option = Read-Host '(Y/N)'
                Switch($option){
                    'Y'{
                        Write-Host "Attempting to Edit the permission for $user"
                        Try{
                        Set-MailboxFolderPermission -Identity 'LancasterSmallOffice@ageuklancs.org.uk:\Calendar' -User $user -AccessRights Editor -ErrorAction Stop
                        }Catch{
                            $message = $_
                            Write-Warning $message
                            Write-Host "Editing permission for $user failed."
                        }
                    }
                    'N'{}
                }

            }
        }
}

LoginMFA