#Try to grab credentials for office 365
Try{
  $msolcred = get-credential -message "Enter your Office 365 credentials as Username@levelonebank.com"
  connect-msolservice -credential $msolcred -ErrorAction Stop
}

#Since in testing a failure of authentication didn't mean a
#full-on failure of the script there's not reason to stop altogether
#when an error is caught.  We won't worry about it unless the user
#can't seem to get a user result.

Catch{
  #Create a variable to say that Auth failed for later
  $ThisError = "Auth"
}
clear

#Do a loop until we come out with some usable results
#This could also indicate Auth issues since I don't know how to target those
Do
  {
  #No valid results yet since we haven't asked
  $validresults = 0
  write-host "What is the username that you would like to convert?" -backgroundcolor DarkGray -foregroundcolor Black -nonewline
  $usrupn = Read-Host
  $usrobj = (get-msoluser|where-object{$_.userprincipalname -like "*$usrupn*"})
  clear
  #Throw the "error" to let the user know something's amiss
  if (-not $usrobj)
    {
      clear
      #Here's where we say maybe it's an auth issue if you can't get a user result.
      if ($ThisError -eq "Auth")
        {
          Write-Host "The Authentication Information that you submitted was incorrect.. Start the script over." -backgroundcolor red
          Exit
        }
      #But if you didn't have an Auth error AND you can't get a result, you probably typed it wrong
      else
        {
          write-host "We didn't find anything for that search.  It looks for UPN, which is flast@levelonebank.com.  If you input first name it won't work as expected." -backgroundcolor red
        }
    }
  #Once we get good results we switch $validresults to 1 which triggers the until clause and continues the script
  else
    {
      $validresults = 1
    }
  }Until($validresults -eq 1)

#Check and see if the $usrobj returned is an array or an object
#If it's an array we present the user with a choice
if ($usrobj -is [system.array])
  {
    $itsacounter = 0
    foreach($user in $usrobj)
      {
        #loop through the results since it'll be dynamic
        $itsacounter++
        $upn1 = $user.userprincipalname
        write-host "$itsacounter  " -foregroundcolor cyan -nonewline
        write-host "-  $upn1" -foregroundcolor white
      }
      Do{
      Write-Host "Select a number from 1 to $itsacounter" -backgroundcolor DarkGray -foregroundcolor Black -nonewline
      [int]$UChoice = read-host
      #Here we just subtract one from the choice since arrays start at 0 rather than 1
      --$UChoice
      clear
      #Now check whether the user input a number that is within the range we listed
      if($UChoice -ge $itsacounter -Or $UChoice -lt 0){write-host "That number is not in the list.  Try something that is" -backgroundcolor red}
      }Until($Uchoice -lt $itsacounter -And $Uchoice -gt 0)
      #The $usrobj variable is reused here so that no code below has to be adapted
      #depending on whether there is one result or multiple results
      $usrobj = $usrobj[$Uchoice]
  }

#Here we build the 3 letter prefix
$upnName = $usrobj.userprincipalname
$dispName = $usrobj.DisplayName
$initials = $dispname.split(" ")
$namecount = $initials.count

#If the user's displayname only has first and last we'll just throw an X in the middle
switch ($namecount)
  {
    2 {$initials = $initials[0][0] + "x" + $initials[1][0]}
    3 {$initials = $initials[0][0]+$initials[1][0]+$initials[2][0]}
    default
      {
        clear
        #If the user's displayname has either 1 word or more than 3 we give up and ask the user
        $initials = read-host "Something is not right.  Displayname is $dispname.  Enter a 3-letter prefix for the user $dispname."
      }
  }

#We check to see if there's a phone number.  If not, ask for the 4 digit suffix
if ($usrobj.PhoneNumber)
  {
    $phone = $usrobj.PhoneNumber
    $ext = $phone.substring($phone.length - 4, 4)  
  }
else
  {
    clear
    write-host "$upnName has no phone number on file." -backgroundcolor Red
    write-host "Enter 4 digits to use as the username suffix"
    $ext = read-host
  }

#Sandwich it all together
$newID = ($initials+$ext).ToUpper()
$validresults2 = "unarmed"

Do{
  #If this is the first loopthrough we clear the screen before moving ahead
  #If it isn't our first time through we leave the clearing to the if command below
  if($validresults2 -like "unarmed"){clear}
  Write-Host "User Display Name:" -nonewline
  Write-Host "   $dispname" -foregroundcolor Cyan
  Write-Host "Current UPN:" -nonewline
  Write-Host "         $upnName" -foregroundcolor Cyan
  Write-Host "New UPN:" -nonewline
  Write-Host "             $newID@levelonebank.com" -foregroundcolor Cyan
  Write-Host "------------------------------------------------------"
  Write-Host "Accept changes? (Y or N)" -backgroundcolor DarkGray -foregroundcolor Black
  $UConfirm = read-host
  if ($Uconfirm -eq "y" -Or $Uconfirm -eq "n" -Or $Uconfirm -eq "Y" -Or $Uconfirm -eq "N")
    {
      $validresults2 = "valid"
    }
  else
    {
      $validresults2 = "errored"
      clear
      Write-Host "Invalid Input.  Answer either 'Y' for Yes or 'N' for No" -backgroundcolor Red
    }
  }Until($validresults2 -like "valid")

#Compile the command because we want to show the end user what's going on
$CmdToRun = "Set-MsolUserPrincipalName -NewUserPrincipalName `"$newID@levelonebank.com`" -UserPrincipalName `"$upnName`""

#Confirm again, since this is kind of a big deal
$validresults3 = "unarmed"
Do{
  if($validresults3 -like "unarmed"){clear}
  write-host "The following command will be run:" -nonewline
  write-host $CmdToRun -foregroundcolor Cyan
  write-host "Hit 'Y' to continue or 'N' to cancel"
  $UConfirm2 = read-host
  if ($Uconfirm2 -eq "y" -Or $Uconfirm2 -eq "n" -Or $Uconfirm2 -eq "Y" -Or $Uconfirm2 -eq "N")
    {
      $validresults3 = "valid"
    }
  else
    {
      $validresults3 = "errored"
      clear
      Write-Host "Invalid Input.  Answer either 'Y' for Yes or 'N' for No" -backgroundcolor Red
    }
  }Until($validresults3 -like "valid")
if ($Uconfirm2 -like "y")
  {
    Set-MsolUserPrincipalName -NewUserPrincipalName "$newID@levelonebank.com" -UserPrincipalName "$upnName"
    if($?)   #If we find an error in the set command we will
      {      #move forward, but make the output red to alert the user
        $FinalResult = "Success"
        $Color1 = "Green"
        $Color2 = "DarkGray"
      }
    else
      {
        $FinalResult = "Failure"
        $Color1 = "Red"
        $Color2 = "Black"
      }
    clear
    #Give the user a decent feedback to confirm that the change worked
    Write-Host "It appears that the user conversion was a $FinalResult" -backgroundcolor $Color1 -foregroundcolor $Color2
    $FinalResultCmd = Get-MSOLUser -UserPrincipalName $newID@levelonebank.com
    $fdispname = $FinalResultCmd.DisplayName
    $fsignin = $FinalResultCmd.SignInName
    $flicensed = $FinalResultCmd.IsLicensed
    $femails = $FinalResultCmd.ProxyAddresses
    Write-Host "Display Name:" -nonewline
    Write-Host "               $fdispname"
    Write-Host "Sign In Name:" -nonewline
    Write-Host "               $fsignin"
    Write-Host "Licensed?" -nonewline
    Write-Host "                   $flicensed"
    Write-Host "List of Email Addresses:"
    $femails
  }
else
  {
    Clear
    #If they hit no at this point I don't care to loop back..
    Write-Host "Quitting since you said N for NO" -backgroundcolor Red
    Exit
  }


#So that we never get to 'Finally'
EXIT

#Obligatory 'Finally'
Finally{"Script End"}