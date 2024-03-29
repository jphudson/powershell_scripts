<#
  .Synopsis
  A GUI for setting an attribute.  The only modifications that need to be made is set the attribute at line 26 and set the AD servers at line 43.
  .Description
  @author=john hudson

#>
Import-Module ActiveDirectory
### Window Box Creation ##################################
[void][System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")
[void][System.Reflection.Assembly]::LoadwithPartialName("System.Windows.Forms")

$Form = New-Object System.Windows.Forms.Form
$Form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedSingle
$Form.Text = "Attribute Setter"
$Form.Size = New-Object System.Drawing.Size(1010, 250)
$Form.StartPosition = "CenterScreen"
$Form.BackgroundImageLayout = "Zoom"
$Form.MinimizeBox = $true
$Form.MaximizeBox = $true
$Form.SizeGripStyle = "Show"
$Icon = [system.drawing.icon]::ExtractAssociatedIcon($PSHOME + "\powershell.exe")
$Form.Icon = $Icon

#################### Variables ######################
$attribute = "extensionAttribute1"

#################### Functions ######################
function Attribute_Change_Click {
  param (
  )
}

function Hide-Console
{
    $consolePtr = [Console.Window]::GetConsoleWindow()
    #0 hide
    [Console.Window]::ShowWindow($consolePtr, 0)
}

#The main function for looking up the user and setting Attribute
function ButtonClick() {
  $servers = @("server1.com", "sub1.server1.com", "sub2.server1.com")
  $onServer = "nothing"
  $output = "nothing" 
  $username = $InputUsernameBox.Text
  $attributeValue = $listBox.Text
  $count = 0
  try{
    foreach ($server in $servers) {
      try{
        Get-ADUser -Identity $username -Server $server
        if ($?) {
          $currentAttribute = $null
          try{
            $currentAttribute = Get-ADUser $username -Properties * -Server $server | Select-Object -Property $attribute -ExpandProperty $attribute
          }
          catch{
            
          }
          if($attributeValue -eq "Delete Attribute" -and $currentAttribute -ne $null){
            Set-ADuser -Identity $username -Remove @{$attribute=$currentAttribute} -Server $server
            $output = Get-ADUser $username -Properties * -Server $server | Select-Object -Property Name,employeeID,$attribute
            $onServer = "On Server: " + [string]$server
          }
          elseif($currentAttribute -eq $null -and $attributeValue -ne "Delete Attribute"){
            Set-ADuser $username -Add @{$attribute=$attributeValue} -Server $server
            $output = Get-ADUser $username -Properties * -Server $server | Select-Object -Property Name,employeeID,$attribute
            $onServer = "On Server: " + [string]$server
          }
          elseif($currentAttribute -eq $null -and $attributeValue -eq "Delete Attribute"){
            $output = Get-ADUser $username -Properties * -Server $server | Select-Object -Property Name,employeeID,$attribute
            $onServer = "On Server: " + [string]$server
          }
          else{
            Set-ADuser $username -Replace @{$attribute=$attributeValue} -Server $server
            $output = Get-ADUser $username -Properties * -Server $server | Select-Object -Property Name,employeeID,$attribute
            $onServer = "On Server: " + [string]$server
          }
        }
      }
      catch{
        $count++
        #$errorOutput += "User: " + $username + " Not on: " + $server + "`n" + $_ + "`n"
        #$outputBox.Text = $errorOutput 
      }
    }
    #if ($output -eq "nothing"){
    if ($count -eq $servers.Count){
      $errorOutput += "Could not find user in:`n"
      foreach ($server in $servers) {
        $errorOutput = $errorOutput + "$server" + "`n"
      }
      $outputBox.Text = $errorOutput
    }
    else{
      $outputBox.Text = [string]$output + "`n" + $onServer
    } 
  }
  #if the user isn't on any of the servers output error
  catch{
    $outputBox.Text = "There was an error in ButtonClick() Function" + $_
  }
  
}

#################### Input ##########################
### Input Username ##################################
$InputUsernameBox = New-Object System.Windows.Forms.TextBox
$InputUsernameBox.Location = New-Object System.Drawing.Size(10,30)
$InputUsernameBox.Size = New-Object System.Drawing.Size(180,20)
$Form.Controls.Add($InputUsernameBox)
$InputUsernameLabel = New-Object System.Windows.Forms.Label
$InputUsernameLabel.Text = "Username: "
$InputUsernameLabel.AutoSize = $true
$InputUsernameLabel.Location = New-Object System.Drawing.Size(15,10)
$Form.Controls.Add($InputUsernameLabel)

### Input Attribute 5 ################################
$listBoxLabel = New-Object System.Windows.Forms.Label
$listBoxLabel.Text = $attribute + ":"
$listBoxLabel.AutoSize = $true
$listBoxLabel.Location = New-Object System.Drawing.Size(15,60)
$Form.Controls.Add($listBoxLabel)

$listBox = New-Object System.Windows.Forms.ComboBox
$listBox.Location = New-Object System.Drawing.Point(10,80)
$listBox.Size = New-Object System.Drawing.Size(180,20)
[void] $listBox.Items.Add('Value1')
[void] $listBox.Items.Add('Value2')
[void] $listBox.Items.Add('Value3')
$listBox.SelectedIndex = $listBox.FindString('Value1')
$Form.Controls.Add($listBox)


#################### Buttons ##########################
### Set Atttribute 5 Button ##############################
$SetAttribute5 = New-Object System.Windows.Forms.Button
$SetAttribute5.Location = New-Object System.Drawing.Size(15,110)
$SetAttribute5.Size = New-Object System.Drawing.Size(100,30)
$SetAttribute5.Text = "Set Attribute"
$SetAttribute5.Add_Click({ButtonClick})
$Form.Controls.Add($SetAttribute5)

#################### Output ##########################
### Output Box Field #############################
$outputBox = New-Object System.Windows.Forms.RichTextBox
$outputBox.Location = New-Object System.Drawing.Size(200,10)
$outputBox.Size = New-Object System.Drawing.Size(780,180)
$outputBox.Font = New-Object System.Drawing.Font("Arial",12)
$outputBox.MultiLine = $true
$outputBox.ScrollBars = "Vertical"
$outputBox.Text = ""
$Form.Controls.Add($outputBox)

### Start the Window #############################
$Form.Add_Shown({$Form.Activate()})
[void] $Form.ShowDialog()
