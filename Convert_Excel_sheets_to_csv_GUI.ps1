###################################################################################
#                                                                                 #
#  GUI interface for converting Excel sheets to CSV                               #
#    - Will work on 1 excel sheet as well as multiple                             #
#    - Names each CSV the name of the sheet                                       #
#                                                                                 #
#                                                                                 #
#  Created By: John Hudson                                                        #
#  Creation Date: 9/19/22                                                         #
#                                                                                 #
#                                                                                 #
###################################################################################

###################################################################################
# Revision History                                                                #
#      9/19/22  -   Creation date                                                 #
#      9/21/22  -   Updated Comments and added seperate folder selection          #
#                                                                                 #
###################################################################################

###################################################################################
#                                                                                 #
#  Window Box Creation                                                            #
#                                                                                 #
###################################################################################

[void][System.Reflection.Assembly]::LoadwithPartialName("System.Drawing")
[void][System.Reflection.Assembly]::LoadwithPartialName("System.Windows.Forms")

$form = New-Object System.Windows.Forms.Form
$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedSingle
$form.text = "Convert Excel Sheet to CSV Tool"
$form.size = New-Object System.Drawing.Size(1400,600)
$form.startposition = "CenterScreen"
$form.backgroundImageLayout = "Zoom"
$form.MinimizeBox = $true
$form.maximizeBox = $true
$form.SizeGripStyle = "Show"
$Icon = [system.drawing.icon]::ExtractAssociatedIcon($PSHOME + "\powershell.exe")
$form.Icon = $Icon
$FormTabControl = New-Object System.Windows.Forms.TabControl
$FormTabControl.Size = "1375,540"
$FormTabControl.Location = "10,10"
$Form.Controls.Add($FormTabControl)

#Convert Excel to CSV Files Tab
$ConvertExcelTab = New-Object System.Windows.Forms.Tabpage
$ConvertExcelTab.DataBindings.DefaultDataSourceUpdateMode = 0
$ConvertExcelTab.UseVisualStyleBackColor = $true
$ConvertExcelTab.Name = "Convert Excel sheets to CSV"
$ConvertExcelTab.Text = "Convert Excel sheets to CSV"
$FormTabControl.Controls.Add($ConvertExcelTab)
#File Text Box
$FileTextBox = New-Object System.Windows.Forms.TextBox
$FileTextBox.Location = New-Object System.Drawing.Size(15,100)
$FileTextBox.Size = New-Object System.Drawing.Size(180,20)
$ConvertExcelTab.Controls.Add($FileTextBox)
#Select File Button
$FileSelectionButton = New-Object System.Windows.Forms.Button
$FileSelectionButton.Location = New-Object System.Drawing.Size(15,30)
$FileSelectionButton.Size = New-Object System.Drawing.Size(180,60)
$FileSelectionButton.Text = "Select Excel File to Convert Sheets to CSVs"
$FileSelectionButton.Add_Click({FileSelectionForExcelToCSV})
$ConvertExcelTab.Controls.Add($FileSelectionButton)
#Select Location to Save Text Box
$LocationTextBox = New-Object System.Windows.Forms.TextBox
$LocationTextBox.Location = New-Object System.Drawing.Size(15,200)
$LocationTextBox.Size = New-Object System.Drawing.Size(180,20)
$ConvertExcelTab.Controls.Add($LocationTextBox)
#Select Location Button
$LocationSelectionButton = New-Object System.Windows.Forms.Button
$LocationSelectionButton.Location = New-Object System.Drawing.Size(15,130)
$LocationSelectionButton.Size = New-Object System.Drawing.Size(180,60)
$LocationSelectionButton.Text = "Select Folder to Put CSVs"
$LocationSelectionButton.Add_Click({LocationSelectionForExcelToCSV})
$ConvertExcelTab.Controls.Add($LocationSelectionButton)
#Convert File Button
$ConvertSelectionButton = New-Object System.Windows.Forms.Button
$ConvertSelectionButton.Location = New-Object System.Drawing.Size(15,400)
$ConvertSelectionButton.Size = New-Object System.Drawing.Size(180,60)
$ConvertSelectionButton.Text = "Convert Sheets to CSV"
$ConvertSelectionButton.Add_Click({ConvertExcelDocument})
$ConvertExcelTab.Controls.Add($ConvertSelectionButton)
#output Box
$ConvertExcelDocumentOutput = New-Object System.Windows.Forms.RichTextBox
$ConvertExcelDocumentOutput.Location = New-Object System.Drawing.Size(220,10)
$ConvertExcelDocumentOutput.Size = New-Object System.Drawing.Size(1130,500)
$ConvertExcelDocumentOutput.Font = New-Object System.Drawing.Font("Arial",12)
$ConvertExcelDocumentOutput.Multiline = $true
$ConvertExcelDocumentOutput.ScrollBars = "ForcedBoth"
$ConvertExcelDocumentOutput.Text = ""
$ConvertExcelTab.Controls.Add($ConvertExcelDocumentOutput)

###################################################################################
#                                                                                 #
#  Functions                                                                      #
#                                                                                 #
###################################################################################

function FileSelectionForExcelToCSV{
    $fileToConvertSelection = New-Object System.Windows.Forms.OpenFileDialog
    $fileToConvertSelection.InitialDirectory = "C:\"
    $fileToConvertSelection.filter = "All Files (*.*)| *.*"
    $null = $fileToConvertSelection.ShowDialog()
    $FileTextBox.text = $fileToConvertSelection.filename
}
function LocationSelectionForExcelToCSV{
    $LocationSelection = New-Object System.Windows.Forms.FolderBrowserDialog
    #$LocationSelection.InitialDirectory = "C:\"
    $null = $LocationSelection.ShowDialog()
    $LocationTextBox.text = $LocationSelection.selectedpath
}

function ConvertExcelDocument{
    try{
        $xls = $FileTextBox.Text
        $ConvertExcelDocumentOutput.Text += "`nStarting CSV creation process...`n"
        $objExcel = New-Object -Comobject Excel.Application
        $objExcel.Visible = $false
        $objExcel.DisplayAlerts = $false
        $workbook = $objExcel.Workbooks.Open($xls)
        $sheets = $workbook.sheets

        foreach ($sheet in $sheets){
            $csvFilename = $LocationTextBox.text + "\" + $sheet.name + ".csv"
            if (-not (Test-Path $csvFilename)){
                $sheet.SaveAs($csvFilename, 6)
                $ConvertExcelDocumentOutput.Text += "Created: " + $csvFilename + " in " + $LocationTextBox.text + "`n"
            }
            else{
                $ConvertExcelDocumentOutput.Text += "Already Exists: " + $csvFilename + " in " + $LocationTextBox.text + "`n"
            }
        }
        $objExcel.quit()
    }
    catch{
        write-host "Error"
    }
} 

#KEEP THIS AT THE END
#This starts the window
$Form.Add_Shown({$form.Activate()})
[void] $form.ShowDialog()