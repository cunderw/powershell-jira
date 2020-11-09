Function Write-Host ($message,$nonewline,$backgroundcolor,$foregroundcolor) {$Message | Out-Host}
[void][System.Reflection.Assembly]::LoadWithPartialName("System.Web.Extensions")
[void][System.Reflection.Assembly]::LoadWithPartialName('presentationframework')
Get-Date -DisplayHint Date
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$inputXML = @"
<Window x:Name="mainWindow" x:Class="WpfApp2.MainWindow"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:local="clr-namespace:WpfApp2"
        mc:Ignorable="d"
        Title="JIRA TCT Find / Replace Utility" Height="800" Width="1500" WindowStartupLocation="CenterScreen" Background="{DynamicResource {x:Static SystemColors.WindowBrushKey}}" ResizeMode="NoResize">
    <Grid>
				<Label Content="Enter Jira Query:" HorizontalAlignment="Left" Margin="10,10,0,0" VerticalAlignment="Top" FontSize="16" FontWeight="Bold" Width="250"/>
				<Label Content="Enter text to find:" HorizontalAlignment="Left" Margin="485,570,0,0" VerticalAlignment="Top" FontSize="16" FontWeight="Bold" Width="250"/>
				<Label Content="Enter text to use for replacement:" HorizontalAlignment="Left" Margin="915,570,0,0" VerticalAlignment="Top" FontSize="16" FontWeight="Bold" Width="300"/>
				
                <TextBox x:Name="query" HorizontalAlignment="Left" Height="200" Margin="10,46,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="400" Background="{DynamicResource {x:Static SystemColors.ControlBrushKey}}" BorderBrush="{DynamicResource {x:Static SystemColors.GrayTextBrushKey}}"/>
                <TextBox x:Name="findText" HorizontalAlignment="Left" Height="100" Margin="485,600,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="318" Background="{DynamicResource {x:Static SystemColors.ControlBrushKey}}" BorderBrush="{DynamicResource {x:Static SystemColors.GrayTextBrushKey}}"/>
                <TextBox x:Name="replaceText" HorizontalAlignment="Left" Height="100" Margin="915,600,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="318" Background="{DynamicResource {x:Static SystemColors.ControlBrushKey}}" BorderBrush="{DynamicResource {x:Static SystemColors.GrayTextBrushKey}}"/>

				<Button x:Name="searchForTests" Content="Search for Tests" Margin="10,260,0,0" VerticalAlignment="Top" Width="300" HorizontalAlignment="Left" FontSize="16" FontWeight="Bold"/>
				
				<Label x:Name="testListLabel" Content="Select Tests To Update" HorizontalAlignment="Left" Margin="725,10,0,0" VerticalAlignment="Top" FontSize="16" FontWeight="Bold" Width="250"/>
				
				<ListBox x:Name="testList" HorizontalAlignment="Left" Height="500" Margin="400,46,0,0" VerticalAlignment="Top" Width="900" Background="{DynamicResource {x:Static SystemColors.WindowBrushKey}}" BorderBrush="{DynamicResource {x:Static SystemColors.GrayTextBrushKey}}"/>

				<Button x:Name="updateTests" Content="Search Selected Tests for Text" Margin="675,720,0,0" VerticalAlignment="Top" Width="300" HorizontalAlignment="Left" FontSize="16" FontWeight="Bold"/>

    </Grid>
</Window>
"@       
 
$inputXML = $inputXML -replace 'mc:Ignorable="d"','' -replace "x:N",'N'  -replace '^<Win.*', '<Window'
[xml]$XAML = $inputXML
#Read XAML
 
$reader=(New-Object System.Xml.XmlNodeReader $xaml)

try{
    $Form=[Windows.Markup.XamlReader]::Load( $reader )
}
catch{
    Write-Host "Unable to load Windows.Markup.XamlReader. Double-check syntax and ensure .net is installed."
}

#===========================================================================
# Popup Form Addition
#=========================================================================== 

# Create objects for popup form
$updateStepsForm = New-Object System.Windows.Forms.Form
$updateStepsList = New-Object System.Windows.Forms.ListView
$updateStepsFormSubmitBtn = New-Object System.Windows.Forms.Button
$updateStepsFormCancelBtn = New-Object System.Windows.Forms.Button

$issuesStepFields = @{}
$newIssuesStepFields= @{}
$global:startAt = 0

#
# jiraFindReplace - Create popup form table
# @function
# 
function createPopupTable(){
    # Add Properties to Form
    $updateStepsForm.Width = 1500
    $updateStepsForm.Height = 800
    $updateStepsForm.Text = "Updates Available"
    $updateStepsForm.Size = New-Object System.Drawing.Size(1500,800)

    # Add Properties to Listview
    $updateStepsList.View = 'Details'
    $updateStepsList.Width = 1490
    $updateStepsList.Height = 725
    $updateStepsList.FullRowSelect = $true
    $updateStepsList.Size = New-Object System.Drawing.Size(1500,725)

    # Make Columns for Listview
    $nameColumn = New-Object System.Windows.Forms.ColumnHeader
    $nameColumn.name = "TestName"
    $nameColumn.Text = "Test"
    $nameColumn.Width = "100"

    $rowIndexColumn = New-Object System.Windows.Forms.ColumnHeader
    $rowIndexColumn.name = "RowIndex"
    $rowIndexColumn.Text = "Row"
    $rowIndexColumn.Width = "50"

    $columnIndexColumn = New-Object System.Windows.Forms.ColumnHeader
    $columnIndexColumn.name = "ColumnIndex"
    $columnIndexColumn.Text = "Column"
    $columnIndexColumn.Width = "50"

    $oldStepColumn = New-Object System.Windows.Forms.ColumnHeader
    $oldStepColumn.name = "OldStep"
    $oldStepColumn.Text = "Old Step"
    $oldStepColumn.Width = "640"

    $newStepColumn = New-Object System.Windows.Forms.ColumnHeader
    $newStepColumn.name = "NewStep"
    $newStepColumn.Text = "New Step"
    $newStepColumn.Width = "640"

    # Add Columns to Listview
    $ret = $updateStepsList.Columns.Add($nameColumn)
    $ret = $updateStepsList.Columns.Add($rowIndexColumn)
    $ret = $updateStepsList.Columns.Add($columnIndexColumn)
    $ret = $updateStepsList.Columns.Add($oldStepColumn)
    $ret = $updateStepsList.Columns.Add($newStepColumn)

    # Add Listview and buttons to form
    addControls
}

#
# jiraFindReplace - adds updates to popup form table
# @function
# @param {string} key - name of issue
# @param {string} rowIndex - row of issue step
# @param {string} columnIndex - column of issue step
# @param {string} oldText - old step text
# @param {string} newText - new step text
#
function addUpdatesToTable($key, $rowIndex, $columnIndex, $oldText, $newText){
    Write-Host "Adding update Item for test $key at $rowIndex , $columnIndex from $oldText to $newText"
    # Create ListView Item
    $stepToUpdate = New-Object System.Windows.Forms.ListViewItem($key)
    $ret = $stepToUpdate.SubItems.Add($rowIndex)
    $ret = $stepToUpdate.SubItems.Add($columnIndex)
    $ret = $stepToUpdate.SubItems.Add($oldText)
    $ret = $stepToUpdate.SubItems.Add($newText)

    # Add Item to the ListView
    $ret = $updateStepsList.Items.Add($stepToUpdate)
}

#
# jiraFindReplace - adds controls to popup form
# @function
# 
function addControls(){
    # Add the ListView to the form
    $ret = $updateStepsForm.Controls.Add($updateStepsList)

    # Add Button controls to the form
    $updateStepsFormSubmitBtn.Text = "Update Selected Steps"
    $updateStepsFormSubmitBtn.Width = 200
    $updateStepsFormSubmitBtn.Location = New-Object System.Drawing.Size(600,730) 

    $updateStepsFormCancelBtn.Text = "Cancel"
    $updateStepsFormCancelBtn.Width = 200
    $updateStepsFormCancelBtn.Location = New-Object System.Drawing.Size(800,730) 

    $ret = $updateStepsForm.Controls.Add($updateStepsFormSubmitBtn)
    $ret = $updateStepsForm.Controls.Add($updateStepsFormCancelBtn)
}

#
# jiraFindReplace - based on step info from selected steps to update, creates new fields objects for issues with updated steps
# @function
# @returns {Object} newIssuesStepFields - the updated issue step fields based on the steps chosen by the user
# 
function createNewStepsFieldsFromChosenSteps(){
    Write-Host "Updating steps"
    $selectedSteps = $updateStepsList.SelectedItems
    $newIssuesStepFields = @{}
    foreach($step in $selectedSteps) {
        $stepInfo = $step.SubItems
        $key = $step.text
        Write-Host "Saving selected step information for test: $key"
        $rowIndex = $stepInfo[1].text
        $columnIndex = $stepInfo[2].text
        $oldText = $stepInfo[3].text
        $newText = $stepInfo[4].text
        $testProp = $key.Replace("-", "")
        # Create property for issue in new issue steps fields if not found
        if($newIssuesStepFields.$testProp -eq $Null) {
            $newIssuesStepFields.$testProp = @{}
            $newIssuesStepFields.$testProp.key = $key
            $keyValueSaved = $newIssuesStepFields.$testProp.key
            Write-Host "key saved: $keyValueSaved"
            $newIssuesStepFields.$testProp.customfield_14883 = $issuesStepFields.$testProp.customfield_14883
        }

        # Update appropriate step in new issue steps field
        if($newIssuesStepFields.$testProp) {
            $newIssuesStepFields.$testProp.customfield_14883.stepsRows[$rowIndex].cells[$columnIndex] = $newText
        }
    }
   Return $(New-Object -TypeName PSObject -Prop $newIssuesStepFields)
}

#
# jiraFindReplace - updates chosen test steps in JIRA
# @function
# @param {Object} stepIssues - the issueSteps objects that need to be sent for updates (wrapper object containing updated customfields of tests needing updates)
#--------------------------------------------------------------------------------------------------------------------
# @property {Object} <testName without dash>
#    @property {string} key - the name of the test
#    @property {Object} customfield_11888 - holds the customfield object for test steps (with updated steps included)
# -------------------------------------------------------------------------------------------------------------------
function updateChosenSteps($stepIssues) {
    Write-Host "Update Chosen Steps"
    foreach($newIssue in $stepIssues.psobject.properties.name) {
        $key = $stepIssues.$newIssue.key
        $newStepsField = $stepIssues.$newIssue.customfield_14883
        $url = "issue/$key/"
        $updatedFieldsObj = @{"fields" = @{ "customfield_14883" = $newStepsField}}
        $updatedFieldsJSON = ConvertTo-Json $updatedFieldsObj -Depth 12

        Write-Host "Modification needed for test $key"
        Write-Host "Updating steps field with chosen text"
        Invoke-Expression -Command ".\runJIRARequest.ps1 -requestMethod PUT -requestUrl $url -requestBody '$updatedFieldsJSON'"
    }
}

#===========================================================================
# Set the popupForm handlers
#===========================================================================

#
# jiraFindReplace - click event handler for "Update Selected Steps" button on popup form
# @function
# 
$updateStepsFormSubmitBtn.Add_Click({
	if($updateStepsList.SelectedItems.Count -eq 0) {
        Write-Host "No steps selected for update"
		[System.Windows.Forms.MessageBox]::Show("You must select a JIRA step to update")
	}
	else {
        $stepIssues = createNewStepsFieldsFromChosenSteps
        updateChosenSteps -stepIssues $stepIssues

        Write-Host "Closing Form"
        $ret = $updateStepsList.Items.Clear()
        $ret = $updateStepsForm.Close()
        $filter = $query.text
        $issues = getIssues -filter $filter -start $global:startAt
        if($issues.issues.length -gt 0) {
            submitQuery
        }      
        
    }
})

#
# jiraFindReplace - click event handler for "Cancel" button on popup form
# @function
# 
$updateStepsFormCancelBtn.Add_Click({
    Write-Host "Closing Form"
    $ret = $updateStepsList.Items.Clear()
    $ret = $updateStepsForm.Close()       
})

#===========================================================================
# App logic
#=========================================================================== 

#
# jiraFindReplace - gets issues based on jira query provided
# @function
# @param {string} filter - jira query to use in GET request
# @param {number} start - jira start index to use in GET request
# @returns {Object} issues - the parsed response for the GET request
#
function getIssues($filter, $start){
    Write-Host ""
    Write-Host "Getting issues using query: $filter"
    #Query for tests
    $issues = Invoke-Expression -Command ".\getJIRASearchResult.ps1 -query '$filter' -max 200 -start $start"
    Write-Host "Finished"
    Write-Host ""
    Return $issues
}

#
# jiraFindReplace - populates issues list based on issues object passed
# @function
# @param {string} issues - an issues object from a GET request
#
function populateIssuesList($issues) {
    #$issuesStepFields = @{}
    Write-Host ""
    Write-Host "Populating issues list"
    $testList.Items.Clear()
    #Add tests to display
    foreach ($test in $issues.issues) {
        $key = $test.key
        $title = $test.fields.summary
        $testProp = $key.Replace("-", "")
        [void] $testList.Items.Add("$key : $title")

        # Add test to issuesStepsFields object for use later      
        $issuesStepFields.$testProp = @{}
        $issuesStepFields.$testProp.customfield_14883 = $test.fields.customfield_14883
    }
    $testList.SelectionMode = "Extended"
    Write-Host "Finished"
    Write-Host ""
}

#
# jiraFindReplace - event handler helper for query submission in app. 
# Gets the issues associated with the query and uses them to populate the issues list.
# Also hides query box and query submission button and makes second part of app visible
# @function
#	
function submitQuery() {
	$filter	= $query.text
    $issues = getIssues -filter $filter -start $global:startAt
    $global:startAt += 200
    if($issues.issues.length -gt 0) {
    	populateIssuesList -issues $issues

	    $testList.Visibility       = "Visible"
	    $findText.Visibility       = "Visible"
	    $replaceText.Visibility    = "Visible"
        $updateTests.Visibility    = "Visible"

	    $query.Visibility          = "Hidden"
        $searchForTests.Visibility = "Hidden"
    }
}

#
# jiraFindReplace - finds text to replace in issue steps and adds to updates table for possible update
# @function
# @param {Object} stepsField - the customfield_11888 field in the issue (holds the steps of the test)
# @param {string} textToFind - the text to find in the steps
# @param {string} textToReplace - the text to replace the text found in the steps with
# @param {string} key - the name of the issue
# @returns {number} numMods - the number of updates that would be made for that particular issue
#
function findTextInIssueRows($stepsField, $textToFind, $textToReplace, $key) {
    $numMods = 0
    Write-Host ""
    Write-Host "Checking rows of test $key"
    if($stepsField) {
        $rowsObj = $stepsField.stepsRows
        for($rowIndex=0; $rowIndex -lt $rowsObj.length; $rowIndex++) {
            $row = $rowsObj[$rowIndex].cells
            $columns = $row.length
            for($columnIndex=0; $columnIndex -lt $columns; $columnIndex++) {
                $oldText = $row[$columnIndex]
                $newText = $oldText.Replace($textToFind, $textToReplace)
                if($newText -cne $oldText) {
                    Write-Host "Found text $oldText. Would replace with $newText"
                    Write-Host "Adding update information to table"
                    addUpdatesToTable -key $key -rowIndex $rowIndex -columnIndex $columnIndex -oldText $oldText -newText $newText
                    $numMods++
                }
            }
        }
    }
    Write-Host "Finished"
    Write-Host ""
    return $numMods
}

#
# jiraFindReplace - searches for text to replace within user-selected tests
# If text is found, the information for the step is saved to a table for further user interactions
# @function
#	
function searchSelectedTests() {
	$selectedTests = $testList.SelectedItems
	$textToFind = $findText.text
    $textToReplace = $replaceText.text
    Write-Host ""
    Write-Host "Finding text $textToFind to replace with text $textToReplace"

	#Based on user selection, update appropriate tests
    foreach ($test in $selectedTests) {
        $key = $test.Split(":")[0].Trim()
        $testProp = $key.Replace("-", "")
        # Get steps field
        $stepsField = $issuesStepFields.$testProp.customfield_14883

        #Attempt to find text to replace in the issue 
        $numMods = findTextInIssueRows -stepsField $stepsField -textToFind $textToFind -textToReplace $textToReplace -key $key 
        #Update the issue steps field if needed
        if($numMods -eq 0) {
            Write-Host "No Modifications needed for test $key"
        } 
    }

    Write-Host "Finished"
    Write-Host ""

    #Reset app to run program again
	$query.Visibility            = "Visible"
    $searchForTests.Visibility   = "Visible"
	
    $findText.Visibility         = "Hidden"
    $replaceText.Visibility      = "Hidden"
    $testList.Visibility         = "Hidden"
    $updateTests.Visibility      = "Hidden"
}

#===========================================================================
# Get handles on the form objects
#===========================================================================

$query                       = $Form.FindName('query')
$searchForTests              = $Form.FindName('searchForTests')

$findText                    = $Form.FindName('findText')
$replaceText                 = $Form.FindName('replaceText')
$testList                    = $Form.FindName('testList')
$updateTests                 = $Form.FindName('updateTests')
$testList.SelectionMode      = "Extended"

$findText.Visibility         = "Hidden"
$replaceText.Visibility      = "Hidden"

$testList.Visibility         = "Hidden"
$updateTests.Visibility      = "Hidden"

#===========================================================================
# Set Main Form Event Handlers
#===========================================================================

#
# jiraFindReplace - click event handler for "Search for Tests" button on form (beneath jira query box)
# @function
# 
$searchForTests.Add_Click({
	if($query.text -eq "") {
		[System.Windows.Forms.MessageBox]::Show("You must enter a JIRA query to use for finding tests")
	}
	else {
        $global:startAt = 0
        submitQuery        
    }
})

#
# jiraFindReplace - click event handler for "Search Selected Tests for Text" button on form (beneath tests listbox)
# @function
# 
$updateTests.Add_Click({
	if($testList.SelectedItem.Count -eq 0) {
		[System.Windows.Forms.MessageBox]::Show("You Must Select Tests To Update")
	} elseif($findText.text -eq "") {
        [System.Windows.Forms.MessageBox]::Show("You Must Enter text to find in each test")
    } elseif($replaceText.text -eq "") {
        [System.Windows.Forms.MessageBox]::Show("You Must Enter text to replace found text with in each test")
    }
	else {
        createPopupTable
        searchSelectedTests
        [void]$updateStepsForm.ShowDialog()
    }
})
 
#===========================================================================
# Show form
#===========================================================================

$ret = 0
while($ret -ne 1) {
    $ret = Invoke-Expression -Command ".\setJiraAuth.ps1"
}
[void]$Form.ShowDialog()
