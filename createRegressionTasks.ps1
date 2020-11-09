#
# for creating project tasks (auto activities and bug logging(
# qaProject - the functional epic link for the testing release
# project - the JIRA project associated with the issue (EQAT, RQAT)
# sprintProj - the JIRA sprint project associated with the issue (ETEST, PTEST)
# epic - the automation epic link for the testing release
param(	
    [Parameter(Mandatory=$true)][string]$qaProject = "",
    [Parameter(Mandatory=$true)][string]$project = "",
    [Parameter(Mandatory=$true)][string]$sprintProj = "",
    [Parameter(Mandatory=$true)][string]$epic = ""
)

#use these powershell script units
. ".\jiraRequestUtilities.ps1"

$jiraConfig = Get-Content -Raw -Path "..\json\CONFIG_JSON\jiraConfig.json" | ConvertFrom-Json


$taskListXML = @"
<Window x:Name="mainWindow" x:Class="WpfApp2.MainWindow"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:local="clr-namespace:WpfApp2"
        mc:Ignorable="d"
        Title="Create Tasks" Height="300" Width="300" WindowStartupLocation="CenterScreen" Background="{DynamicResource {x:Static SystemColors.WindowBrushKey}}" ResizeMode="NoResize">

    <Grid>

        <Label Content="Assign Regression Tasks" HorizontalAlignment="Left" Margin="45,10,0,0" VerticalAlignment="Top" FontSize="16" FontWeight="Bold" Width="250"/>
        <ListBox x:Name="list" HorizontalAlignment="Left" Height="180" Margin="20,46,0,0" VerticalAlignment="Top" Width="250" Background="{DynamicResource {x:Static SystemColors.WindowBrushKey}}" BorderBrush="{DynamicResource {x:Static SystemColors.GrayTextBrushKey}}"/>

        <Button x:Name="createButton" Content="OK" Margin="100,235,0,0" VerticalAlignment="Top" Width="80" FontSize="14" HorizontalAlignment="Left"/>

    </Grid>
</Window>




"@  

#===========================================================================
# read XML to build windows form
#=========================================================================== 

$taskListXML = $taskListXML -replace 'mc:Ignorable="d"','' -replace "x:N",'N'  -replace '^<Win.*', '<Window'
[xml]$listXML = $taskListXML
#Read XAML
 
$infoReader=(New-Object System.Xml.XmlNodeReader $listXML)
try{
    $taskForm=[Windows.Markup.XamlReader]::Load( $infoReader )
}
catch{
    Write-Host "Unable to load Windows.Markup.XamlReader. Double-check syntax and ensure .net is installed."
}

#activate the objects

$list                    = $taskForm.FindName('list')
$createButton            = $taskForm.FindName('createButton')

$list.SelectionMode = "Multiple"


$createButton.Add_Click({
    
    $taskForm.Hide()

    $tms = $list.SelectedItems
    $projectInfo = Get-JiraIssue -issue $epic
    $summary = $projectInfo.fields.summary
    if($tms -ne $false){
        Foreach ($tm in $tms){
           # create activities task
           $ret = createTask -asn $tm -sum "$summary Activities" -desc $summary -epicLink $epic -qaProject $qaProject
           # create bug logging task
           $ret = createTask -asn $tm -sum "$summary Bug Logging" -desc $summary -epicLink $epic -qaProject $qaProject
        }
    }

})


function createTask($asn="", $sum="", $desc="", $epicLink="", $qaProject=""){

    Write-Host "=============================="
    Write-Host "Assigning regression task to $asn"
    Write-Host "=============================="

    $post = @{
                "fields" = @{
                                "project" = @{
                                    "key" = $project;
                                    };
                                "summary" = "$sum";
                                "description" = "$desc";
                                "issuetype" = @{
                                    "name" = "Task";
                                    };
                    }
    }
	if($asn) {
        $assigneeObj = @{
            "self" = "$($jiraConfig.bases.api)user?username=$asn"
            "name" = $asn;
            "key" = $asn
        };
        $ret = $post.fields.Add("assignee", $assigneeObj) 
	}
	if($epicLink) {			
        $ret = $post.fields.Add("customfield_11080", $epicLink)
	}
	if($qaProject -ne "") {
        $qaProjArr = @($qaProject)
        $ret = $post.fields.Add("customfield_11980", $qaProjArr)
	}

    $body = ConvertTo-Json $post

    $url = "issue/"
    $returnValueFromJIRA = & "C.\runJIRARequest.ps1" -requestMethod POST -requestUrl $url -requestBody $body

    $taskKey = $returnValueFromJIRA.key  
	$sprintID = getActiveSprint($sprintProj)
	if($sprintID) {
        $sprintPost = @{
            "issues" = @($taskKey)
        }
        $sprintBody = ConvertTo-Json $sprintPost
        # Assign to sprint
        $ret = & ".\runJIRARequest.ps1" -requestMethod POST -requestUrl "sprint/$sprintID/issue" -requestBody $sprintBody -apiBase "agile"
	}
    Write-Host "=============================="
	Write-Host "The Task is $taskKey"
    Write-Host "=============================="

}

$taskForm.ShowDialog()



