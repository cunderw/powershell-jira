[void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
[void][System.Reflection.Assembly]::LoadWithPartialName("System.Web.Extensions")        
$jsonserial= New-Object -TypeName System.Web.Script.Serialization.JavaScriptSerializer 
$jsonserial.MaxJsonLength = [int]::MaxValue

$jiraConfig = Get-Content -Raw -Path "..\jiraConfig.json" | ConvertFrom-Json
$apiBase = $jiraConfig.bases.api
$greenhopper = $jiraConfig.bases.greenhopper

Function Write-Host ($message,$nonewline,$backgroundcolor,$foregroundcolor) {
    $timestamp = Get-Date -Format "hh:mm:ss MM/dd/yy"
    $Message = "$timestamp [$env:computername] - $Message"
    $Message | Out-Host
}

#
# jiraRequestUtilities - prompts for test plan info for createTestPlan function if fails to create
# @function 
# @param {Object} testPlanInfo - the test plan info currently used for creation
# @return {Object} testPlanInfo - the updated test plan info to use for creation
#
function promptForTestPlanInfo($testPlanInfo) {
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    $form = New-Object System.Windows.Forms.Form
    $form.Text = 'Test Plan Information'
    $form.Size = New-Object System.Drawing.Size(300,600)
    $form.StartPosition = 'CenterScreen'

    $title = New-Object System.Windows.Forms.Label
    $title.Location = New-Object System.Drawing.Point(10,20)
    $title.Size = New-Object System.Drawing.Size(280,50)
    $title.Text = "TEST PLAN CREATION FAILED. `n`nPlease check entry information for errors. `nRe-enter information if inappropriate."
    $form.Controls.Add($title)

    $assigneeLabel = New-Object System.Windows.Forms.Label
    $assigneeLabel.Location = New-Object System.Drawing.Point(10,100)
    $assigneeLabel.Size = New-Object System.Drawing.Size(280,20)
    $assigneeLabel.Text = 'Assignee (username)'
    $form.Controls.Add($assigneeLabel)

    $assigneeTextBox = New-Object System.Windows.Forms.TextBox
    $assigneeTextBox.Location = New-Object System.Drawing.Point(10,120)
    $assigneeTextBox.Size = New-Object System.Drawing.Size(260,20)
    $form.Controls.Add($assigneeTextBox)

    $qaProjectLabel = New-Object System.Windows.Forms.Label
    $qaProjectLabel.Location = New-Object System.Drawing.Point(10,160)
    $qaProjectLabel.Size = New-Object System.Drawing.Size(280,20)
    $qaProjectLabel.Text = 'Epic Link'
    $form.Controls.Add($qaProjectLabel)

    $qaProjectTextBox = New-Object System.Windows.Forms.TextBox
    $qaProjectTextBox.Location = New-Object System.Drawing.Point(10,180)
    $qaProjectTextBox.Size = New-Object System.Drawing.Size(260,20)
    $form.Controls.Add($qaProjectTextBox)

    $summaryLabel = New-Object System.Windows.Forms.Label
    $summaryLabel.Location = New-Object System.Drawing.Point(10,220)
    $summaryLabel.Size = New-Object System.Drawing.Size(280,20)
    $summaryLabel.Text = 'TP Summary'
    $form.Controls.Add($summaryLabel)

    $summaryTextBox = New-Object System.Windows.Forms.TextBox
    $summaryTextBox.Location = New-Object System.Drawing.Point(10,240)
    $summaryTextBox.Size = New-Object System.Drawing.Size(260,20)
    $form.Controls.Add($summaryTextBox)

    $descriptionLabel = New-Object System.Windows.Forms.Label
    $descriptionLabel.Location = New-Object System.Drawing.Point(10,280)
    $descriptionLabel.Size = New-Object System.Drawing.Size(280,20)
    $descriptionLabel.Text = 'TP Description'
    $form.Controls.Add($descriptionLabel)

    $descTextBox = New-Object System.Windows.Forms.TextBox
    $descTextBox.Location = New-Object System.Drawing.Point(10,300)
    $descTextBox.Size = New-Object System.Drawing.Size(260,20)
    $form.Controls.Add($descTextBox)

    $appComponentLabel = New-Object System.Windows.Forms.Label
    $appComponentLabel.Location = New-Object System.Drawing.Point(10,340)
    $appComponentLabel.Size = New-Object System.Drawing.Size(280,20)
    $appComponentLabel.Text = 'App Component'
    $form.Controls.Add($appComponentLabel)

    $appComponentTextBox = New-Object System.Windows.Forms.TextBox
    $appComponentTextBox.Location = New-Object System.Drawing.Point(10,360)
    $appComponentTextBox.Size = New-Object System.Drawing.Size(260,20)
    $form.Controls.Add($appComponentTextBox)

    $projectLabel = New-Object System.Windows.Forms.Label
    $projectLabel.Location = New-Object System.Drawing.Point(10,400)
    $projectLabel.Size = New-Object System.Drawing.Size(280,20)
    $projectLabel.Text = 'Project'
    $form.Controls.Add($projectLabel)

    $projectTextBox = New-Object System.Windows.Forms.TextBox
    $projectTextBox.Location = New-Object System.Drawing.Point(10,420)
    $projectTextBox.Size = New-Object System.Drawing.Size(260,20)
    $form.Controls.Add($projectTextBox)

    $OKButton = New-Object System.Windows.Forms.Button
    $OKButton.Location = New-Object System.Drawing.Point(100,520)
    $OKButton.Size = New-Object System.Drawing.Size(75,23)
    $OKButton.Text = 'Resubmit'
    $OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.AcceptButton = $OKButton
    $form.Controls.Add($OKButton)

    # set defaults
    $assigneeTextBox.Text = $testPlanInfo.assignee
    $qaProjectTextBox.Text = $testPlanInfo.epicLink
    $descTextBox.Text = $testPlanInfo.desc
    $summaryTextBox.Text = $testPlanInfo.summary
    $appComponentTextBox.Text = $testPlanInfo.appComponent
    $projectTextBox.Text = $testPlanInfo.project
    $form.Topmost = $true
    $result = $form.ShowDialog()

    if ($result -eq [System.Windows.Forms.DialogResult]::OK)
    {
        $testPlanInfo.assignee = $assigneeTextBox.Text
        $testPlanInfo.epicLink = $qaProjectTextBox.Text
        $testPlanInfo.desc = $descTextBox.Text
        $testPlanInfo.summary = $summaryTextBox.Text
        $testPlanInfo.appComponent = $appComponentTextBox.Text
        $testPlanInfo.project = $projectTextBox.Text
    }
    return $testPlanInfo
}

#
# jiraRequestUtilities - converts to safe uri
# @function 
# @param {string} uri - the uri component to convert
# @return {string} uri - converted to safe uri
#
Function ConvertTo-SafeUri($uri) {
    Return [System.Uri]::EscapeDataString($uri)
}

#
# jiraRequestUtilities - gets search result for jira issue using query
# @function 
# @param {string} query - jql query to use for getting jira isssues
# @param {string} [max=5000] - max amount to return
# @param {string} [start=0] - index to start at for returned amount
# @returns {Object} responseVal - the related jira issues
#
Function Get-JiraSearchResult($query, $max=5000, $start=0, $fields="") {
    $queryStr = "search?jql=$(ConvertTo-SafeUri $query)&maxResults=$max&startAt=$start"
    if($fields -ne "") {
        $queryStr += "&fields=$fields"
    }
    $responseVal = & ".\runJIRARequest.ps1" -requestMethod GET -requestUrl $queryStr
    Return $responseVal
}

#
# jiraRequestUtilities - gets jira issue using issue name
# @function 
# @param {string} issue - name of issue
# @returns {Object} responseVal - the jira issue's data
#
Function Get-JiraIssue($issue) {
    $responseVal = & ".\runJIRARequest.ps1" -requestMethod GET -requestUrl "issue/$(ConvertTo-SafeUri $issue)"
    Return $responseVal
}

#
# jiraRequestUtilities - gets jira issue creation format
# @function 
# @param {string} project - name of issue
# @param {string} issuetype - name of issue
# @returns {Object} responseVal - the correct format for the meta data
#
Function getIssueCreationFormat($project, $issuetype){
    $queryStr = "issue/createmeta?projectKeys=$(ConvertTo-SafeUri $project)&issuetypeNames=$(ConvertTo-SafeUri $issuetype)&expand=projects.issuetypes.fields"
    $responseVal = & ".\runJIRARequest.ps1" -requestMethod GET -requestUrl $queryStr
    $retText = ConvertTo-Json $responseVal -Depth 10
    if($responseVal) {
        Write-Host "Creation available to user. See format below`n`n$retText"
    } else {
        Write-Host "Creation not available to user for project $project issuetype $issuetype"
    }
    Return $responseVal
}

#
# jiraRequestUtilities - post a comment
# @function 
# @param {string} issue - name of issue to comment to
# @param {string} comment - comment to make
#
function postComment($issue, $comment) {
    $url = "issue/$issue/comment"
    $body = '{ "body": "' + $comment + '" }';
    Write-Host "Posting to issue $issue comment: $comment"
    $ret = & ".\runJIRARequest.ps1" -requestMethod "POST" -requestUrl "$url" -requestBody $body
}

#
# jiraRequestUtilities - starts test plan
# @function 
# @param {string} testPlan - the test plan to start
# @return {object} returnValueFromJIRA - generated return value for request to start test plan
#
Function startTestPlan($testPlan) {
    Write-Host ""
    Write-Host "-------------- START jiraRequestUtilities.startTestPlan -----------------"
    Write-Host ""
    Write-Host "Getting test plan status"
    $issue = Get-JiraIssue -issue $testPlan
    $status = $issue.fields.status.name
    if($status -eq "Closed") {
        Write-Host "Test Plan Closed. Reopening closed test plan"
        $url = "issue/$(ConvertTo-SafeUri $testPlan)/transitions?expand=transitions.fields"
        $body = "{""transition"": { ""id"": ""61""}}"
        $returnValueFromJIRA = & ".\runJIRARequest.ps1" -requestMethod POST -requestUrl $url -requestBody $body
    }

    $url = "issue/$(ConvertTo-SafeUri $testPlan)/transitions?expand=transitions.fields"
    $body = "{""transition"": { ""id"": ""11""}}"
    $returnValueFromJIRA = & ".\runJIRARequest.ps1" -requestMethod POST -requestUrl $url -requestBody $body
    Write-Host ""
    Write-Host "-------------- END jiraRequestUtilities.startTestPlan -----------------"
    Write-Host ""
    Return $returnValueFromJIRA
}

#
# jiraRequestUtilities - gets the active sprint id for project
# @function 
# @param {string} project - project to use
# -----------------------------------------
#   project Options
#   1. ETEST
#   2. PTEST
#   3. MTEST
# -----------------------------------------
Function getActiveSprint($project) {
    Write-Host ""
    Write-Host "-------------- START jiraRequestUtilities.getActiveSprint -----------------"
    Write-Host ""
    if($project -eq "APITEST") {
        if($appComponent -eq "OPS") {
            $project = "PTEST"
        } else {
            $project = "ETEST"
        }
    }
	Write-Host "Getting Active Sprint For Project: $project" 
	if($project -eq "ETEST" -or $project -eq "YTEST") {
        $url = 'sprintquery/533'
        $contains = "eComm"
	}
	elseif($project -eq "PTEST") {
        $url = 'sprintquery/527'
        $contains = "Retail"
	}
	elseif($project -eq "MTEST") {
        $url = 'sprintquery/708'
        $contains = "Mobile"
	}
	elseif($project -eq "TTEST") {
        $url = 'sprintquery/534'
        $contains = "TNQAT"
	}
	$responce =  & ".\runJIRARequest.ps1" -requestMethod GET -requestUrl $url -apiBase "greenhopper"
	$sprints =  $responce.sprints
	foreach($sprint in $sprints) {
	 if($sprint.state -eq "ACTIVE" -and $sprint.name -like "$($contains)*") {
		$activeSprint = $sprint.id
		break
	 }
	}
	Write-Host "Active Sprint: $activeSprint" 
    Write-Host ""
    Write-Host "-------------- END jiraRequestUtilities.getActiveSprint -----------------"
    Write-Host ""
	Return $activeSprint
}

#
# jiraRequestUtilities - creates a blank test plan
# @function 
# @param {string} project - project to use
# -----------------------------------------
#   project Options
#   1. ETEST
#   2. PTEST
#   3. MTEST
# -----------------------------------------
# @param {string} summary - summary of the test plan (usually the release version) 
# @param {string} description - description of the test plan type (based on filter used)
# @param {string} assignee - assignee
# @param {string} qaProject - E-XXX
# @param {number} sprintID - Sprint ID from JIRA
# @param {string} epicLink - EQAT-XXXX
# @param {string} appComponent - Component for the app. OS 3.0, RWD, OPS, etc...
# @param {boolean} [prompt=true] - if false, does not prompt on retry
# @return {string} testPlan - the test plan key
#
Function createTestPlan($project, $summary, $description, $assignee, $qaProject="", $sprintID, $epicLink="", $appComponent, $prompt=$true){
    $returnValueFromJIRA = $null
    $tries = 0
    Write-Host ""
    Write-Host "-------------- START jiraRequestUtilities.createTestPlan -----------------"
    Write-Host ""
    while($returnValueFromJIRA.key -eq $null -and $tries -lt 3) {
        $post = @{}
        if($project -eq "YTEST") {
            $post = @{
                "fields" = @{
			        "project" = @{
			    	    "key" = $project;
			         };
			        "summary" = $summary;
			        "description" =  $description;
			        "issuetype" = @{
			    	    "name" = "Test Plan";
			         }
			        
                }
		    }
        } else {
            $post = @{
                "fields" = @{
			        "project" = @{
			    	    "key" = $project;
			         };
			        "summary" = $summary;
			        "description" =  $description;
			        "issuetype" = @{
			    	    "name" = "Test Plan";
			         };
			         "components" = @(
			    	    @{
			    		    "name" = $appComponent;
			    	    };
			        );
                }
		    }
        }
        
		if($assignee) {
            $assigneeObj = @{
                "self" = $apiBase + "user?username=$assignee";
                "name" = $assignee;
                "key" = $assignee
            };
            $ret = $post.fields.Add("assignee", $assigneeObj) 
		}
		if($qaProject -ne "") {
            $qaProjArr = @($qaProject)
            $ret = $post.fields.Add("customfield_11980", $qaProjArr)
		}
		if($epicLink -ne "") {			
            $ret = $post.fields.Add("customfield_11080", $epicLink)
		}
		$body = ConvertTo-Json $post -Depth 4
		$url = "issue/"
        Write-Host $body
		$returnValueFromJIRA = & ".\runJIRARequest.ps1" -requestMethod POST -requestUrl $url -requestBody $body
		if($returnValueFromJIRA.key -eq $null) {
            $testPlanInfo = @{
                "assignee" = $assignee;
                "epicLink" = $epicLink;
                "desc" = $description;
                "summary" = $summary;
                "appComponent" = $appComponent;
                "project" = $project
            }
            $qaProject = ""
            if($prompt) {
                $testPlanInfo = promptForTestPlanInfo -testPlanInfo $testPlanInfo
                $assignee = $testPlanInfo.assignee
                $epicLink = $testPlanInfo.epicLink
                $description = $testPlanInfo.desc
                $summary = $testPlanInfo.summary
                $appComponent = $testPlanInfo.appComponent
                $project = $testPlanInfo.project
                $qaProject = ""
            }
		}
        $tries++
        sleep -Seconds 5
	}
	$testPlan = $returnValueFromJIRA.key  
	$sprintID = getActiveSprint($project)
	if($sprintID) {
       $sprintPost = @{
            "issues" = @($testPlan)
       }
       $sprintBody = ConvertTo-Json $sprintPost
       # Assign to sprint
       $ret = & ".\runJIRARequest.ps1" -requestMethod POST -requestUrl "sprint/$sprintID/issue" -requestBody $sprintBody -apiBase "agile"
	}
	Write-Host "Test Plan name is $testPlan"
    Write-Host ""
    Write-Host "-------------- END jiraRequestUtilities.createTestPlan -----------------"
    Write-Host ""
	Return $testPlan
}


#
# jiraRequestUtilities - clones a jira tct
# @function
# @param {string} issue - the issue to clone
# @param {string} [labelForIssue = "AUTO_CLONE_JEFF"] - whatever label you wish to attach to the clones
# @param {string} [assignee=""] - the assignee for tct (if left blank, then will clone assignee
function cloneJIRATCT($issue, $labelForIssue = "AUTO_CLONE_JEFF", $assignee="") {
    $issueInfo = Get-JiraIssue -issue $issue
    Write-Host "Attempting to clone issue $issue"
    $labels = $issueInfo.fields.labels
    [array]$labels += $labelForIssue
    $components = $issueInfo.fields.components
    $project = $issueInfo.fields.project.key
    $summary = $issueInfo.fields.summary
    $priority = $issueInfo.fields.priority.name
    $qaTestCaseType = $issueInfo.fields.customfield_12680.value
    $description = $issueInfo.fields.description
    $qaExecutionEstimate = $issueInfo.fields.customfield_12580
    $qaPrepEstimate = $issueInfo.fields.customfield_12880
    $testSteps = $issueInfo.fields.customfield_14883
    $issueType = $issueInfo.fields.issuetype.name
    $timeTracking = $issueInfo.fields.timetracking
    if($reporter -eq "") {
        $reporter = $issueInfo.fields.reporter.name
    }
    if($assignee -eq "") {
        $assignee = $issueInfo.fields.assignee.name
    }
    $body = @{
                "fields" = @{
			        "project" = @{
			    	    "key" = $project
			         };
			        "summary" = $summary;
			        "description" =  $description;
			        "issuetype" = @{
			    	    "name" = $issueType
			         };
                     "customfield_12680" = @(@{
                        "value" = $qaTestCaseType
                     });
                     "labels" = $labels;
			         "components" = $components;
                     "assignee" =  @{
                        "name" = $assignee
                     };
                     "priority" = @{
                        "name" = $priority
                     };
                     "customfield_12880" = $qaExecutionEstimate;
			         "customfield_12580" = $qaPrepEstimate;
                     "customfield_14883" = $testSteps;
                     "timetracking" = $timeTracking
                }
              }
    $post = ConvertTo-Json -InputObject $body -Depth 20
    # create issue
    $ret = & ".\runJIRARequest.ps1" -requestMethod POST -requestUrl "issue" -requestBody $post
    $newTestKey = $ret.key
    # link created issue as clone
    linkIssue -issueToLink "$issue" -issue "$newTestKey" -linkType "Cloners"
    Write-Host "New test $newTestKey created"
}

#
# jiraRequestUtilities - links one issue to another
# @function
# @param {string} issueToLink - the bug to link to the issue
# @param {string} issue - the issue to link the bug to
# @param {string} [linkType = "Defect"] - the type of link to use
function linkIssue($issueToLink, $issue, $linkType = "Defect") {
  $body = @{
    "type"= @{
      "name"= $linkType
    };
    "inwardIssue"= @{
      "key"= $issue
    };
    "outwardIssue"= @{
      "key"= $issueToLink
    }
  };
  
  Write-Host "Linking issue $issueToLink to issue $issue with link $linkType";
  $post = ConvertTo-Json $body -Depth 4
  $ret = & ".\runJIRARequest.ps1" -requestMethod POST -requestUrl "issueLink" -requestBody $post
}


#
# jiraRequestUtilities - Checks Test Case Templates from filter to make sure not already in test plan before adding
# @function
# @param {string} tctInProjFilter - the filter to get the TCTs from the project with the filter specified
# @param {string} tcInProjInTPFilter - the filter to use to get the Test cases already in the Test Plan that match that project 
#
function checkTCTInTestPlan($tctInProjFilter, $tcInProjInTPFilter) {
    Write-Host ""
    Write-Host "-------------- START jiraRequestUtilities.checkTCTInTestPlan -----------------"
    Write-Host ""
    Write-Host "Getting Test Cases in Test Plan using query $tcInProjInTPFilter"
    $startNum = 0
    # get first 100 tests
    $testsInProjInTestPlan = Get-JiraSearchResult -query $tcInProjInTPFilter -fields 'customfield_12480,customfield_11883' -max 100 -start $startNum
    Write-Host ""
    $testsInPlan = [System.Collections.ArrayList]@()
    $tcProjs = [System.Collections.ArrayList]@()
    $allTestsThatShouldBeInPlan = [System.Collections.ArrayList]@()
    $testsNotInPlan = [System.Collections.ArrayList]@()
    # Get all the tests in the test plan that have already been added
    Write-Host ""
    Write-Host "--------------------------------------"
    Write-Host "Checking for tests found in test plan"
    do{
        foreach($test in $testsInProjInTestPlan.issues) {
            $testName = $test.fields.customfield_11883
            $qaatProject = $test.fields.customfield_12480
            Write-Host "Test $testName with QA Automation Project = $qaatProject in Test Plan"
            $ret = $testsInPlan.Add($testName);
            $ret = $tcProjs.Add($qaatProject);

        }
            #set this var to get the next 100 tests from the JIRA API
        $startNum = $startNum + 100
    
        #get the next 100 tests and go back through the loop
        $testsInProjInTestPlan = Get-JiraSearchResult -query $tcInProjInTPFilter -fields 'customfield_12480,customfield_11883' -max 100 -start $startNum
    
    } while($testsInProjInTestPlan.issues.length -gt 0)
    Write-Host "Requesting Additional Information concerning tests to be added using query $tctInProjFilter"
    $startNum = 0
    $testsInProj =  Get-JiraSearchResult -query $tctInProjFilter -fields 'customfield_12480,customfield_11883' -max 100 -start $startNum
    do{
        foreach($test in $testsInProj.issues) {
            $key = $test.key
            $qaatProject = $test.fields.customfield_12480
            $currTest = @{
                "key" = $key;
                "qaatProject" = $qaatProject
            }
            $ret = $allTestsThatShouldBeInPlan.Add($currTest);


        }
            #set this var to get the next 100 tests from the JIRA API
        $startNum = $startNum + 100
    
        #get the next 100 tests and go back through the loop
        $testsInProj = Get-JiraSearchResult -query $tctInProjFilter -fields 'customfield_12480,customfield_11883' -max 100 -start $startNum
    } while($testsInProj.issues.length -gt 0)
    # Get all the tests that need to be added to test plan
    for($index = 0; $index -lt $allTestsThatShouldBeInPlan.Count; $index++) {
        $test = $allTestsThatShouldBeInPlan[$index]
        Write-Host "Searching for Test $($test.key) in Test Plan"
        #If the TCT is not in the test plan, add it to the list of TCTs that need to be added
        $indexOfTC = $testsInPlan.IndexOf($test.key)
        if($indexOfTC -eq -1) {
            Write-Host "Test $($test.key) NOT found in Test Plan"
            $ret =$testsNotInPlan.Add($test.key)
        } else {
            $qaatProject = $tcProjs[$indexOfTC]
            if($qaatProject -eq $test.qaatProject) {
                Write-Host "Test $($test.key) found in Test Plan!"
            } else {
                Write-Host "QA Automation Project not set for TC associated with $($test.key) in TP." 
                $ret =$testsNotInPlan.Add($test.key)
            }
        }
    }
    Write-Host "--------------------------------------"
    Write-Host ""
    $testsToAdd = $testsNotInPlan -join ","
    Write-Host "-------------- END jiraRequestUtilities.checkTCTInTestPlan -----------------"
    Write-Host ""
    return $testsToAdd
}

#
# jiraRequestUtilities - adds tests from query to test plan
# @function 
# @param {string} testPlan - the test plan to add to
# @param {string} query - the JIRA query to use
# @param {Object} testCases - list of testCase issues added
# @param {number} [attemptNumber=0] - this is for the looped portion (DO NOT pass in); tells what attempt the function is on for calls
# @param {number} [originalQuery=""] - this is for the looped portion (DO NOT pass in); tells what the original query was for getting the test cases to add
# @return {object} returnedTestCases - test case templates added to test plan
#
function addTestsToTestPlan($testPlan, $query, $attemptNumber = 0, $originalQuery="") {
    Write-Host ""
    Write-Host "-------------- START jiraRequestUtilities.addTestsToTestPlan -----------------"
    Write-Host ""
    Write-Host "Adding tests to TP $testPlan with query $query"
	$testCases = Get-JiraSearchResult -query $query -fields 'customfield_12480,customfield_12680,description,labels,components,summary,project,customfield_14883' 
    $plan = Get-JiraIssue $testPlan
    $qaatProject = ""

    foreach($test in $testCases.issues) {
        $qaatProject = $test.fields.customfield_12480
        $description = $test.fields.description;
        if($description -eq $null) {
            $description = ""
        }
        $qaTestType = $test.fields.customfield_12680;
        $components = $test.fields.components;
        for($compIndex = 0; $compIndex -lt $components.length; $compIndex++) {
            $compName = $components[$compIndex].name;
            if($compName.Contains("`?`?")) {
                write-host "WARNING: component '$compName' not formatted correctly!!!"
            }
        }
        $post = @{ 
            "fields" = @{
                 "project" = @{
                    "key" = $test.fields.project.key
                  };
                  "summary" = $test.fields.summary;
                  "description" = $description;
                  "labels" = $test.fields.labels;
                  "components" = $test.fields.components;
                  "issuetype" = @{
                    "name" = "Test Case"
                  };
                 "parent" = @{
                    "key" = "$testPlan"
                  };
                 "customfield_11883" = $test.key;
                 "customfield_12480" = $qaatProject;
                 "customfield_12680" = $qaTestType;
             }
         }
         if($test.fields.customfield_14883 -ne $null) {
            $post.fields.customfield_14883 = $test.fields.customfield_14883
         }
        $body = ConvertTo-Json $post -Depth 20
        $returnValueFromJIRA = & ".\runJIRARequest.ps1" -requestMethod POST -requestUrl 'issue' -requestBody $body
        Write-Host "Test Case Addition Response: $returnValueFromJIRA"
    }
    $qaatProject = $qaatProject.Replace("\", "\\")
    $qaatProject = $qaatProject.Replace("`"", "")
    $issuesToAdd = "Waiting for issues to add"
    Write-Host $issuesToAdd
    $timer = 0
    #Wait for additions to add
    while($issuesToAdd -ne "" -and $timer -lt 5) {
        $sleepTime = Get-Random -Maximum 9 -Minimum 1
        sleep -Seconds $sleepTime
        $issuesToAdd = checkTCTInTestPlan -tctInProjFilter $query -tcInProjInTPFilter "parent = $testPlan and status = Open and `"QA Automation Project`" ~ `"$qaatProject`""
        Write-Host "The following issues not added: $issuesToAdd"
        $timer++
    }
    # Try again to add if additions failed. Try up to 3 times
    if($issuesToAdd -ne "" -and $attemptNumber -lt 3) {
        if($originalQuery -eq "") {
            $originalQuery = $query
        }
        $query = "key in ($($issuesToAdd))"
        $attemptNumber++
        $ret = addTestsToTestPlan -testPlan $testPlan -query $query -attemptNumber $attemptNumber -originalQuery $originalQuery
    } else {
        if($originalQuery -eq "") {
            $returnedTestCases = $testCases
        } else {
            $returnedTestCases = Get-JiraSearchResult -query $originalQuery -fields 'customfield_12480'
        }
        if($issuesToAdd -ne "") {
            Write-Host ""
            Write-Host "WARNING: WHILE ATTEMPTING TO ADD TESTS TO TESTPLAN, THE FOLLOWING ISSUES FAILED TO ADD:"
            Write-Host "$issuesToAdd"
            Write-Host ""
            $ret = postComment -issue $testPlan -comment "Tests failed to add: $issuesToAdd"
        }
    }
    Write-Host "Waiting 15 seconds after additions complete"
    sleep -Seconds 15
    Write-Host ""
    Write-Host "-------------- END jiraRequestUtilities.addTestsToTestPlan -----------------"
    Write-Host ""
    return $returnedTestCases
}

function addTestsToTestPlanOld($testPlan, $query) {
    Write-Host "Adding tests to TP $testPlan with query $query"
    $tcts= Get-JiraSearchResult -query $query
    $url = 'issue'
    foreach($test in $tcts.issues) {
        $post = @{ 
            "fields" = @{
                 "project" = @{
                    "key" = $test.fields.project.key
                  };
                  "summary" = $test.fields.summary;
                  "description" = $test.fields.description;
                  "issuetype" = @{
                    "name" = "Test Case"
                  };
                 "parent" = @{
                    "key" = "$testPlan"
                  };
                 "customfield_11883" = $test.key;
                 "customfield_14784" = $test.fields.customfield_14784
             }
         }
        $body = ConvertTo-Json $post -Depth 20
        $returnValueFromJIRA = & ".\runJIRARequest.ps1" -requestMethod POST -requestUrl $url -requestBody $body
    }
    return $testCases
}

#
# jiraRequestUtilities - stops test plan
# @function 
# @param {string} testPlan - the test plan to stop
# @return {object} returnValueFromJIRA - generated return value for request to stop test plan
#
Function stopTestPlan($testPlan) {
    Write-Host ""
    Write-Host "-------------- START jiraRequestUtilities.stopTestPlan -----------------"
    Write-Host ""
    $url = "issue/$(ConvertTo-SafeUri $testPlan)/transitions?expand=transitions.fields"
    $body = "{""transition"": { ""id"": ""51""}}"
    $returnValueFromJIRA = & ".\runJIRARequest.ps1" -requestMethod POST -requestUrl $url -requestBody $body
    Write-Host ""
    Write-Host "-------------- END jiraRequestUtilities.stopTestPlan -----------------"
    Write-Host ""
    Return $returnValueFromJIRA
}

#
# jiraRequestUtilities - completes test plan if possible
# @function 
# @param {string} testPlan - the test plan to complete
# @return {object} returnValueFromJIRA - generated return value for request to complete test plan
#
Function completeTestPlan($testPlan) {
    Write-Host ""
    Write-Host "-------------- START jiraRequestUtilities.completeTestPlan -----------------"
    Write-Host ""
    $testPlanData = Get-JiraIssue -issue $testPlan
    $returnValueFromJIRA = $null
    if($testPlanData.fields.customfield_11890 -eq "100.0") {
        try{
            $url = "issue/$(ConvertTo-SafeUri $testPlan)/transitions?expand=transitions.fields"
            $body = "{""transition"": { ""id"": ""71""}}"
            $returnValueFromJIRA = & ".\runJIRARequest.ps1" -requestMethod POST -requestUrl $url -requestBody $body
        } catch {
            Write-Host "Was Not able to Complete Test Plan."
        } 
    }
    Write-Host ""
    Write-Host "-------------- END jiraRequestUtilities.completeTestPlan -----------------"
    Write-Host ""
    Return $returnValueFromJIRA
}

#
# jiraRequestUtilities - gets all components for a project. Ex ptest, etest, etc
# @function 
# @param {string} project - PTEST, ETEST, etc
#
Function getComponentList($project){
    Write-Host ""
    Write-Host "-------------- START jiraRequestUtilities.getComponentList -----------------"
    Write-Host ""

    $project = $project.ToUpper()

    Write-Host "Getting a list of components from project: $project"

    $components = & ".\runJIRARequest.ps1" -requestMethod "GET" -requestUrl "project/$project/components"

    $array = @()

    foreach($component in $components){

        $array += $component.name

    }
    Write-Host ""
    Write-Host "-------------- END jiraRequestUtilities.getComponentList -----------------"
    Write-Host ""

    return $array
}

#
# jiraRequestUtilities - queries for issues and gets all the components off of them
# @function 
# @param {string} query - 'project = OPSR and status = "Ready For QA"'
#
Function getAllComponentsFromJQL($query){
    Write-Host ""
    Write-Host "-------------- START jiraRequestUtilities.getAllComponentsFromJQL -----------------"
    Write-Host ""
    Write-Host "Running JQL to get components from issues."
    $issues = Get-JiraSearchResult -query $query

    $array = @()

    Write-Host "Creating array of components."

    Foreach($issue in $issues.issues){

        #Write-Host $issue.key

        $components = $issue.fields.components

        Foreach($component in $components){

            $array += $component.name

        }
    

    }

    Write-Host "Removing duplicates."
    $array = $array | select -uniq

    Write-Host "Sorting the list alphabetically."
    $array = $array | Sort-Object

    Write-Host ""
    Write-Host "Components: $array"
    Write-Host ""
    Write-Host ""
    Write-Host "-------------- END jiraRequestUtilities.getAllComponentsFromJQL -----------------"
    Write-Host ""
    return $array
}

#
# jiraRequestUtilities - transitions test case to pass / fail
# @function 
# @param {string} testCase - name of test to transition
# @param {string} status - status test is currently in
# @param {string} transition - state to transition test to
#
function transitionTestCase($testCase, $status, $transition){
    Write-Host "Transitioning $testCase with status $status using transition $transition"
    $postPASS = '{ "fields": {"customfield_13680": 1} , "transition": { "id": "21"} }'
    $postFAIL = '{ "fields": {"customfield_13680": 1} , "transition": { "id": "271"} }'
    $postTEST = '{"transition": { "id": "11"}}'
    $postRETEST = '{"transition": { "id": "81"}}'
    $postRETESTPASS = '{"transition": { "id": "91"}}'
    $postRETESTFAIL = '{"transition": { "id": "51"}}'
    $postNEEDSAUTOREVIEW  = '{"update": { "comment": [ { "add": {"body": "Automated Test Failed. Please Investigate."} }]},"fields": {}, "transition": { "id": "161"}}'
    $postINAUTOREVIEW = '{"transition": { "id": "221"}}'
    $postNEEDSFunctionalReview = '{"update": { "comment": [ { "add": {"body": "Automated API Test Failed. Please Investigate."} }]},"fields": {}, "transition": { "id": "171"}}' 
    #$postDELETEISSUE = '{"transition": { "id": "381"}}'
    $postDELETEISSUE = '{"transition": { "id": "401"}}'
    $url = "issue/$($testCase)/transitions?expand=transitions.fields"
    if($status -eq "Pass" -or $status -eq "Fail" -or $status -eq "Retest") {
        if($status -eq "Fail") {
            # Hit retest
            Write-Host "Hitting Retest for failed test"
            $ret = & ".\runJIRARequest.ps1" -requestMethod "POST" -requestUrl "$url" -requestBody $postRETESTFAIL
        } elseif($status -eq "Pass") {
            # Hit retest
            Write-Host "Hitting Retest for passed test"
            $ret = & ".\runJIRARequest.ps1" -requestMethod "POST" -requestUrl "$url" -requestBody $postRETESTPASS
        }
        sleep -Seconds 1
        # Hit test
        Write-Host "Hitting Test to Retest $testCase"
        $ret = & ".\runJIRARequest.ps1" -requestMethod "POST" -requestUrl "$url" -requestBody $postRETEST
    }
    if($status -eq "Open" -and $transition -ne "Test") {
        Write-Host "Hitting Test"
        $ret = & ".\runJIRARequest.ps1" -requestMethod "POST" -requestUrl "$url" -requestBody $postTEST
    }
    if($transition -eq "Fail") {
        Write-Host "Moving to Needs Automation Review"
        $ret = & ".\runJIRARequest.ps1" -requestMethod "POST" -requestUrl "$url" -requestBody $postNEEDSAUTOREVIEW
        sleep -Seconds 1
        Write-Host "Moving to In Automation Review"
        $ret = & ".\runJIRARequest.ps1" -requestMethod "POST" -requestUrl "$url" -requestBody $postINAUTOREVIEW
    }
    switch($transition) {
        'Pass' {
            Write-Host "Passing test"
            $ret = & ".\runJIRARequest.ps1" -requestMethod "POST" -requestUrl "$url" -requestBody $postPASS
            break;
        }
        'Fail' {
            Write-Host "Failing test"
            $ret = & ".\runJIRARequest.ps1" -requestMethod "POST" -requestUrl "$url" -requestBody $postFAIL
            break;
        }   
        'Needs Functional Review' {
            Write-Host "Failing test"
            $ret = & ".\runJIRARequest.ps1" -requestMethod "POST" -requestUrl "$url" -requestBody $postNEEDSFunctionalReview
            break;
        } 
        'Delete' {
            Write-Host "Deleting issue"
            $ret = & ".\runJIRARequest.ps1" -requestMethod "POST" -requestUrl "$url" -requestBody $postDELETEISSUE
            break;
        }     
    }
}

#
# jiraRequestUtilities - gets test plan progress info
# @author {LA}
# @param {string} testPlan - key for test plan
#
function getTestPlanProgress($testPlan){
    $passedQuery = "parent = $testPlan and status = 'Pass'"
    $tpResponse = Get-JiraSearchResult -query $passedQuery -fields "key"
    $numPassed = $tpResponse.total

    $failedQuery = "parent = $testPlan and status = 'Fail'"
    $tpResponse = Get-JiraSearchResult -query $failedQuery -fields "key"
    $numFailed = $tpResponse.total

    $numCompletedQuery = "parent = $testPlan and statusCategory = Done and status != Deleted"
    $numTestsInPlanQuery = "parent = $testPlan and status != Deleted"

    $numCompletedResponse = Get-JiraSearchResult -query $numCompletedQuery -fields "key"
    $numTestsInPlanResponse = Get-JiraSearchResult -query $numTestsInPlanQuery -fields "key"
    $percentComplete = ($numCompletedResponse.total/$numTestsInPlanResponse.total) * 100
    $progress = @{
        "testPlan" = $testPlan;
        "passed" = $numPassed;
        "failed" = $numFailed;
        "percentComplete" = [math]::Round($percentComplete, 2)
    }
    return $progress
}

#
# jiraRequestUtilities - gets issues blocking tests
# @function 
# @param {string} [appComponent="OPS"] - the app component
#
function getIssuesBlockingTests($appComponent=""){
    if($appComponent -eq "OPS") {
        $query = "type = bug and resolution is empty and issueFunction in linkedIssuesOf(`"project = PTEST and type = 'Test Case Template' and status = 'Automated Test'`")"
    } elseif($appComponent -eq "YETI") {
        $query = "type = bug and resolution is empty and issueFunction in linkedIssuesOf(`"project = YTEST and type = 'Test Case Template' and status = 'Automated Test'`")"
    } else {
        $query = "type = bug and resolution is empty and issueFunction in linkedIssuesOf(`"project = ETEST and component = '$appComponent' and type = 'Test Case Template' and status = 'Automated Test'`")"
    }
    $returnedIssues = Get-JiraSearchResult -query $query -fields "key,summary"
    $issuesBlockingTestsArr = [System.Collections.ArrayList]@()
    foreach($issue in $returnedIssues.issues) {
        $key = $issue.key
        $summary = $issue.fields.summary
        $linkedIssuesQuery = "Type = `"Test Case Template`" and status = `"Automated Test`" and issueFunction in linkedIssuesOf(`"id = $key`")"
        $linkedIssues = Get-JiraSearchResult -query $linkedIssuesQuery -fields "key"
        $linkedIssuesCount = $linkedIssues.total 
        $linkedIssuesUrl = ""
        $bugLinkStr = "<a href=`"`">$key</a>"
        $linkedIssuesStr = "<a href=`"$linkedIssuesUrl`">$linkedIssuesCount issue(s)</a>"
        $priority = "Low"
        if($linkedIssuesCount -ge 5) {
            $priority = "Medium"
        }
        $obj = New-Object -TypeName psobject 
        $obj | Add-Member -MemberType NoteProperty -Name "Issue" -Value $bugLinkStr
        $obj | Add-Member -MemberType NoteProperty -Name "Description" -Value $summary
        $obj | Add-Member -MemberType NoteProperty -Name "Tests Blocked" -Value $linkedIssuesStr
        $obj | Add-Member -MemberType NoteProperty -Name "Notes" -Value ""
        $obj | Add-Member -MemberType NoteProperty -Name "Priority" -Value $priority
        $ret = $issuesBlockingTestsArr.Add($obj)
    }
    return $issuesBlockingTestsArr
}

#
# jiraRequestUtilities - gets total issues found via automation for a project
# @function 
# @param {string} qaProject - the qa project to search under
#
function getTotalIssuesFoundViaAutomationForProject($qaProject){
    $query = "`"QA Project`" = $qaProject and `"QA Identifiers`" = `"Found Via Automation`""
    $foundIssues = Get-JiraSearchResult -query $query -max $max -fields "key"
    return $foundIssues
}

#
# jiraRequestUtilities - gets total tests blocked due to issues
# @function 
# @param {string} [appComponent="OPS"] - the app component
#
function getTestsBlockedDueToIssues($appComponent="OPS") {
if($appComponent -eq "OPS") {
        $query = "project = `"`" and status = `"Automated Test`" and issueFunction in linkedIssuesOf(`"type = bug and resolution is empty`")"
    } elseif($appComponent -eq "") {
       $query = "project = `"`" and status = `"Automated Test`" and issueFunction in linkedIssuesOf(`"type = bug and resolution is empty`")"
     } else {
        $query = "project = `"`" and component = `"$appComponent`" and status = `"Automated Test`" and issueFunction in linkedIssuesOf(`"type = bug and resolution is empty`")"
    }
    $returnedIssues = Get-JiraSearchResult -query $query -fields "key"
    $returnedIssuesObj = @{
        "total" = $returnedIssues.total;
        "link" = "";
        "issues" = $returnedIssues
    }
    return $returnedIssuesObj
}

#
# jiraRequestUtilities - gets total issues in test plans
# @function 
# @param {string[]} [testPlans] - array of test plans
#
function getTotalIssuesInTestPlans($testPlans) {
    $testPlansStr = $testPlans -join ","
    $query = "type = Bug and resolution is empty and issueFunction in linkedIssuesOf(`"type = 'Test Case' and issueFunction in subtasksOf(\`"Type = 'Test Plan' and id in ($testPlansStr) \`")`")"
    $returnedIssues = Get-JiraSearchResult -query $query -fields "key"
    $returnedIssuesObj = @{
        "total" = $returnedIssues.total;
        "link" = "";
        "issues" = $returnedIssues
    }
    return $returnedIssuesObj
}

#
# jiraRequestUtilities - gets total reopened issues in test plans
# @function 
# @param {string[]} [testPlans] - array of test plans
#
function getTotalReopenedIssuesInTestPlans($testPlans) {
    $testPlansStr = $testPlans -join ","
    $query = "type = Bug and resolution is empty and status was in (reopened) and issueFunction in linkedIssuesOf(`"type = 'Test Case' and issueFunction in subtasksOf(\`"Type = 'Test Plan' and id in ($testPlansStr)\`")`")"
    $returnedIssues = Get-JiraSearchResult -query $query -fields "key"
    $returnedIssuesObj = @{
        "total" = $returnedIssues.total;
        "link" = "";
        "issues" = $returnedIssues
    }
    return $returnedIssuesObj
}

#
# jiraRequestUtilities - delete issue
# @function 
# @param {string} query - the query to use for finding issues to delete (WARNING: will delete any and all issues returned from the query
# @param {number} repeatLimit - the limit on the number of times the loop can be executed (how many sets of 50)
# @param {number} [max=50] - the maximum number of tests to return for performing one cycle of deletes
#
Function deleteIssues($query, $repeatLimit, $max=50) {
    Write-Host ""
    Write-Host "-------------- START jiraRequestUtilities.deleteIssues -----------------"
    Write-Host ""
    Write-Host "Deleting issues...."
    Write-Host "-------------------"
    $returnedIssues = Get-JiraSearchResult -query $query -max $max -fields "key"
    $total = $returnedIssues.total
    $repeats = 0
    while($repeats -lt $repeatLimit -and $total -gt 0) {
        foreach($issue in $returnedIssues.issues) {
            $key = $issue.key
            $url = 'issue/' + $key + '?deleteSubtasks=true'
            Write-Host "Deleting $key"
            transitionTestCase -testCase "$key" -status "" -transition "Delete"
        }
        $repeats++
        if($repeats -lt $repeatLimit) {
            Write-Host "Querying for next set of issues"
            sleep -Seconds 1
            $returnedIssues = Get-JiraSearchResult -query $query -max $max -fields "key"
            $total = $returnedIssues.total
        }
        
    }
    Write-Host "-------------------"
    Write-Host ""
    Write-Host "-------------- END jiraRequestUtilities.deleteIssues -----------------"
    Write-Host ""
}

#
# jiraRequestUtilities - deletes any open duplicates within a test plan
# @function 
# @param {string} testPlan - the key for the test plan to search in (Ex. PTEST-XXXX)
#
function deleteDuplicateTestCases($testPlan) {
    $query = "parent = $testPlan and status = 'Open'"
    $returnedIssues = Get-JiraSearchResult -query $query -max 500 
    $total = $returnedIssues.total
    $dupKeyArray = @()
    foreach($issue in $returnedIssues.issues) {
        $key = $issue.key
        $tcTemplate = $issue.fields.customfield_11883
        $dupQuery = "parent = $testPlan and status != 'Open' and 'TC Template' ~ $tcTemplate"
        $dupIssues = Get-JiraSearchResult -query $dupQuery -max 500
        if($dupIssues.total -ne 0) {
            Write-Host "$key is a duplicate"
            deleteIssues -query "key = $key" -repeatLimit 1 -max 1
        }
    }
}

#
# jiraRequestUtilities - Resets a TCT's remaining estimate equal to its original estimate
# @function 
# @param {string} testItem - TCT key name ("ETEST-4866", "PTEST-100")
# @param {string} estimate - estimate value
#
function resetJIRARemainingEstimate($testItem, $estimate) {
    Write-Host "Resetting TCT remaining estimate for $testItem"
    if($estimate) {
        $auth = $env:JIRA_AUTH
        $putURL =  "issue/$($testItem)/editmeta"
        $putBody = "{`"update`":{`"timetracking`":[{`"edit`":{`"remainingEstimate`": `"$($estimate)`"}}]}}"
        $jiraCon = .\runJIRARequest.ps1 -requestMethod "PUT" -requestUrl $putURL -requestBody $putBody
        #PUT doesn't return a parsable response body so unable to check response code
    } else {
        Write-Host "No estimate value found/given, aborting estimate reset"
    }
}


