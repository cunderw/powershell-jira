param(	
	[Parameter(Mandatory=$false)][string]$requestMethod = "",
	[Parameter(Mandatory=$false)][string]$requestUrl = "",
	[Parameter(Mandatory=$false)][string]$requestBody = "",
    [Parameter(Mandatory=$false)][string]$apiBase = "api"
)	

Function Write-Host ($message,$nonewline,$backgroundcolor,$foregroundcolor) {
    $timestamp = Get-Date -Format "hh:mm:ss MM/dd/yy"
    $Message = "$timestamp [$env:computername] - $Message"
    $Message | Out-Host
}

$jiraConfig = Get-Content -Raw -Path "..\jiraConfig.json" | ConvertFrom-Json
[Net.ServicePointManager]::SecurityProtocol = "tls12, tls11, tls"
[void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
[void][System.Reflection.Assembly]::LoadWithPartialName("System.Web.Extensions")        
$jsonserial= New-Object -TypeName System.Web.Script.Serialization.JavaScriptSerializer 
$jsonserial.MaxJsonLength = [int]::MaxValue

# Invoke-JiraRequest - Runs an http request for JIRA
# @function
# @param {Object} $method - the request method to use
# ---------------------------------------------------
#		$method Options
#		1. GET
#		2. POST
#		3. REST
# ---------------------------------------------------
# @param {string} $request - the request to send to jira API service
# @param {string} [$body] - the body to send for the request (optional)
# @param {string} [$userAgent='Jira.psm1'] - the UserAgent for the request
# @returns {Object} $jiraResponse - the response object returned from the request

Function Invoke-JiraRequest($method, $request, $body=$Null, $userAgent='Jira.psm1') {
  $ServicePoint = [System.Net.ServicePointManager]::FindServicePoint(${env:JIRA_API_BASE})
  $ServicePoint.CloseConnectionGroup("")
  $uri = "${env:JIRA_API_BASE}${request}"
  If ($env:JIRA_API_BASE -eq $Null) {
      Write-Error "JIRA API Base has not been set, please run ``Set-JiraApiBase'"
  }
  If ($env:JIRA_AUTH -eq $Null) {
      Write-Error "No JIRA credentials have been set, please run 'setJiraAuth.ps1'"
  }
  Write-Debug "Calling $method $env:JIRA_API_BASE$request $env:JIRA_HTTP_PROXY"
  If ($body -eq $Null) {
      $responce = Invoke-WebRequest -Uri $uri -Headers @{"AUTHORIZATION"="Basic $env:JIRA_AUTH"} -Method $method -ContentType "application/json" -UserAgent $userAgent
	  $obj = $jsonserial.DeserializeObject($responce)
	  return $obj
  }
  else {
      $responce = Invoke-RestMethod -Uri $uri -Headers @{"AUTHORIZATION"="Basic $env:JIRA_AUTH"} -Method $method -Body $body -ContentType "application/json" -UserAgent $userAgent 
      # Invoke-RestMethod returns a powershell object for the response
      return $responce
  }
}

Function Set-JiraApiBase {    
    if($apiBase -eq "api") {
       $env:JIRA_API_BASE = $jiraConfig.bases.api
    } elseif ($apiBase -eq "agile"){
        $env:JIRA_API_BASE = $jiraConfig.bases.agile
    } elseif ($apiBase -eq "greenhopper"){
        $env:JIRA_API_BASE = $jiraConfig.bases.greenhopper
    }
}

Set-JiraApiBase

If ($env:JIRA_AUTH -eq $Null) {
  Invoke-Expression -Command ".\setJiraAuth.ps1"
}

if($requestBody -eq ""){
    $jiraResponse = Invoke-JiraRequest -method $requestMethod -request $requestUrl
} else {
    $jiraResponse = Invoke-JiraRequest -method $requestMethod -request $requestUrl -body $requestBody
}

$jiraResponse