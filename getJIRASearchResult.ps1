param(	
	[Parameter(Mandatory=$false)][string]$query = "",
    [Parameter(Mandatory=$false)][string]$max = 5000,
    [Parameter(Mandatory=$false)][int]$start = 0
)

Function ConvertTo-SafeUri($uri) {
  Return [System.Uri]::EscapeDataString($uri)
}

Function Get-JiraSearchResult($query, $max=5000, $start=0) {
    $queryStr = "search?jql=$(ConvertTo-SafeUri $query)&maxResults=$max&startAt=$start"
    $responseVal = & ".\runJIRARequest.ps1" -requestMethod GET -requestUrl $queryStr
    Return $responseVal
}

Get-JiraSearchResult -query $query -max $max -start $start