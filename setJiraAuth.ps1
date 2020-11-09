param(	
	[Parameter(Mandatory=$false)][string]$username = "",
    [Parameter(Mandatory=$false)][string]$password = ""
)
[void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')

#
# setJiraAuth - sets JIRA_AUTH environment variable for JIRA http requests
# @function
# @author {CU} {LS}
#
function setAuth(){
    $ret = 1
    Write-Host "Setting JIRA_AUTH"
    try{
	    if($username -eq "" -or $password -eq "") {
			$cred = Get-Credential -Message "Please enter your Jira Credentials"
			$username = $cred.UserName
			$password = $cred.Password
		}
        $auth = [System.Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes("${username}:$([System.Runtime.InteropServices.Marshal]::PtrToStringAuto([System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($password)))"))
        [Environment]::SetEnvironmentVariable('JIRA_AUTH', $auth, "User")
      } catch {
        Write-Host "An error occurred while setting jira auth. See additional information below."
        Write-Host "$_.ErrorDetails.Message.message"
        $ret = 0
      } 
    Return $ret
        
}
setAuth