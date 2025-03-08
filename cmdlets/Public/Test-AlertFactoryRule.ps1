<#
  .DESCRIPTION
  This cmdlet is designed to test rules against emails that are in the mailbox for alert factory
  .PARAMETER rule
  The rule we are testing against
  .PARAMETER emails
  The emails that will be part of this incident
#>
function Test-AlertFactoryRule{
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][PSCustomObject]$rule,
    [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][PSCustomObject]$emails
  )
  foreach($obj in $rule.PSObject.Properties){
    if($null -ne $obj.value.type){
      switch($obj.name){
        "sender" {
          $lookupField = "from.emailAddress.Address"
        }
        "body" {
          $lookupField = "body.content"
        }
        "subject" {
          $lookupField = "subject"
        }
      }        
      switch($obj.value.type){
        "equals" {
          $command = "[System.Collections.Generic.List[PSObject]]`$emails = `$emails | Where-Object {`$_.$($lookupField) -eq `"$($obj.value.expression)`"}"
          
        }
        "contains" {
          $command = "[System.Collections.Generic.List[PSObject]]`$emails = `$emails | Where-Object {`$_.$($lookupField) -match `"$($obj.value.expression)`"}"
        }
        "regex" {
          $command = "[System.Collections.Generic.List[PSObject]]`$emails = `$emails | Where-Object {`$_.$($lookupField) -match `"$($obj.value.expression)`"}"
        }        
      }
      Invoke-Expression -Command $command
    }
  }
  return $emails
}