<#
  .DESCRIPTION
  This cmdlet is designed to parse out data from emails for Alert Factory
#>
function Get-AlertFactoryTicketDetails{
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][PSCustomObject]$details,
    [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][PSCustomObject]$emails
  )
  # Generate blank hashtable for ticket
  $ticketDetails = @{}
  foreach($obj in $details.PSObject.Properties){
    if($null -ne $obj.value.type){
      if($null -ne $obj.value.attribute){
        switch($obj.value.attribute){
          "createdDateTime" {
            $content = $emails[0].createdDateTime
          }
          "receivedDateTime" {
            $content = $emails[0].receivedDateTime
          }
          "sentDateTime" {
            $content = $emails[0].sentDateTime
          }
          "subject" {
            $content = $emails[0].subject
          }
          "bodyPreview"{
            $content = $emails[0].bodyPreview
          }
          "body" {
            $content = $emails[0].body.content
          }
          "sender" {
            $content = $emails[0].sender.emailAddress.address
          }
          "from" {
            $content = $emails[0].from.emailAddress.address
          }
        }
      }
      else{
        $content = $obj.value.value
      }
      switch($obj.value.type){
        "source" {
          $ticketDetails.Add($obj.Name,$content) | Out-Null
        }
        "string" {
          $ticketDetails.Add($obj.Name,$obj.value.value) | Out-Null
        }
        "regex" {
          $match = [Regex]::Match($content,$obj.value.value)
          $val = $null          
          if($match.Success){
            $val = $match.Value
          }
          $ticketDetails.Add($obj.Name,$val) | Out-Null
        }
      }
    }
  }
  return $ticketDetails
}