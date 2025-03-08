<#
  .DESCRIPTION
  This cmdlet is designed to create a new Topdesk Support Ticket
  .PARAMETER details
  The details that we will be using for the ticket
  .PARAMETER emails
  The emails that will be part of this incident
  .PARAMETER sourceDir
  The directory that the process is running from, so it can save the emails to html files
#>
function New-AlertFactoryIncident{
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][hashtable]$details,
    [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][PSCustomObject]$emails,
    [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$sourceDir,
    [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$mailbox,
    [Parameter()][ValidateNotNullOrEmpty()][int]$count = 1
  )
  $body = @{}
  foreach($item in $details.GetEnumerator()){
    if($null -ne $item.Value){
      switch($item.key){
        "email"{
          $body.Add("caller_dynamicName",$item.Value)
          $body.Add("caller_email",$item.Value)
        }
        "location" {
          $body.Add("caller_branch_id",(Get-TopdeskBranches -clientReferenceNumber $item.Value | Select-Object -first 1).id)
        }
        "briefDescription" {
          $description = $item.Value
          if($count -gt 1){
            $description = "$($description) (Has Occured $($count) times)"
          }
          $body.Add("briefDescription",$description)
        }
        "callType" {
          $body.Add("callType_name",$item.Value)
          
        }
        "category" {
          $body.Add("category_name",$item.Value)
        }
        "subcategory" {
          $body.Add("subcategory_name",$item.Value)
        }
        "impact" {
          $body.Add("impact_name",$item.Value)
        }
        "urgency" {
          $body.Add("urgency_name",$item.Value)
        }
        "priorty" {
          $body.Add("priority_name",$item.Value)
        }
        "duration" {
          $body.Add("duration_name",$item.Value)
        }
        "operatorGroup" {
          $body.Add("operatorGroup_id",(Get-TopdeskOperatorGroup -query "groupName=='$($item.value)'" -fields id | Select-Object -first 1).id)
        }
        "operator" {
          $body.Add("operator_id",(Get-TopdeskOperator -query "dynamicName=='$($item.value)'" -fields id | Select-Object -first 1).id)
        }
        "status" {
          $body.Add("processingStatus_name",$item.Value)
        }
        "request" {
          $body.Add("request",$item.value)
        }
      }
    }
  }
  try{
    $incident = New-TopdeskIncident @body
    foreach($email in $emails)
    {
      $tempfolder = Join-Path -Path $sourceDir -ChildPath "Temp"
      if(-not (Test-Path $tempfolder)){
        New-Item -Path $tempfolder -Force -ItemType Directory
      }
      $filename = "$(Get-Random).html"
      $tempfile = Join-Path -Path $tempfolder -ChildPath $filename
      $email.body.content | Out-File $tempFile
      Add-TopdeskIncidentAttachment -id $incident.id -filepath $tempFile -filename $filename
      Remove-Item $tempfile -Force
      if($email.hasAttachments -and $rule.saveattachments){
        $attachments = Get-GraphMailAttachment -mailbox $mailbox -messageid $email.id
        foreach($attachment in $attachments){
          Add-TopdeskIncidentAttachment -id $incident.id -base64 $attachment.contentBytes -contenttype $attachment.contentType -filename $attachment.name
        }
      }      
    }
    return $incident
  }
  catch{
    throw "Unable to create incident. $($Error[0])"
  }
}