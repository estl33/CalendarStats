[CmdletBinding()]
Param (
[Parameter(Position = 0, Mandatory = $False)]
[DateTime] $From = "08/01/2018",

[Parameter(Position = 0, Mandatory = $False)]
[DateTime] $To = "08/31/2018"
)

 
Try {
$outlook = New-Object -Com Outlook.Application -ErrorAction Stop
$mapi = $outlook.GetNameSpace("MAPI")
$roomList = $mapi.Folders
} Catch {
Write-Warning "Unable to create Outlook COM Object and connect to Outlook"
Exit
}

 
# Array to include all the objects for all meeting rooms
[Array] $roomsCol = @()
[Array] $meetingCol = @()

# Process all the meeting rooms one by one. Please note that mailboxes to which we have access in Outlook will also be processed
ForEach ($room in $roomList) {
$roomName = $room.Name

# Exclude our own mailbox from processing.
If (($roomName -eq "user1@yvr.ca") -or ($roomName -eq "user2@yvr.ca"))  {Continue}

# Print the name of the current meeting room being processed
Write-Host $roomName -ForegroundColor Green

# Check if the meeting room has any calendar items. If yes, then get all the items within the timespan specified at the beginning of the script
If (($room.Folders.Item("Calendar").Items) -ne $null) {
$calItems = $room.Folders.Item("Calendar").Items
$calItems.Sort("[Start]")
$calItems.IncludeRecurrences = $True
$dateRange = "[End] >= '{0}' AND [Start] <= '{1}'" -f $From.ToString("g"), $To.ToString("g")
$calItems = $calItems.Restrict($dateRange)
$totalItems = ($calItems | Measure-Object).Count

# Set some variables that will be used to save meeting information and process all meetings one by one
[Int] $count = 0
ForEach ($meeting in $calItems) {
Write-Progress -Activity "Processing $count / $totalItems"

[Int] $TotalAccpAttendees = $TotalReqAttendees = 0
For($i=1; $i -lt $meeting.Recipients.Count ;$i++){
if(($meeting.Recipients.Count -gt 0) -and ($meeting.Recipients -ne $null)){
if($meeting.Recipients.Item($i).MeetingResponseStatus -ne $null){
if (($meeting.Recipients.Item($i).MeetingResponseStatus -eq 3) -or ($meeting.Recipients.Item($i).MeetingResponseStatus -eq 0)) {$TotalAccpAttendees ++}
}
}
}

# Save the information gathered into an object and add the object to our object collection
$mObj = New-Object PSObject -Property @{
Room               = $roomName
Meeting            = $meeting.Subject
TotalReqAttendees  = ($meeting.RequiredAttendees.Split(";")).Count
TotalOptAttendees = If (($meeting.OptionalAttendees.toString()) -eq "") {0} Else{($meeting.OptionalAttendees.Split(";")).Count}
TotalAccpAttendees = $TotalAccpAttendees
StartTime          = $meeting.Start
EndTime            = $meeting.End
}
$meetingCol += $mObj
$count++
}

}
Else {Write-Host "calendar null" -ForegroundColor Green}
}

$meetingCol | Select Room, Meeting, TotalReqAttendees, TotalOptAttendees, TotalAccpAttendees, StartTime, EndTime | Sort Room | Export-Csv "C:\user\Desktop\FolderName\meeting_room_export$(Get-Date -f MMddyy_HHmm).csv" -NoTypeInformation
Write-Host "Finished exporting calendar to csv" -ForegroundColor Green 
