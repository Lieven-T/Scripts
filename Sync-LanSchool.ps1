Connect-Graph -Scopes @("User.ReadBasic.All","User.Read.All","Directory.Read.All","Group.Read.All")

$Classes = @()
$Users = @()
$Assignments = @()
$TenantID = "82812c36-6990-4cdc-a7f0-c481f0f68262"
$ClassId = 0
$UserId = 0
$EnrollmentId = 0
Get-MgGroup -Filter "startswith(displayname,'2122_')" -All | ? DisplayName -match "2122_(3BO1|3BW)" | % {
    $Class = $_
    $ClassId++
    $Classes += [PSCustomObject][Ordered]@{
        sourcedId = $ClassId
        status = ""
        dateLastModified = ""
        title = $_.DisplayName
        grades = ""
        courseSourcedId = "1_$($ClassId)"
        classCode = ""
        classType = "homeroom"
        location = "" 
        schoolSourcedId = 1
        termSourcedIds = "2122"
        subjects=""
        subjectCodes=""
        periods=""
    } 
    Get-MgGroupMember -GroupId $_.Id -Property @('givenName','surname','id','userPrincipalName') | ? displayname -notmatch "adweaver" | % {
        if ($_.AdditionalProperties.displayname -match "ADWeaver") { return }
        $UserId++
        $EnrollmentId++
        $Users += [PSCustomObject][Ordered]@{
            sourcedId = $UserId
            status = ""
            dateLastModified = ""
            enabledUser = "true"
            orgSourcedIds = "1"
            role = "student"
            username = $_.AdditionalProperties.userPrincipalName
            userIds = ""
            givenName = $_.AdditionalProperties.givenName ?? "Empty"
            familyName = $_.AdditionalProperties.surname ?? "Empty"
            middleName = ""
            identifier = ""
            email= $_.AdditionalProperties.userPrincipalName
            sms = ""
            phone = ""
            agentSourcedIds = ""
            grades = ""
            password = ""
        }   

        $Assignments += [PSCustomObject]@{
            sourcedId=$EnrollmentId
            status = ""
            dateLastModified = ""
            classSourcedId=$ClassId
            schoolSourcedId = 1
            userSourcedId = $UserId
            role = "student"
            primary = ""
            beginDate = ""
            endDate = ""
        }
    }

    Get-MgGroupOwner -GroupId $_.Id -Property @('givenName','surname','id','userPrincipalName') | % {
        if ($_.AdditionalProperties.displayname -match "ADWeaver") { return }
        $UserId++
        $EnrollmentId++
        $Users += [PSCustomObject][Ordered]@{
            sourcedId = $UserId
            status = ""
            dateLastModified = ""
            enabledUser = "true"
            orgSourcedIds = 1
            role = "teacher"
            username = $_.AdditionalProperties.userPrincipalName
            userIds = ""
            givenName = $_.AdditionalProperties.givenName ?? "Empty"
            familyName = $_.AdditionalProperties.surname ?? "Empty"
            middleName = ""
            identifier = ""
            email= $_.AdditionalProperties.userPrincipalName
            sms = ""
            phone = ""
            agentSourcedIds = ""
            grades = ""
            password = ""
            
        }        
        $Assignments += [PSCustomObject]@{
            sourcedId = $EnrollmentId
            status = ""
            dateLastModified = ""
            classSourcedId = $ClassId
            schoolSourcedId = 1
            userSourcedId = $UserId
            role = "teacher"
            primary = ""
            beginDate = ""
            endDate =""
        }
    }
}

$Assignments | Export-Csv c:\temp\assignments.csv -Delimiter ","
$Classes | Export-Csv c:\temp\classes.csv -Delimiter ","
$Users | Export-Csv c:\temp\users.csv -Delimiter ","