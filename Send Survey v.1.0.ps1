# ----------------------------------------------------------
# -- Send Survey v.1.0
# -- Last modified by Anders Asp, Innofactor - 2018-05-31
# ----------------------------------------------------------

#region Configuration
$DatabaseServer = 'STOMSSQL05\INST3'
$DatabaseName = 'OrchestratorWorkingDB'
$ServideDeskSurveyGUID = '841dc385-b43b-9447-012c-7e881ac28450'
$PortalURL = "http://servicedesk.swedishmatch.com"

#endregion

#region Functions

Function Insert-SQL {
    param(
        [parameter(Mandatory=$True)][string]$databaseServer,
        [parameter(Mandatory=$True)][string]$database,
        [parameter(Mandatory=$True)][string]$sqlCommand
    )

    $connectionString = "Data Source=$databaseServer; " +
                        "Integrated Security=SSPI; " +
                        "Initial Catalog=$database"
    $connection = New-Object System.Data.SqlClient.SqlConnection($connectionString)
    $command = New-Object System.Data.SqlClient.SqlCommand($sqlCommand,$connection)

    try{
        $connection.Open()
    } catch {
        Throw "Insert-SQL: Error connecting to SQL server and/or database!"
    }

    try {
        $command.ExecuteNonQuery()
    } catch {
        $connection.Close()
        Throw "Insert-SQL: Error executing INSERT statement: $($_)"
    }

    $connection.Close()

}

Function Query-SQL {
    param(
        [parameter(Mandatory=$True)][string]$databaseServer,
        [parameter(Mandatory=$True)][string]$database,
        [parameter(Mandatory=$True)][string]$sqlCommand
    )

    $connectionString = "Data Source=$databaseServer; " +
                        "Integrated Security=SSPI; " +
                        "Initial Catalog=$database"
    $connection = New-Object System.Data.SqlClient.SqlConnection($connectionString)
    $command = New-Object System.Data.SqlClient.SqlCommand($sqlCommand,$connection)
    try{
        $connection.Open()
    } catch {
        Throw "Query-SQL: Error connecting to SQL server and/or database!"
    }

    try{
        $adapter = New-Object System.Data.SqlClient.SqlDataAdapter $command
        $dataset = new-object System.Data.DataSet
        $adapter.Fill($dataset) | Out-Null
    } catch {
        $connection.Close()
        Throw "Query-SQL: Error executing SQL query: $($_)"
    }

    $connection.Close()

    $resultArray = @()
    foreach ($obj in $dataset.Tables.Rows){

        $Columns = $obj | Get-Member -MemberType Property | Select-Object Name
        $hash = @{}

        foreach ($column in $Columns) {
        
            $hash.Add($column.Name,$obj.$($column.Name))
        }
        $temp = New-Object -TypeName psobject -Property $hash
        $resultArray += $temp
    }

    Return $resultArray

}

Function Get-SCSMEmailAddress {
    param(
    [parameter(Mandatory=$True)]$UserObject
    )

    $UserPrefClass = Get-SCSMClass System.UserPreference$
    $EndPoint = Get-SCSMRelatedObject -SMObject $UserObject -Relationship $userPref| where{$_.DisplayName -like '*SMTP'}

    If($EndPoint){
        return $EndPoint.TargetAddress
    } else {
        return $null
    }
}

#endregion

#region Logics

# Import modules
Import-Module SMLets

# Get SCSM classes
$SRClass = Get-SCSMClass System.WorkItem.ServiceRequest$
$IRClass = Get-SCSMClass System.WorkItem.Incident$
$SurveyTemplateClass = Get-SCSMClass Cireson.Survey.Template$

# Get SCSM relationship classes
$AffectedCIRelClass = Get-SCSMRelationshipClass System.WorkItemAboutConfigItem$
$AffectedUserRelClass = Get-SCSMRelationshipClass System.WorkItemAffectedUser$

# Define output arrays
$OutPutEmailAddress = @()
$OutPutSurveyURL = @()
$OutPutSurveyMode = @()
$OutPutWIId = @()
$OutPutWITitle = @()

# -- Get LastSurveyDate
$Query = "Select LastResolvedDateCheck from SCSMSurvey_LastResolvedCheck"
try{
    $QueryResult = Query-SQL -databaseServer $DatabaseServer -database $DatabaseName -sqlCommand $Query
    [datetime]$LastResolvedDateCheck = $QueryResult.LastResolvedDateCheck
    [datetime]$NewResolvedDateCheck = (Get-Date).ToUniversalTime()
}
catch{
    Throw "Unable to query $DatabaseServer / $DatabaseName!"
}

If(!$LastResolvedDateCheck){Throw "Unable to retrieve LastResolvedDateCheck!"}

# -- Get Resolved/Completed WIs
try{
    $WIs = @()
    $WIs += Get-SCSMObject -Class $IRClass -Filter "ResolvedDate -gt $LastResolvedDateCheck"
    $WIs += Get-SCSMObject -Class $SRClass -Filter "CompletedDate -gt $LastResolvedDateCheck"
}
catch{
    Throw "Unable to retrieve data from SCSM!"
}

# -- Get all enabled Surveys
$AllSurveys = Get-SCSMObject -Class $SurveyTemplateClass -Filter "Enabled -eq $true"

# -- Loop through all WIs
Foreach($WI in $WIs){

    # - Get Affected User
    $AffectedUser = Get-SCSMRelatedObject -SMObject $WI -Relationship $AffectedUserRelClass

    If(!$AffectedUser){
        Write-Warning "$WI does not have an Affected User!"
        continue
    }

    # - Get Affected User emailaddress
    $AffectedUserEmail = Get-SCSMEmailAddress -UserObject $AffectedUser

    If(!$AffectedUserEmail){
        Write-Warning "User $AffectedUser does not have a valid email address specified within SCSM!"
        continue
    }

    # - Get Affected services
    $AffectedServices = Get-SCSMRelatedObject -SMObject $WI -Relationship $AffectedCIRelClass | Where-Object {$_.Classname -match 'System.Service|Microsoft.SystemCenter.BusinessService|.DA$|^Service_63a65c6fc67d43b2ac295399d63205eb$'}

    # - Determine which Survey that is applicable
    $SurveyMode = $null

    If(($AffectedServices|measure).Count -eq 1){
        
        $MatchingSurvey = $null

        # Find matching Service Survey
        Foreach($Survey in $AllSurveys){
            
            If($Survey.DisplayName -match "^$($AffectedServices.DisplayName) Survey"){
                $MatchingSurvey = $Survey
                break
            }
        }

        # Did we find a matching survey for the Affected Service on the WI?
        If($MatchingSurvey){
            $SurveyMode = $AffectedServices.DisplayName
        } else {
            $SurveyMode = "ServiceDesk"
            $MatchingSurvey = Get-SCSMObject -Id $ServideDeskSurveyGUID
        }

    } else {
        $SurveyMode = "ServiceDesk"
        $MatchingSurvey = Get-SCSMObject -Id $ServideDeskSurveyGUID
    }

    # - Determine if we should send survey or not

    $QueryLastSurvey = "Select Username, SurveyTemplateGuid, LastSurveyDate, TicketID From SCSMSurvey_SurveySendList where Username = '" + $AffectedUser.Username + "' And SurveyTemplateGuid = '" + $MatchingSurvey.Id + "'"
    $QueryResultLastSurvey = Query-SQL -databaseServer $DatabaseServer -database $DatabaseName -sqlCommand $QueryLastSurvey

    If($QueryResultLastSurvey.LastSurveyDate -ne $null){
        # Check if it's time to re-send the survey...

        # Do the matching survey have a recurrance set?
        If($MatchingSurvey.Notes -ne "" -and $MatchingSurvey.Notes -ne $null -and $MatchingSurvey.Notes -ne "0"){

            # Ensure that we have a valid value in $MatchingSurvey.Notes
            try{
                [int]$Recurrance = $MatchingSurvey.Notes
            }
            catch{
                Throw "Unable to convert recurrance ""$($MatchingSurvey.Notes)"" for $($MatchingSurvey.DisplayName) to INTEGER!"
            }

            # Is the LastSurveyDate for the User and Survey 
            If($QueryResultLastSurvey.LastSurveyDate -lt (Get-Date).AddDays(-$Recurrance)){
                # Add to "Send survey output"
                $OutPutEmailAddress += $AffectedUserEmail
                $OutPutSurveyURL += "$PortalURL/View/SurveyApp#/survey/create?templateId=$($MatchingSurvey.Id.Guid)&workItemId=$($WI.Get_ID().Guid)"
                $OutPutSurveyMode += $SurveyMode
                $OutPutWIId += $WI.Id
                $OutPutWITitle += $WI.Title

                Write-Output "Send survey: $($MatchingSurvey.DisplayName) to: $($AffectedUser.DisplayName) with e-mail: $($AffectedUserEmail)"

                # Update SCSMSurvey_SurveySendList with Survey/LastSurveyDate information for User
                try{
                    $QueryUpdateSurveySendList = "Update SCSMSurvey_SurveySendList Set LastSurveyDate = '" + $NewResolvedDateCheck + "', TicketID = '" + $WI.Id + "' Where Username = '" + $AffectedUser.UserName + "' And SurveyTemplateGuid = '" + $MatchingSurvey.Id + "'"
                    $Result = Insert-SQL -databaseServer $DatabaseServer -database $DatabaseName -sqlCommand $QueryUpdateSurveySendList
                }
                catch{
                    Throw "Unable to update SCSMSurvey_SurveySendList! $_"
                }
            }

        } else {
            # Recurrance is not specified / set to 0 - send the query on every single WI
            # Add to "Send survey output"
            $OutPutEmailAddress += $AffectedUserEmail
            $OutPutSurveyURL += "$PortalURL/View/SurveyApp#/survey/create?templateId=$($MatchingSurvey.Id.Guid)&workItemId=$($WI.Get_ID().Guid)"
            $OutPutSurveyMode += $SurveyMode
            $OutPutWIId += $WI.Id
            $OutPutWITitle += $WI.Title

            Write-Output "Send survey: $($MatchingSurvey.DisplayName) to: $($AffectedUser.DisplayName) with e-mail: $($AffectedUserEmail)"

            # Update SCSMSurvey_SurveySendList with Survey/LastSurveyDate information for User
            try{
                $QueryUpdateSurveySendList = "Update SCSMSurvey_SurveySendList Set LastSurveyDate = '" + $NewResolvedDateCheck + "', TicketID = '" + $WI.Id + "' Where Username = '" + $AffectedUser.UserName + "' And SurveyTemplateGuid = '" + $MatchingSurvey.Id + "'"
                $Result = Insert-SQL -databaseServer $DatabaseServer -database $DatabaseName -sqlCommand $QueryUpdateSurveySendList
            }
            catch{
                Throw "Unable to update SCSMSurvey_SurveySendList! $_"
            }
        }


    } else {
        # The Affected User never recieved this Survey, proceed and send it
        # Add to "Send survey output"
        $OutPutEmailAddress += $AffectedUserEmail
        $OutPutSurveyURL += "$PortalURL/View/SurveyApp#/survey/create?templateId=$($MatchingSurvey.Id.Guid)&workItemId=$($WI.Get_ID().Guid)"
        $OutPutSurveyMode += $SurveyMode
        $OutPutWIId += $WI.Id
        $OutPutWITitle += $WI.Title

        Write-Output "Send survey: $($MatchingSurvey.DisplayName) to: $($AffectedUser.DisplayName) with e-mail: $($AffectedUserEmail)"

        # Insert into SCSMSurvey_SurveySendList for User
        try{
            $QueryInsertSurveySendList = "Insert into SCSMSurvey_SurveySendList (Username, SurveyTemplateGuid, LastSurveyDate, TicketID) Values ('" + $AffectedUser.UserName + "','" + $MatchingSurvey.Id + "','" + $NewResolvedDateCheck + "','" + $WI.Id + "')"
            $Result = Insert-SQL -databaseServer $DatabaseServer -database $DatabaseName -sqlCommand $QueryInsertSurveySendList
        }
        catch{
            Throw "Unable to Insert Into SCSMSurvey_SurveySendList! $_"
        }
    }


}

# Update LastSurveyDate
$Query = "Update SCSMSurvey_LastResolvedCheck Set LastResolvedDateCheck = '$NewResolvedDateCheck'"
try{
    $Result = Insert-SQL -databaseServer $DatabaseServer -database $DatabaseName -sqlCommand $Query
}
catch{
    Throw "Unable to set new LastResolvedDateCheck in SQL!"
}




#endregion