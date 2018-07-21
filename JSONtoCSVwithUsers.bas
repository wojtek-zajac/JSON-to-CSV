Attribute VB_Name = "JSONtoCSVwithUsers"
Sub JSONtoCSVwithUsers()
Dim jsonData As String, JSON As Object, JsonString As String, rowData As String

jsonFile = Application.ActiveWorkbook.Path & ":users.json" 'Input json file
jsonfileNum = FreeFile
Open jsonFile For Input As #jsonfileNum
JsonString = Input(LOF(jsonfileNum), jsonfileNum)
Set JSON = ParseJson(JsonString)
Close #jsonfileNum

csvFile = Application.ActiveWorkbook.Path & ":users.csv" 'Output csv file
csvFileNum = FreeFile
Open csvFile For Output As #csvFileNum
Print #csvFileNum, "jobId,jobRefId,jobTitle,jobCreateDate,jobStatus,jobOpenPositions,candidatesRejected,candidatesWithdrawn,candidatesInReview,candidatesOffer,candidatesTotal,candidatesHired,locationCountry,locationState,locationCity"

rowData = ""
For Each Item In JSON
            jobIdField = formatField(Item("jobId"))
            jobRefIdField = formatField(Item("jobRefId"))
            jobTitleField = formatField(Item("jobTitle"))
            jobCreateDateField = formatField(Item("jobCreateDate"))
            jobStatusField = formatField(Item("jobStatus"))
            jobOpenPositionsField = formatField(Item("jobOpenPositions"))
            candidatesRejectedField = formatField(Item("candidatesRejected"))
            candidatesWithdrawnField = formatField(Item("candidatesWithdrawn"))
            candidatesInReviewField = formatField(Item("candidatesInReview"))
            candidatesInterviewField = formatField(Item("candidatesInterview"))
            candidatesOfferField = formatField(Item("candidatesOffer"))
            candidatesTotalField = formatField(Item("candidatesTotal"))
            candidatesHiredField = formatField(Item("candidatesTotal"))
            locationCountryField = formatField(Item("locationCountry"))
            locationStateField = formatField(Item("locationState"))
            locationCityField = formatField(Item("locationCity"), False)

    rowData = jobIdField + jobRefIdField + jobTitleField + jobCreateDateField + jobStatusField + jobOpenPositionsField + candidatesRejectedField + candidatesWithdrawnField + candidatesInReviewField + candidatesOfferField + candidatesTotalField + candidatesHiredField + locationCountryField + locationStateField + locationCityField
    
    Print #csvFileNum, rowData
Next
Close #csvFileNum
End Sub



Function formatField(val, Optional addComma As Boolean = True)
If (addComma = False) Then
formatField = Chr(34) + Trim(val) + Chr(34)
Else
formatField = Chr(34) + Trim(val) + Chr(34) + ","
End If
End Function
