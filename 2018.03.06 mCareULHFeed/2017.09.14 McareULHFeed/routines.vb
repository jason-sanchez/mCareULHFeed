Imports System
Imports System.IO
Imports System.Collections
Imports System.data.sqlclient
Public Module routines
    Dim connectionString As String
    Dim dictNVP As New Hashtable
    Dim sql As String
    Dim sql2 As String
    Dim myfile As StreamReader
    Dim dir As String
    Dim gblORIGINAL_PA_NUMBER As String
    Dim globalError As Boolean
    Dim bolDebug As Boolean
    Dim visitPaNumExists As Boolean
    Dim bolProcessInsurer As Boolean
    Dim theError As String
    Public Function parsePhone(ByVal data As String) As String
        'converts an hl7 phone number to format nnn.nnn.nnnn
        '7/28/2000 changed to handle 7 digits as a number without areacode
        If Len(data) = 10 Then
            parsePhone = Mid$(data, 1, 3) & "-" & Mid$(data, 4, 3) & "-" & Mid$(data, 7, 4)
        ElseIf Len(data) = 7 Then
            parsePhone = Mid$(data, 1, 3) & "-" & Mid$(data, 4, 4)
        Else
            parsePhone = ""
        End If
    End Function

    Public Function ConvertDate(ByVal datedata As String) As String
        'convert the hl7 date to a database date
        'hl7 in format: yyyymmdd or yyyymmddhhmm
        'returns now if the string is not in one
        'of the two formats.
        '
        Dim strYear As String
        Dim strMonth As String
        Dim strDay As String
        Dim strHour As String
        Dim strMinute As String

        If Len(Trim(datedata)) = 8 Then
            strYear = Mid$(datedata, 1, 4)
            strMonth = Mid$(datedata, 5, 2)
            strDay = Mid$(datedata, 7, 2)
            ConvertDate = strMonth & "/" & strDay & "/" & strYear

        ElseIf Len(Trim(datedata)) >= 12 Then
            strYear = Mid$(datedata, 1, 4)
            strMonth = Mid$(datedata, 5, 2)
            strDay = Mid$(datedata, 7, 2)
            strHour = Mid$(datedata, 9, 2)
            strMinute = Mid$(datedata, 11, 2)

            If strHour = "24" Then
                ConvertDate = strMonth & "/" & strDay & "/" & strYear
            Else
                ConvertDate = strMonth & "/" & strDay & "/" & strYear & " " & strHour & ":" & strMinute
            End If


        Else
            ConvertDate = DateTime.Now

        End If


    End Function
    
    Public Function ConvertSOS(ByVal data As String) As String
        'converts a sos number without delimiters to:
        ' sss-ss-ssss
        If Len(data) = 9 Then
            ConvertSOS = Mid$(data, 1, 3) & "-" & Mid$(data, 4, 2) & "-" & Mid$(data, 6, 4)
        ElseIf Len(data) = 11 Then
            ConvertSOS = data
        Else
            ConvertSOS = ""
        End If
    End Function

    
    
    
End Module
