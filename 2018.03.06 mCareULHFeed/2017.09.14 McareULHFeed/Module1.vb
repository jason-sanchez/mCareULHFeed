'20091117 - McareNewFeed VB 2008 version
'20100112 - fixed spelling error from PCPno to PCPNo
'20100218 - added ZMI procedures.
'20100222 - modified ZMI procedures to write to [03InsurerSupplement] db table
'20101001 - start code for multiple AL1 segments for Mcare

'20121010 - New version for McKesson testin as McareFeed_MC
'20121013 - changed to use NVP files.
'20121206 - add corpno for A01,A04,A05
'20130129 - McKesson Production
'20130508 - Start mods for Star feed
'20130620 - added calcStatus and xlateClass functions. MOd code to update patstatus with calcStatus value.
'20140202 - mods for wave3 testin on cscsysfeed5
'20140205 - use text log instead of event viewer
'20140215 - code for A31 records with multiple allergies.
'20140220 - new OEC calculation code in a separate procedure
'20140228 - changed logic to handle A31 records.
'20140303 - added severity, reaction and IDDate to 222AL1 and to processAL1 routine
'20140310 - corpno now in PID_2 = CORPNO. Also new status routine.
'20140311 - add call to process OEC routine
'20140317 - handle A11 like A03
'20140317 - new status and patstatus for A06, A07
'20140317 - added new calcPatStatus routine.
'20140318 - several bug fixes
'20140319 - don't change status on A13 and add new code for A13 in calcPatStatus, returning patient to Active (A)
'20140319 - added E patient class for patstatus and status.
'20140321 - use cerner version for extractCorpNoCerner if A31 else use the extractCorpNo version. Also only processError if not A31.
'20140402  - don't process OEC information for certain reord types.
'20140403 - fixed patStatus issue with datareader connection.
'20140405 - fixed datareader connection in OEC routine.
'20140724 - Added new ALI processing to table in PatientGlobal Database.
'20140817 - mods for W3 Production.
'20140915 - modified search criteria for processAL1
'20160916 - capture all AL1 data.
'20140917 - zero gblAL1Count before counting.
'20140918 - accept more than 9 AL1 segments in processAL1
'20140924 - remove length restriction on Patient Religion in patient routine.
'20141002 - fixed ARO.
'20141002 - fix patient religion error
'20141021 - write A34 Information to PatientGlobal database, table = A34Queue
'20141206 - added dictNVP.Clear() if no panum
'20140112 = added oldcorpNo to A34 Processing
'20150408 - remove processing of orphans.

'20150413 - VS2013 version
'20150418 - process orphans but don't log them.
'20150717 - remove orphan processing.

'20150821 - Mcare feed modified to support ULH on test system. 
'20151222 - start modes to run on KY2 Production interface


'20160113 - ProcessOEC: change patient type criteria for McareULH Star to OVT only
'20160203 - add Patient Type = "OMT" to processOEC
'20160203 - add B patient class (IP) to calcStarStatus
'20160421 - add B patient class (IA) to calcPatStatus

Imports System
Imports System.IO
Imports System.Collections
Imports System.data.sqlclient

'6/26/2006 - added A17 processing.
Module Module1
    Dim connectionString As String
    Dim dictNVP As New Hashtable
    'Dim dictNVP
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
    Dim gblInsCount As Integer
    Dim consultingArray() As String
    Dim strConsultingText As String
    Dim boolConsultingExists As Boolean
    Dim gblLogString As String = ""

    '4/26/2007 Error Handler Global Variables
    Dim functionError As Boolean
    Dim dbError As Boolean
    Dim continueProcessing As Boolean
    Dim orphanFound As Boolean


    Private fullinipath As String = Path.GetFullPath(Path.Combine(Directory.GetCurrentDirectory, "..\..\..\Configs\ULH\HL7Mapper.ini")) ' New test
    Public objIniFile As New INIFile(fullinipath) '20140817 - New Test
    'Public objIniFile As New INIFile("C:\KY2 Test Environment\HL7Mapper.ini") '20151222 'Local
    'Public objIniFile As New INIFile("C:\ULHTest\ULHMapper.ini") '20151222 'Test
    'Public objIniFile As New INIFile("c:\newfeeds\HL7Mapper.ini") '20151222 'Prod

    Private fullconinipath As String = Path.GetFullPath(Path.Combine(Directory.GetCurrentDirectory, "..\..\..\Configs\ULH\ConnProd.ini")) 'New Test
    Public conIniFile As New INIFile(fullconinipath) '20140805 New Test
    'Public conIniFile As New INIFile("C:\KY2 Test Environment\KY2ConnDev.ini") 'Local
    'Public conIniFile As New INIFile("C:\ULHTest\KY2ConnTest.ini") 'Test
    'Public conIniFile As New INIFile("C:\newfeeds\KY2ConnProd.ini") '20140805 Prod

    Dim strOutputDirectory As String = ""
    Dim strMapperFile As String = ""
    '20091118 - Public thefile As FileInfo added for exception processing
    Public thefile As FileInfo
    '20100211
    Dim gblZMICount As Integer = 0
    '20101001 - added global AL1count
    Dim gblAL1Count As Integer = 0
    '21011206
    Dim gblCorporateNumber As Integer = 0

    '20140205
    Dim strLogDirectory As String = ""

    Dim boolProcessOEC As Boolean = False
    Dim gblMrnum As String = ""


    Sub Main()

        bolDebug = True


        Try
            'declarations for split function
            Dim delimStr As String = "="
            Dim delimiter As Char() = delimStr.ToCharArray()
            'Dim theFile As FileInfo
            'declarations for stream reader
            Dim strLine As String
            Dim strServer As String = ""

            Dim dir As String = objIniFile.GetString("Settings", "directory", "(none)") & ":\"
            Dim parent As String = objIniFile.GetString("Settings", "parentDir", "(none)") & "\"

            strOutputDirectory = dir & parent & objIniFile.GetString("ULH_MCARE", "ULH_MCAREoutputdirectory", "(none)") '20150821
            strMapperFile = dir & parent & objIniFile.GetString("ULH_MCARE", "ULH_MCAREmapper", "(none)") '20150821
            'strServer = objIniFile.GetString("Settings", "server", "(none)")
            '20140205 - add logfile location
            strLogDirectory = dir & parent & objIniFile.GetString("Settings", "logs", "(none)")
            'setup directory
            '20121013 - changed to read NVP files
            Dim dirs As String() = Directory.GetFiles(strOutputDirectory, "NVP.*")


            'connectionString = "server=10.48.64.5\sqlexpress;database=mcareULHTest;uid=sysmax;pwd=Condor!"
            'connectionString = "server=10.48.242.249,1433;database=mcareULH;uid=sysmax;pwd=Condor!" '20151222

            connectionString = conIniFile.GetString("Strings", "MCAREULH", "(none)")


            Dim myConnection As New SqlConnection(connectionString)
            Dim updatecommand As New SqlCommand
            updatecommand.Connection = myConnection
            Dim objCommand As New SqlCommand
            objCommand.Connection = myConnection

            For Each dir In dirs
                thefile = New FileInfo(dir)
                If thefile.Extension <> ".$#$" Then
                    '1.set up the streamreader to get a file
                    myfile = File.OpenText(dir)
                    'and read the first line
                    strLine = myfile.ReadLine()
                    '20091123 - Catch a problem if the NVP file is messes up
                    Try
                        'Do While Not strLine Is Nothing
                        Do While Not myfile.EndOfStream
                            Dim myArray As String() = Nothing
                            strLine = myfile.ReadLine()
                            If strLine <> "" Then
                                myArray = strLine.Split(delimiter, 2)
                                'add array key and item to hashtable
                                Try
                                    dictNVP.Add(myArray(0), myArray(1))
                                Catch
                                End Try

                            End If


                        Loop
                    Catch ex As Exception
                        'make copy in the problems directory delete any previous ones with same name
                        Dim fi2 As FileInfo = New FileInfo(strOutputDirectory & "problems\" & thefile.Name)
                        fi2.Delete()
                        thefile.CopyTo(strOutputDirectory & "problems\" & thefile.Name)

                        gblLogString = gblLogString & "Dictionary Error" & " - " & thefile.Name & vbCrLf
                        gblLogString = gblLogString & ex.Message & vbCrLf
                        writeToLog(gblLogString, 1)
                        'get rid of the file so it doesn't mess up the next run.
                        myfile.Close()
                        If thefile.Exists Then
                            thefile.Delete()
                            Exit Sub
                        End If
                    End Try
                    '20091123 - Catch a problem if the NVP file is messes up

                    myfile.Close()
                    gblLogString = ""
                    gblORIGINAL_PA_NUMBER = ""

                    
                    gblMrnum = dictNVP.Item("mrnum")

                    '20180305 - Move files that do not have an observation datetime to the NO OBS Datetime Folder
                    If dictNVP.Item("TriggerEventID") = "Z47" Then
                        checkZ47(thefile, myfile)
                    Else

                        '20130523 - extract the corpNo from the fullpid data

                        '20140310 - corpno now in PID_2 = CORPNO
                        '20140321 - if A31 use cerner corp no extraction
                        If dictNVP.Item("TriggerEventID") = "A31" Then
                            gblCorporateNumber = extractCorpNoCerner(dictNVP.Item("fullpid"))
                        Else
                            gblCorporateNumber = extractCorpNo(dictNVP.Item("CORPNO"))
                        End If

                        '20141021 - Add A34 processing.
                        If dictNVP.Item("TriggerEventID") = "A34" Then 'added 20141021
                            Call ProcessA34(dictNVP)
                        End If

                        '20140402 only process OEC for certain record types.
                        boolProcessOEC = False
                        If dictNVP.Item("TriggerEventID") = "A01" Then boolProcessOEC = True
                        If dictNVP.Item("TriggerEventID") = "A02" Then boolProcessOEC = True
                        If dictNVP.Item("TriggerEventID") = "A04" Then boolProcessOEC = True
                        If dictNVP.Item("TriggerEventID") = "A07" Then boolProcessOEC = True
                        If dictNVP.Item("TriggerEventID") = "A08" Then boolProcessOEC = True

                        'If (dictNVP.Item("fullpid") <> "") Then
                        'gblCorporateNumber = extractCorpNo((dictNVP.Item("fullpid")))
                        'End If

                        If (dictNVP.Item("panum") <> "") Then
                            gblORIGINAL_PA_NUMBER = extractPanum(dictNVP.Item("panum"))
                            '===================================================================================================

                            '===================================================================================================

                            gblInsCount = 0
                            gblZMICount = 0
                            gblAL1Count = 0 ' 20101006 added initialization in file loop
                            '===================================================================================================
                            'call subdirectories here

                            '20140228 - if this is an A31 then just run the process Al1 routine otherwise handles as normal.
                            If dictNVP.Item("TriggerEventID") = "A31" Then
                                gblAL1Count = 0 '20140917
                                Call checkSegments(dictNVP)
                                Call processAL1(dictNVP)
                            Else
                                Call processError(dictNVP) ' 20140321 - proceserror only if not A31
                                If continueProcessing Then
                                    Call checkSegments(dictNVP)

                                    '20100218 - added checkZMI procedure.
                                    Call checkZMI(dictNVP)

                                    Call AddPatient(dictNVP)
                                    Call addVisit(dictNVP)
                                    Call addContact(dictNVP)
                                    '20130515 - don't process insurer if the plancode has double quotes or is blank
                                    If dictNVP.Item("iplancode") <> """""" And dictNVP.Item("iplancode") <> "" Then
                                        Call AddInsurer(dictNVP)
                                    End If

                                    '20140311 - add call to process OEC routine
                                    '20140402 don't process OEC information if boolProcessOEC is false.
                                    If boolProcessOEC Then
                                        Call processOEC(dictNVP)
                                    End If

                                End If 'If continueProcessing
                            End If ' if A31
                            '===================================================================================================
                            dictNVP.Clear()
                            If functionError Then

                                gblLogString = thefile.Name & vbCrLf & gblLogString
                                writeTolog(gblLogString, 2)
                                Dim fi2 As FileInfo = New FileInfo(strOutputDirectory & "backup\" & thefile.Name)
                                fi2.Delete()
                                thefile.CopyTo(strOutputDirectory & "backup\" & thefile.Name)
                                thefile.Delete()



                            ElseIf dbError Then

                                gblLogString = thefile.Name & vbCrLf & gblLogString

                                writeTolog(gblLogString, 2)
                                Dim fi2 As FileInfo = New FileInfo(strOutputDirectory & "reprocess\" & thefile.Name)
                                fi2.Delete()
                                thefile.CopyTo(strOutputDirectory & "reprocess\" & thefile.Name)
                                thefile.Delete()

                            ElseIf orphanFound Then
                                '20150408 - remove processing of orphans.
                                '20150418 - process orphans but don't log
                                'gblLogString = thefile.Name & vbCrLf & gblLogString

                                'writeTolog(gblLogString, 3)


                                Dim fi2 As FileInfo = New FileInfo(strOutputDirectory & "orphans\" & thefile.Name)
                                fi2.Delete()
                                thefile.CopyTo(strOutputDirectory & "orphans\" & thefile.Name)
                                thefile.Delete()

                            ElseIf globalError Then
                                gblLogString = "Global Error - " & thefile.Name & vbCrLf & gblLogString
                                writeTolog(gblLogString, 2)
                                gblLogString = ""

                                Dim fi2 As FileInfo = New FileInfo(strOutputDirectory & "backup\" & thefile.Name)
                                fi2.Delete()
                                thefile.CopyTo(strOutputDirectory & "backup\" & thefile.Name)
                                thefile.Delete()

                            Else
                                '20121210 - make a backup copy for review during McKesson Testing
                                'Dim fi2 As FileInfo = New FileInfo(strOutputDirectory & "mckesson\" & thefile.Name)
                                'fi2.Delete()
                                'thefile.CopyTo(strOutputDirectory & "mckesson\" & thefile.Name)
                                thefile.Delete()

                            End If

                        Else
                            '20080110: added code to delete NVP files with no panum
                            dictNVP.Clear() '20141206 - added  if no panum
                            thefile.Delete()
                        End If 'If (dictNVP.Item("panum") <> ""

                    End If 'If dictNVP.Item("TriggerEventID") = "Z47"

                End If 'If theFile.Extension <> ".$#$"
            Next
            'End main Processing -

        Catch ex As Exception
            globalError = True

            gblLogString = gblLogString & "Main Routine Error" & vbCrLf
            gblLogString = gblLogString & ex.Message & vbCrLf

            writeToLog(gblLogString, 1)
            '20130409 - issue with corpo > int. Processing hangs if record not deleted
            Dim fi2 As FileInfo = New FileInfo(strOutputDirectory & "backup\" & thefile.Name)
            fi2.Delete()
            thefile.CopyTo(strOutputDirectory & "backup\" & thefile.Name)
            thefile.Delete()

            Exit Sub
        End Try
    End Sub

    Public Sub AddPatient(ByVal dictNVP As Hashtable)
        '20130122 - add corpNo to patient table
        '20130513 - modifications for star feed
        Dim myConnection As New SqlConnection(connectionString)
        Dim objCommand As New SqlCommand
        Dim updatecommand As New SqlCommand
        updatecommand.Connection = myConnection

        objCommand.Connection = myConnection
        Dim dataReader As SqlDataReader
        Dim mrExists As Boolean
        Dim sql As String
        Dim addit As Boolean
        Dim updateit As Boolean

        addit = False
        updateit = False

        Try



            If dictNVP.Item("Event Type Code") = "A01" Then addit = True
            If dictNVP.Item("Event Type Code") = "A04" Then addit = True
            If dictNVP.Item("Event Type Code") = "A05" Then addit = True
            If dictNVP.Item("Event Type Code") = "A08" Then addit = True
            '20100114 - add non existing mrnum if A18
            If dictNVP.Item("Event Type Code") = "A18" Then addit = True

            '20090818 deal with phone numbers=======================================================
            Dim primaryPhone As String = ""
            Dim altPhone As String = ""
            Dim businessPhone As String = ""
            
            '20130813
            primaryPhone = Replace(dictNVP.Item("01patient.patPhone STAR"), "'", "''")
            'End If
            businessPhone = Replace(dictNVP.Item("Patient Business Phone STAR"), "'", "''")
            '20090818==================================================================================


            If (addit) Then
                'test to see if the mr number exists
                sql = "select mrnum from [01patient] where mrnum = " & dictNVP.Item("mrnum")
                objCommand.CommandText = sql
                myConnection.Open()
                dataReader = objCommand.ExecuteReader()


                If dataReader.HasRows Then
                    mrExists = True
                Else
                    mrExists = False
                End If
                myConnection.Close()
                dataReader.Close()

                'if the mr number does not exist add the record. Also there must be a visit record
                If (Not mrExists) Then 'process an add record

                    '20130122 - add corpNo on insert
                    sql = "INSERT [01patient] (mrnum, corpNo, patlast, patfirst, patSex, patRace, patmi, pataddr1, employer, "

                    '20090803 - add files for AltPhone and BusinessPhone after employer
                    sql = sql & "patphone, altphone, businessphone, "

                    If Len(dictNVP.Item("Patient Religion")) > 0 Then
                        sql = sql & "pataddr2, patcity, patstate, patzip, dob, patss, patCounty, religionID, added) "
                    Else
                        sql = sql & "pataddr2, patcity, patstate, patzip, dob, patss, patCounty, added) "
                    End If

                    sql = sql & "VALUES ("
                    sql = sql & CLng(dictNVP.Item("mrnum")) & ", "
                    sql = sql & gblCorporateNumber & ", "
                    sql = sql & "'" & Replace(dictNVP.Item("01patient.patlast"), "'", "''") & "', "
                    sql = sql & "'" & Replace(dictNVP.Item("01patient.patfirst"), "'", "''") & "', "
                    sql = sql & "'" & Replace(dictNVP.Item("Patient Sex"), "'", "''") & "', "
                    sql = sql & "'" & Replace(dictNVP.Item("Patient Race"), "'", "''") & "', "
                    sql = sql & "'" & Left$(dictNVP.Item("01patient.patmi"), 1) & "', "
                    sql = sql & "'" & Replace(dictNVP.Item("01patient.pataddr1"), "'", "''") & "', "
                    sql = sql & "'" & Replace(dictNVP.Item("PatientEmployer"), "'", "''") & "', "

                    '20090818 add processing for altphone and businessPhone
                    If primaryPhone <> "" Then
                        sql = sql & "'" & primaryPhone & "', "
                    Else
                        sql = sql & "NULL, "
                    End If
                    If altPhone <> "" Then
                        sql = sql & "'" & altPhone & "', "
                    Else
                        sql = sql & "NULL, "
                    End If
                    If businessPhone <> "" Then
                        sql = sql & "'" & businessPhone & "', "
                    Else
                        sql = sql & "NULL, "
                    End If

                    sql = sql & "'" & Replace(dictNVP.Item("01patient.pataddr2"), "'", "''") & "', "
                    sql = sql & "'" & Replace(dictNVP.Item("01patient.patcity"), "'", "''") & "', "
                    '20090612 don't process a state longer than two characters
                    If Len(dictNVP.Item("01patient.patstate")) <> 2 Then
                        sql = sql & "'KY', "
                    Else
                        sql = sql & "'" & dictNVP.Item("01patient.patstate") & "', "
                    End If

                    'sql = sql & "'" & dictNVP.Item("01patient.patstate") & "', "

                    '20090612 if no numeric zip don't enter
                    If IsNumeric(dictNVP.Item("01patient.patzip")) Then
                        sql = sql & "'" & dictNVP.Item("01patient.patzip") & "', "
                    Else
                        sql = sql & "'00000', "
                    End If
                    sql = sql & "'" & ConvertDate(dictNVP.Item("01patient.DOB")) & "', "
                    sql = sql & "'" & ConvertSOS(dictNVP.Item("01patient.patSS")) & "', "

                    sql = sql & "'" & Replace(dictNVP.Item("Patient County/Parish Code"), "'", "''") & "', "

                    If Len(dictNVP.Item("Patient Religion")) > 0 Then
                        sql = sql & "'" & Replace(dictNVP.Item("Patient Religion"), "'", "''") & "', "

                    End If

                    sql = sql & "'" & DateTime.Now & "') "
                    updatecommand.CommandText = sql
                    myConnection.Open()
                    updatecommand.ExecuteNonQuery()
                    myConnection.Close()

                Else
                    'update existing patient record
                    If Len(dictNVP.Item("01patient.patlast")) > 1 Then

                        sql = "UPDATE [01patient] "
                        sql = sql & "SET patLast = '" & Replace(dictNVP.Item("01patient.patlast"), "'", "''") & "' "

                        sql = sql & ",corpNo = " & gblCorporateNumber & " "

                        'sql = sql & ",patFirst = '" & Replace(dictNVP.Item("01patient.patfirst"), "'", "''") & "'"
                        If Len(dictNVP.Item("01patient.patfirst")) > 0 Then
                            sql = sql & STARupdateString("patFirst", dictNVP.Item("01patient.patfirst"))
                        End If

                        If Len(dictNVP.Item("01patient.patmi")) > 0 And dictNVP.Item("01patient.patmi") <> """""" Then
                            sql = sql & STARupdateString("patMI", Replace(Left$(dictNVP.Item("01patient.patmi"), 1), "'", "''"))
                        Else
                            sql = sql & STARupdateString("patMI", dictNVP.Item("01patient.patmi"))
                        End If

                        sql = sql & ",updated = '" & DateTime.Now & "' "
                        ' don't process address if pataddr1 is blank
                        If dictNVP.Item("01patient.pataddr1") <> "" Then
                            If Len(dictNVP.Item("01patient.pataddr1")) > 0 Then
                                sql = sql & STARupdateString("patAddr1", dictNVP.Item("01patient.pataddr1"))
                            End If

                            'sql = sql & ",patAddr2 = '" & Replace(dictNVP.Item("01patient.pataddr2"), "'", "''") & "' "
                            If Len(dictNVP.Item("01patient.pataddr2")) > 0 Then
                                sql = sql & STARupdateString("patAddr2", dictNVP.Item("01patient.pataddr2"))
                            End If

                            If Len(dictNVP.Item("01patient.patcity")) > 0 Then
                                sql = sql & STARupdateString("patCity", dictNVP.Item("01patient.patcity"))
                            End If
                        End If
                        '20090612 don't process a state longer than two characters
                        If Len(dictNVP.Item("01patient.patstate")) = 2 Then

                            If Len(dictNVP.Item("01patient.patstate")) > 0 Then
                                sql = sql & STARupdateString("patState", dictNVP.Item("01patient.patstate"))
                            End If
                        End If
                        '20090612 if no numeric zip don't enter
                        If IsNumeric(dictNVP.Item("01patient.patzip")) Then

                            If Len(dictNVP.Item("01patient.patZip")) > 0 Then
                                sql = sql & STARupdateString("patZip", dictNVP.Item("01patient.patZip"))
                            End If
                        End If

                        sql = sql & ",patSS = '" & Replace(ConvertSOS(dictNVP.Item("01patient.patSS")), "'", "''") & "' "
                        sql = sql & ",patSex = '" & Replace(dictNVP.Item("Patient Sex"), "'", "''") & "' "
                        sql = sql & ",patRace = '" & Replace(dictNVP.Item("race code STAR"), "'", "''") & "' "
                        sql = sql & ",raceDesc = '" & Replace(dictNVP.Item("race description STAR"), "'", "''") & "' "
                        sql = sql & ",employer = '" & Replace(dictNVP.Item("PatientEmployer"), "'", "''") & "' "
                        sql = sql & ",patCounty = '" & Replace(dictNVP.Item("Patient County/Parish Code"), "'", "''") & "' "
                        If Len(dictNVP.Item("Patient Religion")) > 0 Then
                            sql = sql & ",religionID = '" & Replace(dictNVP.Item("Patient Religion"), "'", "''") & "' "
                        End If


                        '20090818==============================================================
                        If primaryPhone <> "" Then
                            sql = sql & ",patPhone = '" & primaryPhone & "' "
                        End If

                        If altPhone <> "" Then
                            sql = sql & ",altPhone = '" & altPhone & "' "
                        End If

                        If businessPhone <> "" Then
                            sql = sql & ",businessPhone = '" & businessPhone & "' "
                        End If
                        '=====================================================================

                        If Not ConvertDate(dictNVP.Item("01patient.DOB")) = "" Then

                            sql = sql & ",DOB = '" & ConvertDate(dictNVP.Item("01patient.DOB")) & "' "
                        Else
                            sql = sql & ",DOB = null "
                        End If

                        sql = sql & "Where MRNum = " & dictNVP.Item("mrnum")
                        updatecommand.CommandText = sql
                        myConnection.Open()
                        updatecommand.ExecuteNonQuery()
                        myConnection.Close()

                        '8/6/2003
                    End If ' If Len(dictNVP.Item("01patient.patlast")) > 1

                    End If 'mrExists
            End If 'If (addit) Then

        Catch ex As Exception
            functionError = True
            gblLogString = gblLogString & "Patient Process Error" & vbCrLf
            gblLogString = gblLogString & ex.Message & vbCrLf
            'LogFile.Close()
            Exit Sub
        End Try
    End Sub
    Public Sub AddInsurer(ByVal dictNVP As Hashtable)
        '20140603 - code to use star plancode in the iplancode field
        '20140904 - return to integer fclass
        Dim myConnection As New SqlConnection(connectionString)
        Dim objCommand As New SqlCommand
        Dim updatecommand As New SqlCommand
        Dim dataReader As SqlDataReader
        Dim addit As Boolean
        Dim updateit As Boolean
        Dim tempstr As String
        Dim tempstr2 As String
        Dim sql As String
        Dim intFclass As Integer = 0 '20140904
        Dim iPlanCodeExists As Boolean
        Dim bolProcessThis As Boolean
        Dim strAuthServices As String
        Dim paNumExists As Boolean
        Dim i As Integer
        updatecommand.Connection = myConnection
        objCommand.Connection = myConnection
        paNumExists = False
        strAuthServices = ""
        bolProcessThis = False

        '20130513
        Dim star_plancode As String = ""
        Dim strFclass As String = ""

        Try
            If dictNVP.Item("Event Type Code") = "A01" Then bolProcessThis = True
            If dictNVP.Item("Event Type Code") = "A04" Then bolProcessThis = True
            If dictNVP.Item("Event Type Code") = "A08" Then bolProcessThis = True

            If (bolProcessThis) Then
                'See if any records exist in the insurer table for this panum
                sql = "select panum from [03insurer] where panum = '" & gblORIGINAL_PA_NUMBER & "'"
                objCommand.CommandText = sql
                myConnection.Open()
                dataReader = objCommand.ExecuteReader()
                If dataReader.HasRows Then
                    paNumExists = True
                Else
                    paNumExists = False
                End If
                myConnection.Close()
                dataReader.Close()

                intFclass = 0
                sql = "select id from [104finclass] where finclass = '" & dictNVP.Item("insurer.fClass") & "' and inactive = 0" '20140904
                objCommand.CommandText = sql
                myConnection.Open()
                dataReader = objCommand.ExecuteReader()
                While dataReader.Read()
                    intFclass = dataReader.GetInt32(0)
                End While
                myConnection.Close()
                dataReader.Close()

                addit = False
                updateit = False
                'strFclass = dictNVP.Item("insurer.fClass") '20140904

                If dictNVP.Item("Event Type Code") = "A01" Then addit = True
                If dictNVP.Item("Event Type Code") = "A04" Then addit = True


                If dictNVP.Item("Event Type Code") = "A08" Then updateit = True

                If ((addit) And (visitPaNumExists) And Not (paNumExists)) Then
                    Call insertInsurer(dictNVP)
                End If
                tempstr = ""
                If ((updateit) And (visitPaNumExists)) Then
                    i = 0
                    For i = 1 To gblInsCount
                        If i = 1 Then
                            tempstr = ""
                        End If
                        If i > 1 Then
                            tempstr = "_000" & i
                        End If
                        '20130513 - Added STAR_Plancode to [03insurer] table
                        star_plancode = Trim(Replace(dictNVP("iplancode2" & tempstr), "'", "''")) & Trim(Replace(dictNVP("iplancode" & tempstr), "'", "''"))

                        If Len(dictNVP.Item("iplancode" & tempstr)) >= 3 Then
                            iPlanCodeExists = False
                            sql = "SELECT panum FROM [03Insurer] where panum = '" & gblORIGINAL_PA_NUMBER & "' "
                            'sql = sql & "AND star_plancode  = '" & star_plancode & "'" '20140603
                            sql = sql & "AND iPlanCode  = '" & star_plancode & "'" '20140603

                            objCommand.CommandText = sql
                            myConnection.Open()
                            dataReader = objCommand.ExecuteReader()

                            If dataReader.HasRows Then
                                iPlanCodeExists = True
                            Else
                                iPlanCodeExists = False
                            End If
                            myConnection.Close()
                            dataReader.Close()

                            If iPlanCodeExists Then

                                sql = "UPDATE [03Insurer] "
                                sql = sql & "SET updated = '" & DateTime.Now & "'"
                                If Len(dictNVP.Item("CompanyName" & tempstr)) > 0 Then
                                    sql = sql & " ,coName = '" & Replace(dictNVP.Item("CompanyName" & tempstr), "'", "''") & "'"
                                End If

                                If Len(dictNVP.Item("group" & tempstr)) > 0 Then
                                    sql = sql & ", theGroup = '" & Replace(dictNVP.Item("group" & tempstr), "'", "''") & "'"
                                End If

                                If Len(dictNVP.Item("PolicyNumber" & tempstr)) > 0 Then
                                    sql = sql & ", policyNum = '" & Replace(dictNVP.Item("PolicyNumber" & tempstr), "'", "''") & "'"
                                End If

                                '20170512 parse out primary AuthNum from new IN1_14
                                'Dim autharray() As String = dictNVP.Item("AuthNum" & tempstr).split("~")
                                'Dim Authcode As String = Replace(autharray(0), "^", "")

                                'If Len(Authcode) > 0 Then
                                'sql = sql & ", authnum1 = '" & Replace(Authcode, "'", "''") & "'"
                                'End If
                                If Len(dictNVP.Item("AuthNum" & tempstr)) > 0 Then
                                    sql = sql & ", authnum1 = '" & Replace(dictNVP.Item("AuthNum" & tempstr), "'", "''") & "'"
                                End If

                                tempstr2 = ""
                                tempstr2 = dictNVP.Item("Insured First Name" & tempstr) & _
                                        " " & dictNVP.Item("Insured Middle Name" & tempstr) & _
                                        " " & dictNVP("Insured Last Name" & tempstr)

                                '5/29/2003 - check for a string length > 3 vice zero
                                If Len(tempstr2) > 3 Then
                                    sql = sql & ", subname = '" & Replace(tempstr2, "'", "''") & "'"
                                End If

                                If Len(dictNVP.Item("PolicyIssueDate" & tempstr)) = 8 Then
                                    sql = sql & ", PIssue = '" & ConvertDate(dictNVP.Item("PolicyIssueDate" & tempstr)) & "'"
                                End If

                                If dictNVP("COBPriority" & tempstr) = "1" Then
                                    sql = sql & ", aprimary = 1"
                                    sql = sql & ", fclass = " & intFclass & " " '20140904
                                Else
                                    sql = sql & ", aprimary = 0"
                                    sql = sql & ", fclass = 0"
                                End If

                                '5/22/2003 - added insurer address and phone information
                                If Len(dictNVP.Item("addr1" & tempstr)) > 0 Then
                                    sql = sql & ", addr1 = '" & Replace(dictNVP.Item("addr1" & tempstr), "'", "''") & "'"
                                End If
                                If Len(dictNVP.Item("addr2" & tempstr)) > 0 Then
                                    sql = sql & ", addr2 = '" & Replace(dictNVP.Item("addr2" & tempstr), "'", "''") & "'"
                                End If
                                If Len(dictNVP.Item("city" & tempstr)) > 0 Then
                                    sql = sql & ", city = '" & Replace(dictNVP.Item("city" & tempstr), "'", "''") & "'"
                                End If
                                If Len(dictNVP.Item("state" & tempstr)) > 0 Then
                                    sql = sql & ", state = '" & Replace(dictNVP.Item("state" & tempstr), "'", "''") & "'"
                                End If
                                If Len(dictNVP.Item("zip" & tempstr)) > 0 Then
                                    sql = sql & ", zip = '" & Replace(dictNVP.Item("zip" & tempstr), "'", "''") & "'"
                                End If
                                If Len(dictNVP.Item("phone" & tempstr)) > 0 Then
                                    sql = sql & ", phone = '" & Replace(Left(dictNVP.Item("phone" & tempstr), 15), "'", "''") & "'"
                                End If

                                '5/22/2003 - added auth services
                                '5/29/2003 - check for a string length > 1 vice zero
                                strAuthServices = dictNVP.Item("authservices1" & tempstr) & " " & dictNVP.Item("authservices2" & tempstr)
                                If Len(strAuthServices) > 1 Then
                                    sql = sql & ", preCertAuth = '" & Replace(strAuthServices, "'", "''") & "'"
                                End If

                                '5/22/2003 - add subscriber code (subscriber employer)
                                If Len(dictNVP.Item("SubscriberEmployer" & tempstr)) > 0 Then
                                    sql = sql & ", subEmployer = '" & Replace(dictNVP.Item("SubscriberEmployer" & tempstr), "'", "''") & "'"
                                End If

                                '5/22/2003 - add precert contact code
                                If Len(dictNVP.Item("precertcontact" & tempstr)) > 0 Then
                                    sql = sql & ", precertcontact = '" & Replace(dictNVP.Item("precertcontact" & tempstr), "'", "''") & "'"
                                End If

                                '10/19/2005===============================================================================
                                If Len(dictNVP("Insured DOB" & tempstr)) > 1 Then
                                    sql = sql & ", InsuredDOB = '" & ConvertDate(dictNVP.Item("Insured DOB" & tempstr)) & "'"
                                End If
                                If dictNVP("Insured Sex" & tempstr) <> "" Then
                                    sql = sql & ", InsuredSex = '" & Replace(dictNVP.Item("Insured Sex" & tempstr), "'", "''") & "'"
                                End If
                                '10/19/2005 end===========================================================================

                                sql = sql & ", reqCert = 1"
                                '20170804 - Keep track of insurance changes
                                sql = sql & ", lastupdatedStar = '" & DateTime.Now & "' "

                                sql = sql & " Where panum = '" & gblORIGINAL_PA_NUMBER & "'"
                                'sql = sql & " AND star_planCode = '" & star_plancode & "'" '20140603
                                sql = sql & " AND iPlancode = '" & star_plancode & "'" '20140603

                                'LogFile.WriteLine(sql)
                                updatecommand.CommandText = sql
                                myConnection.Open()
                                updatecommand.ExecuteNonQuery()
                                myConnection.Close()
                                gblLogString = gblLogString & "update " & dictNVP.Item("iplancode" & tempstr) & vbCrLf

                                ProcessIN1_14(dictNVP, tempstr)
                            Else                'iPlancode does not exist

                                'Call insertInsurer(dictNVP)
                                Call insertOneInsurer(dictNVP, tempstr)
                            End If              'If iPlanCodeExists Then

                        End If                  'Len(dictNVP.Item("iplancode" & tempstr)) >= 3
                    Next                        'gblInsCount
                End If                          'If ((updateit) And (visitPaNumExists))
            End If                              'If (bolProcessThis)
            'LogFile.WriteLine("Insurer Process Complete")
        Catch ex As Exception
            functionError = True
            gblLogString = gblLogString & "Insurer Process Error" & vbCrLf
            gblLogString = gblLogString & ex.Message & vbCrLf
            'LogFile.Close()
            Exit Sub
        End Try
    End Sub

    Public Sub addVisit(ByVal dictNVP As Hashtable)
        '5/31/2006 - added code to handle primary plan
        '20121206 - add corpNo
        Dim sql As String
        Dim addit As Boolean
        Dim updateit As Boolean
        Dim visitStatus As String
        Dim strCurrman As String
        'Dim intCurrmanNo As Integer
        Dim strRoomBed As String
        Dim boolAccidentExists As Boolean
        boolAccidentExists = False
        Dim paNumExists As Boolean
        paNumExists = False

        Dim tempCOBPriority As String = ""
        Dim i As Integer
        Dim tempstr As String = ""

        Dim myConnection As New SqlConnection(connectionString)
        Dim objCommand As New SqlCommand
        Dim updatecommand As New SqlCommand
        updatecommand.Connection = myConnection

        objCommand.Connection = myConnection
        Dim dataReader As SqlDataReader
        Dim j As Integer
        addit = False
        updateit = False
        Dim strEventTypeCode As String = dictNVP.Item("Event Type Code")

        boolConsultingExists = False

        'If Len(dictNVP.Item("Consulting Physician")) > 5 Then
        'boolConsultingExists = True ' will use this later when updating database
        'strConsultingText = dictNVP.Item("Consulting Physician")
        'consultingArray = Split(strConsultingText, "~")
        'End If

        Try
            'see if panum exixts
            sql = "select panum from [01visit] where panum = '" & gblORIGINAL_PA_NUMBER & "'"
            objCommand.CommandText = sql
            myConnection.Open()
            dataReader = objCommand.ExecuteReader()

            If dataReader.HasRows Then
                paNumExists = True
                visitPaNumExists = True
            Else
                paNumExists = False
                visitPaNumExists = False
            End If

            myConnection.Close()
            dataReader.Close()

            visitStatus = ""

            If (dictNVP("Patient Status Code") = "OP" Or dictNVP("Patient Status Code") = "OA") Then
                bolProcessInsurer = False
            Else
                bolProcessInsurer = True
            End If


            'code here to combine the separate room and bed fields from the generic feed
            strRoomBed = Replace(dictNVP.Item("01visit.room"), "'", "''") & Replace(dictNVP.Item("01visit.bed"), "'", "''")
            If Len(strRoomBed) = 0 Then strRoomBed = "None"
            'If dictNVP.Item("Event Type Code") = "A01" Then visitStatus = "IP"
            'If dictNVP.Item("Event Type Code") = "A03" Then visitStatus = "OEC"
            'If dictNVP.Item("Event Type Code") = "A04" Then visitStatus = "OP"
            'If dictNVP.Item("Event Type Code") = "A05" Then visitStatus = "PRE"
            'If dictNVP.Item("Event Type Code") = "A06" Then visitStatus = "IP"
            'If dictNVP.Item("Event Type Code") = "A07" Then visitStatus = "OP"

            '20140310 - use new status procedure
            visitStatus = calcSTARStatus(dictNVP)

            '20140318 - A06 and 7 is now built into the calsC=STARStatus routine.
            'If dictNVP.Item("Event Type Code") = "A06" Then visitStatus = "IP"
            'If dictNVP.Item("Event Type Code") = "A07" Then visitStatus = "OP"

            If dictNVP.Item("Event Type Code") = "A01" Then addit = True
            If dictNVP.Item("Event Type Code") = "A02" Then updateit = True
            If dictNVP.Item("Event Type Code") = "A03" Then updateit = True
            If dictNVP.Item("Event Type Code") = "A04" Then addit = True
            If dictNVP.Item("Event Type Code") = "A05" Then addit = True
            If dictNVP.Item("Event Type Code") = "A06" Then updateit = True
            If dictNVP.Item("Event Type Code") = "A07" Then updateit = True
            If dictNVP.Item("Event Type Code") = "A08" Then updateit = True
            If dictNVP.Item("Event Type Code") = "A13" Then updateit = True

            '20100114 - change A44 to A18
            If dictNVP.Item("Event Type Code") = "A18" Then updateit = True

            '6/26/2006
            If dictNVP.Item("Event Type Code") = "A17" Then updateit = True

            '4/18/2017 allow the process of A11
            If dictNVP.Item("Event Type Code") = "A11" Then updateit = True

            'check accident table if addit is true
            If (addit) Then
                'check for accidents on an A01,4 or 5
                sql = "select * from [08accidents] where panum = '" & dictNVP.Item("panum") & "'"
                objCommand.CommandText = sql
                myConnection.Open()
                dataReader = objCommand.ExecuteReader()

                If dataReader.HasRows Then
                    boolAccidentExists = True
                End If

                myConnection.Close()
                dataReader.Close()

            End If 'If (addit) Then

            If (addit) And (Not paNumExists) Then
                'this would be the case on an A01, A04 or A05 where the panum does not exist => add a visit record

                insertVisit(dictNVP)

                bolProcessInsurer = True
                visitPaNumExists = True
                'procedure to add entry in the new 08Accidents table for A01,A04 or A05 or update if entry exists

                If Replace(dictNVP.Item("panum"), "'", "''") <> "" And Not boolAccidentExists Then 'add it
                    sql = "INSERT [08Accidents] (panum, accmemo1, accmemo2, obx4, obx5, obx6, obx7, updated) "
                    sql = sql & "VALUES ("
                    sql = sql & "'" & gblORIGINAL_PA_NUMBER & "', "
                    sql = sql & "'" & Replace(dictNVP.Item("03insurer.accmemo1"), "'", "''") & "', "
                    sql = sql & "'" & Replace(dictNVP.Item("03insurer.accmemo2"), "'", "''") & "', "
                    sql = sql & "'" & Replace(dictNVP.Item("obx_4"), "'", "''") & "', "
                    sql = sql & "'" & Replace(dictNVP.Item("obx_5"), "'", "''") & "', "
                    sql = sql & "'" & Replace(dictNVP.Item("obx_6"), "'", "''") & "', "
                    sql = sql & "'" & Replace(dictNVP.Item("obx_7"), "'", "''") & "', "
                    sql = sql & "'" & DateTime.Now & "') "
                    updatecommand.CommandText = sql
                    myConnection.Open()
                    updatecommand.ExecuteNonQuery()
                    myConnection.Close()


                Else ' update the accident record
                    sql = "UPDATE [08accidents] "
                    sql = sql & "SET updated = '" & DateTime.Now & "'"
                    sql = sql & ", accmemo1 = '" & Replace(dictNVP.Item("03insurer.accmemo1"), "'", "''") & "'"
                    sql = sql & ", accmemo2 = '" & Replace(dictNVP.Item("03insurer.accmemo2"), "'", "''") & "'"
                    sql = sql & ", obx4 = '" & Replace(dictNVP.Item("obx_4"), "'", "''") & "'"
                    sql = sql & ", obx5 = '" & Replace(dictNVP.Item("obx_5"), "'", "''") & "'"
                    sql = sql & ", obx6 = '" & Replace(dictNVP.Item("obx_6"), "'", "''") & "'"
                    sql = sql & ", obx7 = '" & Replace(dictNVP.Item("obx_7"), "'", "''") & "'"
                    'sql = sql & ", accdate = '" & ConvertDate(dictNVP.Item("Accident Date/Time")) & "'"
                    sql = sql & " Where panum = '" & gblORIGINAL_PA_NUMBER & "'"
                    updatecommand.CommandText = sql
                    myConnection.Open()
                    updatecommand.ExecuteNonQuery()
                    myConnection.Close()
                End If 'add or update accident record


            Else 'If (addit) And (Not paNumExists)
                Select Case (strEventTypeCode)
                    Case "A01", "A04"
                        sql = "UPDATE [01Visit] "
                        sql = sql & "SET updated = '" & DateTime.Now & "'"

                        '20140220 - new OEC code in separate procedure
                        'If dictNVP.Item("Patient Type") = "OVD" Or dictNVP.Item("Patient Type") = "OVS" Or dictNVP.Item("Patient Type") = "OVC" Or dictNVP.Item("Patient Type") = "OVU" Then
                        'bolProcessInsurer = True
                        'sql = sql & ", Status = 'OEC'"
                        'sql = sql & ", OECLimit = 23"
                        'sql = sql & " ,dtChanged = '" & ConvertDate(dictNVP.Item("01visit.AdminDate")) & "'"
                        'If dictNVP.Item("01visit.AdminDate") <> "" Then
                        'sql = sql & ", OECExpire = '" & DateAdd("h", 23, CDate(ConvertDate(dictNVP.Item("01visit.AdminDate")))) & "'"
                        'End If
                        ''20130315 - moved visit status here so we dont get duplicate visit status data if OEC
                        'Else
                        sql = sql & ", Status = '" & visitStatus & "'"
                        'End If

                        '20121206 - add corpNo
                        sql = sql & " ,corpNo = " & gblCorporateNumber

                        sql = sql & " ,room = '" & strRoomBed & "'"

                        '20130522 - added department
                        sql = sql & STARupdateString("department", dictNVP.Item("department"))

                        '20090920 - add servicing facility from feed.
                        sql = sql & STARupdateString("servFacility", dictNVP.Item("Servicing Facility"))
                       
                        sql = sql & STARupdateString("diagnosis", dictNVP.Item("02diag.diagnosis"))


                        '20080110: added Admit Source to 01visit
                        sql = sql & STARupdateString("admSource", dictNVP.Item("Admit Source"))


                        If IsNumeric(dictNVP.Item("Admitting.patPhysNum")) Then
                            sql = sql & " ,physAdmit = " & dictNVP.Item("Admitting.patPhysNum")
                        ElseIf dictNVP.Item("Admitting.patPhysNum") = """""" Or dictNVP.Item("Admitting.patPhysNum") = "" Then
                            sql = sql & " ,physAdmit = NULL "
                        End If

                        '20121115 - add attending phys during McKesson Testing
                        If IsNumeric(dictNVP.Item("AttendingPhysicianID")) Then
                            sql = sql & " ,physAttend = " & dictNVP.Item("AttendingPhysicianID")
                        ElseIf dictNVP.Item("AttendingPhysicianID") = """""" Or dictNVP.Item("AttendingPhysicianID") = "" Then
                            sql = sql & " ,physAttend = NULL "
                        End If

                        If IsNumeric(dictNVP.Item("Referring.patPhysNum")) Then
                            sql = sql & " ,physRefer = " & dictNVP.Item("Referring.patPhysNum")
                        ElseIf dictNVP.Item("Referring.patPhysNum") = """""" Or dictNVP.Item("Referring.patPhysNum") = "" Then
                            sql = sql & " ,physRefer = NULL "
                        End If

                        If IsNumeric(dictNVP.Item("PCPNo")) Then
                            sql = sql & " ,physPrimary = " & dictNVP.Item("PCPNo")
                        ElseIf dictNVP.Item("PCPNo") = """""" Or dictNVP.Item("PCPNo") = "" Then
                            sql = sql & " ,physPrimary = NULL "
                        End If

                        If dictNVP.Item("AROFieldName") = "ARO" Then
                            sql = sql & " ,ARO = '" & Replace(dictNVP("ARODataField"), "'", "''") & "'"
                        End If
                        If Len(dictNVP.Item("AdvDirective")) >= 6 Then
                            sql = sql & " ,AdvanceDir = '" & Left$(dictNVP.Item("AdvDirective"), 6) & "'"
                        End If
                        '======================================================================================================
                        For i = 1 To gblInsCount
                            If i = 1 Then
                                tempstr = ""
                            End If
                            If i > 1 Then
                                tempstr = "_000" & i
                            End If

                            If dictNVP.Item("COBPriority" & tempstr) = "1" Then

                                '20130522 - use full star plancode including iplancode2
                                tempCOBPriority = Replace(dictNVP.Item("iplancode2" & tempstr), "'", "''") & Replace(dictNVP.Item("iplancode" & tempstr), "'", "''")
                            End If
                        Next
                        If tempCOBPriority <> "" Then
                            sql = sql & ", primaryPlan = '" & tempCOBPriority & "' "
                        End If
                        '=====================================================================================================
                        If Len(dictNVP.Item("Accident Date/Time")) > 4 Then
                            sql = sql & " ,accDate = '" & ConvertDate(dictNVP.Item("Accident Date/Time")) & "'"
                        End If

                        sql = sql & STARupdateString("allergies", dictNVP.Item("Allergy Description"))
                        ' If

                        If dictNVP("patient_height") <> "" Then
                            sql = sql & " ,height = '" & Replace(dictNVP("patient_height"), "'", "''") & "'"
                        End If
                        If dictNVP("patient_weight") <> "" Then
                            sql = sql & " ,weight = '" & Replace(dictNVP("patient_weight"), "'", "''") & "'"
                        End If


                        sql = sql & STARupdateString("patClass", dictNVP.Item("Patient Class"))
                        
                        sql = sql & STARupdateString("patType", dictNVP.Item("Patient Type"))
                        
                        '20140317 - changed to use new calcPatStatus routine
                        sql = sql & STARupdateString("patStatus", calcPatStatus(dictNVP))
                        'End If

                        If Len(dictNVP.Item("01visit.AdminDate")) > 7 Then
                            sql = sql & " ,AdminDate = '" & ConvertDate(dictNVP.Item("01visit.AdminDate")) & "'"
                        End If

                        If boolConsultingExists Then
                            For j = 0 To UBound(consultingArray)
                                If j < 10 Then
                                    If IsNumeric(consultingArray(j)) Then '6/7/2006
                                        If Trim(consultingArray(j)) <> "" Or IsDBNull(consultingArray(j)) Then
                                            sql = sql & ",consult" & Trim(Str(j + 1)) & " = " & (consultingArray(j))
                                        Else
                                            sql = sql & ",consult" & Trim(Str(j + 1)) & " = NULL"
                                        End If
                                    End If
                                End If
                            Next
                            If UBound(consultingArray) < 9 Then
                                For j = (UBound(consultingArray) + 1) To 9
                                    sql = sql & ",consult" & Trim(Str(j + 1)) & " = NULL"
                                Next
                            End If
                        End If

                        sql = sql & ", DCDate = NULL"
                        sql = sql & ", discharged = 0"


                        sql = sql & " Where PANum = '" & gblORIGINAL_PA_NUMBER & "'"
                        updatecommand.CommandText = sql
                        myConnection.Open()
                        updatecommand.ExecuteNonQuery()
                        myConnection.Close()

                        If Replace(dictNVP.Item("panum"), "'", "''") <> "" And Not boolAccidentExists Then 'add it


                            sql = "INSERT [08Accidents] (panum, accmemo1, accmemo2, obx4, obx5, obx6, obx7, updated) "
                            sql = sql & "VALUES ("
                            sql = sql & "'" & gblORIGINAL_PA_NUMBER & "', "
                            sql = sql & "'" & Replace(dictNVP.Item("03insurer.accmemo1"), "'", "''") & "', "
                            sql = sql & "'" & Replace(dictNVP.Item("03insurer.accmemo2"), "'", "''") & "', "
                            sql = sql & "'" & Replace(dictNVP.Item("obx_4"), "'", "''") & "', "
                            sql = sql & "'" & Replace(dictNVP.Item("obx_5"), "'", "''") & "', "
                            sql = sql & "'" & Replace(dictNVP.Item("obx_6"), "'", "''") & "', "
                            sql = sql & "'" & Replace(dictNVP.Item("obx_7"), "'", "''") & "', "
                            sql = sql & "'" & DateTime.Now & "') "
                            updatecommand.CommandText = sql
                            myConnection.Open()
                            updatecommand.ExecuteNonQuery()
                            myConnection.Close()

                        Else ' update the accident record
                            sql = "UPDATE [08accidents] "
                            sql = sql & "SET updated = '" & DateTime.Now & "'"
                            sql = sql & ", accmemo1 = '" & Replace(dictNVP.Item("03insurer.accmemo1"), "'", "''") & "'"
                            sql = sql & ", accmemo2 = '" & Replace(dictNVP.Item("03insurer.accmemo2"), "'", "''") & "'"
                            sql = sql & ", obx4 = '" & Replace(dictNVP.Item("obx_4"), "'", "''") & "'"
                            sql = sql & ", obx5 = '" & Replace(dictNVP.Item("obx_5"), "'", "''") & "'"
                            sql = sql & ", obx6 = '" & Replace(dictNVP.Item("obx_6"), "'", "''") & "'"
                            sql = sql & ", obx7 = '" & Replace(dictNVP.Item("obx_7"), "'", "''") & "'"
                            sql = sql & " Where panum = '" & gblORIGINAL_PA_NUMBER & "'"
                            updatecommand.CommandText = sql
                            myConnection.Open()
                            updatecommand.ExecuteNonQuery()
                            myConnection.Close()
                        End If

                End Select
            End If 'If (addit) And (Not paNumExists)

            If (updateit) And (paNumExists) Then

                Select Case (strEventTypeCode)
                    '======================================================================================
                    Case "A02" ' room change
                        '======================================================================================
                        'get the current manager for the room


                        If Len(strRoomBed) > 4 Then
                            sql = "UPDATE [01Visit] "
                            sql = sql & "SET room = '" & Replace(strRoomBed, "'", "''") & "'"
                            'sql = sql & ", currmanNo = " & intCurrmanNo & " "
                            sql = sql & ", currmantype = 'CM'"
                            '20120817
                            'If dictNVP.Item("Hospital Service") = "OEC" And (Mid$(dictNVP.Item("Message Control ID"), 5, 4) = "OPBD" Or Mid$(dictNVP.Item("Message Control ID"), 5, 4) = "XFR1") Then
                            'If dictNVP.Item("Hospital Service") = "OEC" And dictNVP.Item("OPRP.date") <> "" Then
                            ''-------------------------------------------------------------------------
                            'bolProcessInsurer = True
                            'sql = sql & ", Status = 'OEC'"
                            'sql = sql & ", OECLimit = 23"
                            'sql = sql & " ,dtChanged = '" & ConvertDate(dictNVP.Item("OPRP.date")) & "'"
                            'If dictNVP.Item("OPRP.date") <> "" Then
                            'sql = sql & ", OECExpire = '" & DateAdd("h", 23, CDate(ConvertDate(dictNVP.Item("OPRP.date")))) & "'"
                            'End If
                            'End If


                            sql = sql & ", updated = '" & DateTime.Now & "'"
                            sql = sql & " Where PANum = '" & gblORIGINAL_PA_NUMBER & "'"
                            updatecommand.CommandText = sql
                            myConnection.Open()
                            updatecommand.ExecuteNonQuery()
                            myConnection.Close()
                        End If

            '======================================================================================
                    Case "A03" 'discharged

                        '======================================================================================
                        sql = "UPDATE [01Visit] "
                        sql = sql & "SET updated = '" & DateTime.Now & "' "

                        sql = sql & STARupdateString("DCDisp", dictNVP.Item("Discharge Disposition"))

                        'End If
                        If Len(dictNVP("01visit.dischdate")) > 7 Then
                            sql = sql & ", DCDate = '" & ConvertDate(dictNVP.Item("01visit.dischdate")) & "'"
                        End If

                        '20130620 - update patStatus from calcStatus function;not Patient Status Code
                        sql = sql & ", patStatus = '" & calcPatStatus(dictNVP) & "'"
                        sql = sql & ", discharged = 1"

                        sql = sql & " Where PANum = '" & gblORIGINAL_PA_NUMBER & "'"
                        updatecommand.CommandText = sql
                        myConnection.Open()
                        updatecommand.ExecuteNonQuery()
                        myConnection.Close()

                        '======================================================================================
                    Case "A11"  'cancel admit
                        '20140317 - handle A11 like A03. Change also made to calcSTARStatus
                        '======================================================================================
                        sql = "UPDATE [01Visit] "
                        sql = sql & "SET updated = '" & DateTime.Now & "' "

                        sql = sql & STARupdateString("DCDisp", dictNVP.Item("Discharge Disposition"))

                        'End If
                        If Len(dictNVP("01visit.dischdate")) > 7 Then
                            sql = sql & ", DCDate = '" & ConvertDate(dictNVP.Item("01visit.dischdate")) & "'"
                        End If
                        '20130620 - update patStatus from calcStatus function;not Patient Status Code
                        '20140317 - changed to use new calcPatStatus routine
                        sql = sql & ", patStatus = '" & calcPatStatus(dictNVP) & "'"
                        sql = sql & ", discharged = 1"

                        sql = sql & " Where PANum = '" & gblORIGINAL_PA_NUMBER & "'"
                        updatecommand.CommandText = sql
                        myConnection.Open()
                        updatecommand.ExecuteNonQuery()
                        myConnection.Close()
                        '======================================================================================
                    Case "A06" 'status transfer from outpatient to inpatient
                        '======================================================================================
                        'get the current manager for the room

                        strCurrman = "No Room Number"


                        sql = "UPDATE [01Visit] "

                        sql = sql & "SET updated = '" & DateTime.Now & "' "

                        If dictNVP("Patient Class") <> "" Then
                            sql = sql & " ,patClass = '" & Replace(dictNVP("Patient Class"), "'", "''") & "'"
                        End If

                        '20080110: added Admit Source to 01visit
                        If dictNVP("Admit Source") <> "" Then
                            sql = sql & " ,admSource = '" & Replace(dictNVP("Admit Source"), "'", "''") & "'"
                        End If

                        If dictNVP("Patient Type") <> "" Then
                            sql = sql & " ,patType = '" & Replace(dictNVP("Patient Type"), "'", "''") & "'"
                        End If

                        '20130620 -update patStatus from calcStatus function
                        '20140317 - changed to use new calcPatStatus routine
                        sql = sql & " ,patStatus = '" & calcPatStatus(dictNVP) & "'"

                        
                        If Len(strRoomBed) > 4 Then
                            sql = sql & " ,room = '" & Replace(strRoomBed, "'", "''") & "'"
                            'sql = sql & " ,currmanNo = " & intCurrmanNo & " "
                            sql = sql & " ,currmantype = 'CM'"
                        End If
                        sql = sql & " ,Status = '" & visitStatus & "'"
                        sql = sql & " Where PANum = '" & gblORIGINAL_PA_NUMBER & "'"
                        updatecommand.CommandText = sql
                        myConnection.Open()
                        updatecommand.ExecuteNonQuery()
                        myConnection.Close()
                        '======================================================================================
                    Case "A07" 'status transfer from inpatient to outpatient
                        '======================================================================================

                        strCurrman = "No Room Number"



                        sql = "UPDATE [01Visit] "
                        sql = sql & "SET updated = '" & DateTime.Now & "' "

                        If dictNVP("Patient Class") <> "" Then
                            sql = sql & " ,patClass = '" & Replace(dictNVP("Patient Class"), "'", "''") & "'"
                        End If

                        If dictNVP("Patient Type") <> "" Then
                            sql = sql & " ,patType = '" & Replace(dictNVP("Patient Type"), "'", "''") & "'"
                        End If

                        '20130620 -update patStatus from calcStatus function
                        'sql = sql & " ,patStatus = '" & calcStatus(dictNVP) & "'"

                        '20140317 - new patstatus for A07
                        sql = sql & " ,patStatus = 'OA'"

                        If Len(strRoomBed) > 4 Then
                            sql = sql & " ,room = '" & Replace(strRoomBed, "'", "''") & "'"
                            'sql = sql & " ,currmanNo = " & intCurrmanNo & " "
                            sql = sql & " ,currmantype = 'CM'"
                        End If
                        '20120817
                        'If dictNVP.Item("Hospital Service") = "OEC" Then
                        'If (dictNVP.Item("Hospital Service") = "OEC" And dictNVP.Item("OPRP.date") <> "") Then
                        'sql = sql & ", Status = 'OEC'"
                        ''sql = sql & ", OECExpire = '" & DateAdd("h", 23, CDate(ConvertDate(dictNVP.Item("01visit.AdminDate")))) & "'"
                        'sql = sql & ", OECExpire = '" & DateAdd("h", 23, CDate(ConvertDate(dictNVP.Item("OPRP.date")))) & "'"
                        'Else
                        sql = sql & " ,Status = '" & visitStatus & "'"
                        'End If

                        sql = sql & " Where PANum = '" & gblORIGINAL_PA_NUMBER & "'"
                        updatecommand.CommandText = sql
                        myConnection.Open()
                        updatecommand.ExecuteNonQuery()
                        myConnection.Close()
                        '======================================================================================
                    Case "A08" 'update everything
                        '======================================================================================
                        sql = "UPDATE [01Visit] "
                        sql = sql & "SET updated = '" & DateTime.Now & "'"

                        If Len(dictNVP.Item("01visit.AdminDate")) > 7 Then
                            sql = sql & " ,AdminDate = '" & ConvertDate(dictNVP.Item("01visit.AdminDate")) & "'"
                        End If
                        '20120817
                        'If dictNVP.Item("Hospital Service") = "OEC" And Mid$(dictNVP.Item("Message Control ID"), 5, 4) = "OPRP" Then
                        'If (dictNVP.Item("Hospital Service") = "OEC" And dictNVP.Item("OPRP.date") <> "") Then
                        bolProcessInsurer = True
                        'sql = sql & ", Status = 'OEC'"
                        'sql = sql & ", OECLimit = 23"
                        'sql = sql & " ,dtChanged = '" & ConvertDate(dictNVP.Item("OPRP.date")) & "'"
                        'If dictNVP.Item("OPRP.date") <> "" Then
                        'sql = sql & ", OECExpire = '" & DateAdd("h", 23, CDate(ConvertDate(dictNVP.Item("OPRP.date")))) & "'"
                        'End If

                        'End If

                        'If dictNVP("02diag.diagnosis") <> "" Then
                        'sql = sql & " ,diagnosis = '" & Replace(dictNVP("02diag.diagnosis"), "'", "''") & "'"
                        'End If
                        'If Len(dictNVP.Item("02diag.diagnosis")) > 0 Then
                        sql = sql & STARupdateString("diagnosis", dictNVP.Item("02diag.diagnosis"))
                        'End If

                        '20090920 - add servicing facility from feed.
                        'If Not IsDBNull(dictNVP.Item("Servicing Facility")) Then
                        'sql = sql & " ,servFacility = '" & Replace(dictNVP("Servicing Facility"), "'", "''") & "'"
                        'End If
                        'If Len(dictNVP.Item("Servicing Facility")) > 0 Then
                        sql = sql & STARupdateString("servFacility", dictNVP.Item("Servicing Facility"))
                        'End If
                        '20090920 - end

                        '20130522 - added department
                        sql = sql & STARupdateString("department", dictNVP.Item("department"))

                        '20080110: added Admit Source to 01visit
                        'If dictNVP("Admit Source") <> "" Then
                        'sql = sql & " ,admSource = '" & Replace(dictNVP("Admit Source"), "'", "''") & "'"
                        'End If
                        'If Len(dictNVP.Item("Admit Source")) > 0 Then
                        sql = sql & STARupdateString("admSource", dictNVP.Item("Admit Source"))
                        'End If


                        If IsNumeric(dictNVP.Item("Admitting.patPhysNum")) Then
                            sql = sql & " ,physAdmit = " & dictNVP.Item("Admitting.patPhysNum")
                        ElseIf dictNVP.Item("Admitting.patPhysNum") = """""" Or dictNVP.Item("Admitting.patPhysNum") = "" Then
                            sql = sql & " ,physAdmit = NULL"
                        End If

                        '20121115 - add attending phys during McKesson Testing
                        If IsNumeric(dictNVP.Item("AttendingPhysicianID")) Then
                            sql = sql & " ,physAttend = " & dictNVP.Item("AttendingPhysicianID")
                        ElseIf dictNVP.Item("AttendingPhysicianID") = """""" Or dictNVP.Item("AttendingPhysicianID") = "" Then
                            sql = sql & " ,physAttend = NULL"
                        End If


                        If IsNumeric(dictNVP.Item("Referring.patPhysNum")) Then
                            sql = sql & " ,physRefer = " & dictNVP.Item("Referring.patPhysNum")
                        ElseIf dictNVP.Item("Referring.patPhysNum") = """""" Or dictNVP.Item("Referring.patPhysNum") = "" Then
                            sql = sql & " ,physRefer = NULL"
                        End If

                        If IsNumeric(dictNVP.Item("PCPNo")) Then
                            sql = sql & " ,physPrimary = " & dictNVP.Item("PCPNo")
                        ElseIf dictNVP.Item("PCPNo") = """""" Or dictNVP.Item("PCPNo") = "" Then
                            sql = sql & " ,physPrimary = NULL"
                        End If

                        If dictNVP.Item("AROFieldName") = "ARO" Then
                            sql = sql & " ,ARO = '" & Replace(dictNVP("ARODataField"), "'", "''") & "'"
                        End If

                        If Len(dictNVP.Item("AdvDirective")) >= 6 Then
                            sql = sql & " ,AdvanceDir = '" & Left$(dictNVP.Item("AdvDirective"), 6) & "'"
                        End If

                        '======================================================================================================
                        '5/31/2006
                        'If (dictNVP("1.COBPriority") = "1" And Len(dictNVP("1.iplancode")) > 2) Then
                        'sql = sql & " ,primaryPlan = '" & dictNVP("1.iplancode") & "'"
                        'ElseIf (dictNVP("2.COBPriority") = "1" And Len(dictNVP("2.iplancode")) > 2) Then
                        'sql = sql & " ,primaryPlan = '" & dictNVP("2.iplancode") & "'"
                        'ElseIf (dictNVP("3.COBPriority") = "1" And Len(dictNVP("3.iplancode")) > 2) Then
                        'sql = sql & " ,primaryPlan = '" & dictNVP("3.iplancode") & "'"
                        'End If

                        For i = 1 To gblInsCount
                            If i = 1 Then
                                tempstr = ""
                            End If
                            If i > 1 Then
                                tempstr = "_000" & i
                            End If

                            If dictNVP.Item("COBPriority" & tempstr) = "1" Then
                                'sql = sql & ", primaryIns = '" & Replace(dictNVP.Item("iplancode" & tempstr), "'", "''") & "' "
                                '20130522 - use star plancode witch includes iplancode2
                                tempCOBPriority = Replace(dictNVP.Item("iplancode2" & tempstr), "'", "''") & Replace(dictNVP.Item("iplancode" & tempstr), "'", "''")
                            End If
                        Next
                        If tempCOBPriority <> "" Then
                            sql = sql & ", primaryPlan = '" & tempCOBPriority & "' "
                        End If
                        '=====================================================================================================


                        If Len(dictNVP.Item("Accident Date/Time")) > 4 Then
                            sql = sql & " ,accDate = '" & ConvertDate(dictNVP.Item("Accident Date/Time")) & "'"
                        End If

                        'If dictNVP("Allergy Description") <> "" Then
                        'sql = sql & " ,allergies = '" & Replace(dictNVP("Allergy Description"), "'", "''") & "'"
                        'End If
                        'If Len(dictNVP.Item("Allergy Description")) > 0 Then
                        sql = sql & STARupdateString("allergies", dictNVP.Item("Allergy Description"))
                        'End If

                        If dictNVP("patient_height") <> "" Then
                            sql = sql & " ,height = '" & Replace(dictNVP("patient_height"), "'", "''") & "'"
                        End If
                        If dictNVP("patient_weight") <> "" Then
                            sql = sql & " ,weight = '" & Replace(dictNVP("patient_weight"), "'", "''") & "'"
                        End If

                        'If dictNVP("Patient Class") <> "" Then
                        'sql = sql & " ,patClass = '" & Replace(dictNVP("Patient Class"), "'", "''") & "'"
                        'End If
                        If Len(dictNVP.Item("Patient Class")) > 0 Then
                            sql = sql & STARupdateString("patClass", dictNVP.Item("Patient Class"))
                        End If

                        'If dictNVP("Patient Type") <> "" Then
                        'sql = sql & " ,patType = '" & Replace(dictNVP("Patient Type"), "'", "''") & "'"
                        'End If
                        If Len(dictNVP.Item("Patient Type")) > 0 Then
                            sql = sql & STARupdateString("patType", dictNVP.Item("Patient Type"))
                        End If


                        '20140317 - added back in and use calcPatStatus routine
                        'If dictNVP("Patient Status Code") <> "" Then
                        sql = sql & " ,patStatus = '" & calcPatStatus(dictNVP) & "'"
                        'End If

                        '20130620 - don't update patstatus on an A08 record
                        'If Len(dictNVP.Item("Patient Status Code")) > 0 Then
                        'sql = sql & STARupdateString("patStatus", dictNVP.Item("Patient Status Code"))
                        'End If



                        'If Len(dictNVP.Item("Discharge Disposition")) > 0 Then
                        'sql = sql & ", DCDisp = '" & Replace(dictNVP.Item("Discharge Disposition"), "'", "''") & "'"
                        'End If
                        If Len(dictNVP.Item("Discharge Disposition")) > 0 Then
                            sql = sql & STARupdateString("DCDisp", dictNVP.Item("Discharge Disposition"))
                        End If

                        '2/8/2006 - handle consulting physicians
                        If boolConsultingExists Then
                            For j = 0 To UBound(consultingArray)
                                If j < 10 Then
                                    If Trim(consultingArray(j)) <> "" Or IsDBNull(consultingArray(j)) Then
                                        sql = sql & ",consult" & Trim(Str(j + 1)) & " = " & (consultingArray(j))
                                    Else
                                        sql = sql & ",consult" & Trim(Str(j + 1)) & " = NULL"
                                    End If
                                End If
                            Next
                            If UBound(consultingArray) < 9 Then
                                For j = (UBound(consultingArray) + 1) To 9
                                    sql = sql & ",consult" & Trim(Str(j + 1)) & " = NULL"
                                Next
                            End If
                        End If

                        '20140723 - add visit status update
                        sql = sql & ", status = '" & visitStatus & "' "

                        sql = sql & " Where PANum = '" & gblORIGINAL_PA_NUMBER & "'"
                        updatecommand.CommandText = sql
                        myConnection.Open()
                        updatecommand.ExecuteNonQuery()
                        myConnection.Close()

                        '======================================================================================
                    Case "A13" 'cancel discharge
                        '======================================================================================
                        sql = "UPDATE [01Visit] "
                        sql = sql & "SET updated = '" & DateTime.Now & "'"
                        sql = sql & ", DCDate = NULL"
                        sql = sql & ", discharged = 0"
                        '12/10/2002 - added status
                        '20130620 -update patStatus from calcStatus function
                        '20140317 - changed to calcPatStatus routine and added back in status field.
                        '20140319 - added code for A13 in calcPatStatus
                        sql = sql & " ,patStatus = '" & calcPatStatus(dictNVP) & "'"

                        '20140319 - don't change status on A13 because it was not changed on an A03
                        'sql = sql & " ,Status = '" & calcSTARStatus(dictNVP) & "'"

                        sql = sql & " Where PANum = '" & gblORIGINAL_PA_NUMBER & "'"
                        updatecommand.CommandText = sql
                        myConnection.Open()
                        updatecommand.ExecuteNonQuery()
                        myConnection.Close()
                        '======================================================================================
                    Case "A17" 'roomchange
                        '======================================================================================

                        Dim processA17 As Boolean = True
                        Dim strRoomBed1 As String = ""
                        Dim strRoomBed2 As String = ""
                        Dim panum1 As String = ""
                        Dim panum2 As String = ""
                        strRoomBed1 = Replace(dictNVP.Item("01visit.room"), "'", "''") & Replace(dictNVP.Item("01visit.bed"), "'", "''")
                        strRoomBed2 = Replace(dictNVP.Item("01visit.room_0002"), "'", "''") & Replace(dictNVP.Item("01visit.bed_0002"), "'", "''")
                        If (dictNVP.Item("panum") <> "") Then
                            panum1 = dictNVP.Item("panum")
                        Else
                            processA17 = False
                        End If

                        If (dictNVP.Item("panum_0002") <> "") Then
                            panum2 = dictNVP.Item("panum_0002")
                        Else
                            processA17 = False
                        End If
                        If processA17 Then
                            'process first panum - 6/26/2006
                            sql = "UPDATE [01visit] "
                            sql = sql & "SET updated = '" & DateTime.Now & "'"
                            sql = sql & ", room = '" & strRoomBed2 & "'"
                            sql = sql & " Where panum = '" & panum1 & "'"
                            updatecommand.CommandText = sql
                            myConnection.Open()
                            updatecommand.ExecuteNonQuery()
                            myConnection.Close()
                            'process second panum
                            sql = "UPDATE [0visit] "
                            sql = sql & "SET updated = '" & DateTime.Now & "'"
                            sql = sql & ", room = '" & strRoomBed1 & "'"
                            sql = sql & " Where panum = '" & panum2 & "'"
                            updatecommand.CommandText = sql
                            myConnection.Open()
                            updatecommand.ExecuteNonQuery()

                            myConnection.Close()

                        End If

                End Select

                '======================================================================================

                '20100114 - changes to A18 processing from A44 processing to handle mrnum change in visit table only.
                If dictNVP.Item("Event Type Code") = "A18" Then

                    '20130225 - check for empty values
                    If (Trim(dictNVP.Item("mrnum")) <> "" And Trim(dictNVP.Item("oldMrNum")) <> "") Then

                        sql = "Update [01visit] "
                        sql = sql & "SET updated = '" & DateTime.Now & "'"
                        sql = sql & ", mrnum = " & dictNVP.Item("mrnum")
                        sql = sql & " Where mrnum = " & dictNVP.Item("oldMrNum")
                        updatecommand.CommandText = sql
                        myConnection.Open()
                        updatecommand.ExecuteNonQuery()
                        myConnection.Close()
                    End If
                End If
                '======================================================================================


            End If 'If (updateit) And (paNumExists)
            'LogFile.Write("Visit Process Complete: ")
            'LogFile.WriteLine(gblORIGINAL_PA_NUMBER + " (" + dictNVP.Item("Event Type Code") + ")")
        Catch ex As Exception
            functionError = True
            gblLogString = gblLogString & "Visit Error" & vbCrLf
            gblLogString = gblLogString & ex.Message & vbCrLf
            'LogFile.Close()
            Exit Sub
        End Try
    End Sub 'addVisit
    Public Sub insertOneInsurer(ByVal dictNVP As Hashtable, ByVal tempstr As String)
        '20140603 - code to use star plancode in the iplancode field of 03insurer.
        '20140904 - return to integer fclass.
        Dim strSubName As String
        Dim strPolicyIssueDate As String
        Dim strAuthServices As String
        Dim intFClass As Integer = 0 '20140904
        Dim myConnection As New SqlConnection(connectionString)
        Dim objCommand As New SqlCommand
        Dim updatecommand As New SqlCommand
        updatecommand.Connection = myConnection
        ' Dim strFClass As String = "" '20140904
        Dim STAR_Plancode As String = ""
        objCommand.Connection = myConnection
        Dim dataReader As SqlDataReader '20140904

        Try
            intFClass = 0
            sql = "select id from [104finclass] where finclass = '" & dictNVP.Item("insurer.fClass") & "' and inactive = 0" '20140904
            objCommand.CommandText = sql
            'Console.WriteLine("{0}", sql)
            'Console.ReadLine()
            myConnection.Open()
            dataReader = objCommand.ExecuteReader()
            While dataReader.Read()
                intFClass = dataReader.GetInt32(0)
            End While

            myConnection.Close()
            dataReader.Close()

            STAR_Plancode = Trim(Replace(dictNVP("iplancode2" & tempstr), "'", "''")) & Trim(Replace(dictNVP("iplancode" & tempstr), "'", "''"))

            If STAR_Plancode <> "" Then '6/1/2006

                strAuthServices = dictNVP.Item("authservices1" & tempstr) & " " & dictNVP.Item("authservices2" & tempstr)
                If Len(dictNVP.Item("PolicyIssueDate" & tempstr)) = 8 Then
                    strPolicyIssueDate = ConvertDate(dictNVP.Item("PolicyIssueDate" & tempstr))
                Else
                    strPolicyIssueDate = ""
                End If

                strSubName = dictNVP.Item("Insured First Name" & tempstr)
                strSubName = strSubName & " " & dictNVP.Item("Insured Middle Name" & tempstr)
                strSubName = strSubName & " " & dictNVP("Insured Last Name" & tempstr)
                'remove these later
                gblORIGINAL_PA_NUMBER = dictNVP.Item("panum")

                '20080331 changed updated to created to be populated when insurer added per projman
                sql = "Insert [03insurer] "
                sql = sql & "(panum, coName, coNUm, iPlanCode, iplancode2, star_plancode, subname, policyNum, "
                sql = sql & "authNum1, theGroup, PIssue, aprimary, FClass, reqCert, "
                sql = sql & "addr1, addr2, city, state, zip, phone, preCertAuth, subEmployer, precertcontact, insuredDOB, insuredSex, " ' 10/19/2005
                'sql = sql & "updated) "

                '20170804 - Keep track of insurance changes
                'sql = sql & "created) "
                sql = sql & "created, lastupdatedStar) "

                sql = sql & "VALUES ("
                sql = sql & "'" & gblORIGINAL_PA_NUMBER & "', "
                insertString(dictNVP.Item("CompanyName" & tempstr))
                insertString(dictNVP("iplancode2" & tempstr))

                '20140603 - put star plancode in the iplancode field
                'insertString(dictNVP("iplancode" & tempstr))
                insertString(STAR_Plancode)

                insertString(dictNVP("iplancode2" & tempstr))
                insertString(STAR_Plancode)
                insertString(strSubName)
                insertString(dictNVP.Item("PolicyNumber" & tempstr))
                insertString(dictNVP.Item("AuthNum" & tempstr))
                '20170512 parse out primary AuthNum from new IN1_14
                'Dim autharray() As String = dictNVP.Item("AuthNum" & tempstr).split("~")
                'Dim Authcode As String = Replace(autharray(0), "^", "")
                'sql = sql & "'" & Authcode & "', "
                'insertString(Authcode)

                insertString(dictNVP.Item("group" & tempstr))
                insertString(strPolicyIssueDate)

                If dictNVP("COBPriority" & tempstr) = "1" Then
                    sql = sql & "1, " & intFClass & ", " '20140904
                Else
                    sql = sql & "0, 0, "
                End If

                sql = sql & "1, "
                insertString(dictNVP.Item("addr1" & tempstr))
                insertString(dictNVP.Item("addr2" & tempstr))
                insertString(dictNVP.Item("city" & tempstr))
                insertString(dictNVP.Item("state" & tempstr))
                insertString(dictNVP.Item("zip" & tempstr))
                insertString(Left(dictNVP.Item("phone" & tempstr), 15))
                insertString(strAuthServices)
                insertString(dictNVP.Item("SubscriberEmployer" & tempstr))
                insertString(dictNVP.Item("precertcontact" & tempstr))

                insertString(ConvertDate(dictNVP.Item("Insured DOB" & tempstr)))
                insertString(dictNVP.Item("Insured Sex" & tempstr))

                sql = sql & "'" & DateTime.Now & "', "
                '20170804 - Keep track of insurance changes
                sql = sql & "'" & DateTime.Now & "') "

                updatecommand.CommandText = sql
                myConnection.Open()
                updatecommand.ExecuteNonQuery()
                myConnection.Close()

                ProcessIN1_14(dictNVP, tempstr)
            End If

        Catch ex As Exception
            functionError = True
            gblLogString = gblLogString & "Single Insurer Insert Error" & vbCrLf
            gblLogString = gblLogString & ex.Message & vbCrLf
            'LogFile.Close()
            Exit Sub


        End Try
    End Sub
    Public Sub insertInsurer(ByVal dictNVP As Hashtable)
        '20140603 - code to use star plancode in the iplancode field of 03insurer.
        '20140904 - return to integer fclass.

        Dim i As Integer
        Dim strSubName As String
        Dim strPolicyIssueDate As String
        Dim strAuthServices As String
        Dim tempStr As String
        Dim strFClass As String = ""
        Dim intFClass As Integer = 0 '20140904
        Dim myConnection As New SqlConnection(connectionString)
        Dim objCommand As New SqlCommand
        Dim updatecommand As New SqlCommand
        updatecommand.Connection = myConnection
        Dim STAR_Plancode As String = ""
        objCommand.Connection = myConnection
        Dim dataReader As SqlDataReader '20140904

        Try
            intFClass = 0
            sql = "select id from [104finclass] where finclass = '" & dictNVP.Item("insurer.fClass") & "' and inactive = 0" '20140904
            objCommand.CommandText = sql
            'Console.WriteLine("{0}", sql)
            'Console.ReadLine()
            myConnection.Open()
            dataReader = objCommand.ExecuteReader()
            While dataReader.Read()
                intFClass = dataReader.GetInt32(0)
            End While

            myConnection.Close()
            dataReader.Close()
            'strFClass = dictNVP.Item("insurer.fClass") '20140904
            i = 0
            tempStr = ""
            For i = 1 To gblInsCount
                If i = 1 Then
                    tempStr = ""
                End If
                If i > 1 Then
                    tempStr = "_000" & i
                End If

                sql = ""
                '20130515 - Added STAR_Plancode to [03insurer] table
                STAR_Plancode = Trim(Replace(dictNVP("iplancode2" & tempStr), "'", "''")) & Trim(Replace(dictNVP("iplancode" & tempStr), "'", "''"))

                If STAR_Plancode <> "" Then '6/1/2006

                    strAuthServices = dictNVP.Item("authservices1" & tempStr) & " " & dictNVP.Item("authservices2" & tempStr)
                    If Len(dictNVP.Item("PolicyIssueDate" & tempStr)) = 8 Then
                        strPolicyIssueDate = ConvertDate(dictNVP.Item("PolicyIssueDate" & tempStr))
                    Else
                        strPolicyIssueDate = ""
                    End If

                    strSubName = dictNVP.Item("Insured First Name" & tempStr)
                    strSubName = strSubName & " " & dictNVP.Item("Insured Middle Name" & tempStr)
                    strSubName = strSubName & " " & dictNVP("Insured Last Name" & tempStr)
                    'remove these later
                    gblORIGINAL_PA_NUMBER = dictNVP.Item("panum")

                    '20080331 changed updated to created to be populated when insurer added per projman
                    sql = "Insert [03insurer] "
                    sql = sql & "(panum, coName, coNUm, iPlanCode, iplancode2, star_plancode, subname, policyNum, "
                    sql = sql & "authNum1, theGroup, PIssue, aprimary, FClass, reqCert, "
                    sql = sql & "addr1, addr2, city, state, zip, phone, preCertAuth, subEmployer, precertcontact, insuredDOB, insuredSex, " ' 10/19/2005
                    'sql = sql & "updated) "

                    '20170804 - Keep track of insurance changes
                    'sql = sql & "created) "
                    sql = sql & "created, lastupdatedStar) "

                    sql = sql & "VALUES ("
                    sql = sql & "'" & gblORIGINAL_PA_NUMBER & "', "
                    insertString(dictNVP.Item("CompanyName" & tempStr))
                    insertString(dictNVP("iplancode2" & tempStr))

                    '20140603 - put star plancode in iplancode field
                    'insertString(dictNVP("iplancode" & tempStr))
                    insertString(STAR_Plancode)

                    insertString(dictNVP("iplancode2" & tempStr))
                    insertString(STAR_Plancode)
                    insertString(strSubName)
                    insertString(dictNVP.Item("PolicyNumber" & tempStr))
                    insertString(dictNVP.Item("AuthNum" & tempStr))
                    '20170512 parse out primary AuthNum from new IN1_14
                    'Dim autharray() As String = dictNVP.Item("AuthNum" & tempStr).split("~")
                    'Dim Authcode As String = Replace(autharray(0), "^", "")
                    'sql = sql & "'" & Authcode & "', "
                    'insertString(Authcode)

                    insertString(dictNVP.Item("group" & tempStr))
                    insertString(strPolicyIssueDate)

                    If dictNVP("COBPriority" & tempStr) = "1" Then
                        sql = sql & "1, " & intFClass & ", " '20140904
                    Else
                        sql = sql & "0, 0, "
                    End If

                    sql = sql & "1, "
                    insertString(dictNVP.Item("addr1" & tempStr))
                    insertString(dictNVP.Item("addr2" & tempStr))
                    insertString(dictNVP.Item("city" & tempStr))
                    insertString(dictNVP.Item("state" & tempStr))
                    insertString(dictNVP.Item("zip" & tempStr))
                    insertString(Left(dictNVP.Item("phone" & tempStr), 15))
                    insertString(strAuthServices)
                    insertString(dictNVP.Item("SubscriberEmployer" & tempStr))
                    insertString(dictNVP.Item("precertcontact" & tempStr))

                    insertString(ConvertDate(dictNVP.Item("Insured DOB" & tempStr)))
                    insertString(dictNVP.Item("Insured Sex" & tempStr))

                    sql = sql & "'" & DateTime.Now & "', "
                    '20170804 - Keep track of insurance changes
                    sql = sql & "'" & DateTime.Now & "') "

                    updatecommand.CommandText = sql
                    myConnection.Open()
                    updatecommand.ExecuteNonQuery()
                    myConnection.Close()
                    gblLogString = gblLogString & STAR_Plancode & vbCrLf

                    ProcessIN1_14(dictNVP, tempStr)
                End If 'If Len(dictNVP("iplancode" & tempstr)) = 3 Then '6/1/2006

            Next 'i to gblInsCount
        Catch ex As Exception
            functionError = True
            gblLogString = gblLogString & "Insurer Error" & vbCrLf
            gblLogString = gblLogString & ex.Message & vbCrLf
            'LogFile.Close()
            Exit Sub
        End Try
    End Sub

    Public Sub checkSegments(ByVal dictNVP As Hashtable)
        Dim IN1Count As Integer = 0
        Dim AL1count As Integer = 0
        Try
            gblAL1Count = 0 ' 20140917
            Dim myEnumerator As IDictionaryEnumerator = dictNVP.GetEnumerator()


            'For Each value In dictNVP.values
            'Console.WriteLine(value)

            'Next
            myEnumerator.Reset()
            While myEnumerator.MoveNext()
                If Left$(myEnumerator.Key, 8) = "setIDIN1" Then
                    gblInsCount = gblInsCount + 1
                    IN1Count = IN1Count + 1
                End If

                '20140215 - mod to look for just three letters like itw
                If Left$(myEnumerator.Key, 3) = "AL1" Then
                    gblAL1Count = gblAL1Count + 1
                    AL1count = AL1count + 1
                End If

            End While
            'gblLogString = gblLogString & " IN1count =  "
            'gblLogString = gblLogString & IN1Count & vbCrLf
        Catch ex As Exception
            functionError = True
            gblLogString = gblLogString & "Segment Enumeration Error" & vbCrLf
            gblLogString = gblLogString & ex.Message & vbCrLf
            'LogFile.Close()
            Exit Sub
        End Try
    End Sub

    Public Sub insertVisit(ByVal dictNVP As Hashtable)
        '5/31/2006 - added code to handle primary plan
        '20121206 - add corpNo
        Dim tempPlanCode As String = ""
        Dim i As Integer
        Dim tempstr As String = ""

        Dim myConnection As New SqlConnection(connectionString)
        Dim objCommand As New SqlCommand
        Dim updatecommand As New SqlCommand
        Dim strRoomBed As String = ""
        updatecommand.Connection = myConnection

        objCommand.Connection = myConnection
        'Dim dataReader As SqlDataReader
        'Dim mrExists As Boolean

        Dim addit As Boolean
        Dim updateit As Boolean
        Dim visitStatus As String
        addit = False
        updateit = False
        boolConsultingExists = False
        Dim j As Integer = 0
        Try
            'If Len(dictNVP.Item("Consulting Physician")) > 5 Then
            'boolConsultingExists = True ' will use this later when updating database
            'strConsultingText = dictNVP.Item("Consulting Physician")
            'consultingArray = Split(strConsultingText, "~")

            'End If
            strRoomBed = Replace(dictNVP.Item("01visit.room"), "'", "''") & Replace(dictNVP.Item("01visit.bed"), "'", "''")

            If Len(strRoomBed) = 0 Then strRoomBed = "None"
            visitStatus = "NN"
            'If dictNVP.Item("Event Type Code") = "A01" Then visitStatus = "IP"
            'If dictNVP.Item("Event Type Code") = "A03" Then visitStatus = "OEC"
            'If dictNVP.Item("Event Type Code") = "A04" Then visitStatus = "OP"
            'If dictNVP.Item("Event Type Code") = "A05" Then visitStatus = "PRE"

            '20140310 - use new status procedure
            visitStatus = calcSTARStatus(dictNVP)
            'If dictNVP.Item("Event Type Code") = "A06" Then visitStatus = "IP"
            'If dictNVP.Item("Event Type Code") = "A07" Then visitStatus = "OP"
            'If dictNVP.Item("Event Type Code") = "A08" Then visitStatus = "UP"

            '20121206 - add corpNo
            '20130515 - add star region after corpNo
            '20130522 - added department (PV2_23_3) after star_region
            sql = "Insert [01visit] "
            sql = sql & "(panum, mrnum, corpNo, star_region, department, AdminDate, room, admSource, patStatus, status, dtChanged, "
            sql = sql & "OECLimit, OECExpire, currmantype, "

            '20121115 - add attending phys during McKesson Testing
            sql = sql & "patType, patClass, accDate, diagnosis, physAdmit, physRefer, physAttend, "

            sql = sql & "physPrimary, ARO, advanceDir, primaryPlan, allergies, height, "
            sql = sql & "weight, hospServ, "

            If boolConsultingExists Then
                sql = sql & "consult1, consult2, consult3, consult4, consult5, consult6, consult7, consult8, "
                sql = sql & "consult9, consult10, "
            End If
            '20090920 - add servicing facility if available
            sql = sql & "servFacility, "
            sql = sql & "updated) "

            sql = sql & "VALUES ("
            sql = sql & "'" & gblORIGINAL_PA_NUMBER & "', "
            sql = sql & CLng(dictNVP.Item("mrnum")) & ", "

            '20121206 - add corpNo
            sql = sql & gblCorporateNumber & ", "
            insertString(dictNVP("Sending Facility"))
            '20130522 - add department here
            insertString(dictNVP("department"))

            sql = sql & "'" & ConvertDate(dictNVP.Item("01visit.AdminDate")) & "', "
            insertString(strRoomBed)

            '20080110: added Admit Source to 01visit table
            insertString(dictNVP("Admit Source"))

            '20130620 - get patstatus from calcStatus function
            '20140317 - changed to calcPatStatus routine
            insertString(calcPatStatus(dictNVP))

            '20140317
            'If dictNVP.Item("Hospital Service") = "OEC" And dictNVP.Item("Event Type Code") = "A04" Then
            'sql = sql & "'OEC', '" & ConvertDate(dictNVP.Item("01visit.AdminDate")) & "', " & "23, '"
            'sql = sql & DateAdd("h", 23, CDate(ConvertDate(dictNVP.Item("01visit.AdminDate")))) & "', "
            'Else
            sql = sql & "'" & visitStatus & "', " & "null, null, null, "
            'End If
            insertString("CM")



            insertString(dictNVP.Item("Patient Type"))
            insertString(dictNVP.Item("Patient Class"))

            If Len(dictNVP.Item("Accident Date/Time")) > 4 Then
                sql = sql & "'" & ConvertDate(dictNVP.Item("Accident Date/Time")) & "', "
            Else
                sql = sql & "NULL, "
            End If

            insertString(dictNVP("02diag.diagnosis"))
            insertNumber(dictNVP.Item("Admitting.patPhysNum"))
            insertNumber(dictNVP.Item("Referring.patPhysNum"))
            '20121115 - add attending phys during McKesson Testing
            insertNumber(dictNVP.Item("AttendingPhysicianID"))

            insertNumber(dictNVP.Item("PCPNo"))

            If dictNVP.Item("AROFieldName") = "ARO" Then
                sql = sql & "'" & dictNVP.Item("ARODataField") & "', "
            Else
                sql = sql & "NULL, "
            End If
            If Len(dictNVP.Item("AdvDirective")) >= 6 Then
                sql = sql & "'" & Left$(dictNVP.Item("AdvDirective"), 6) & "', "
            Else
                sql = sql & "NULL, "
            End If


            For i = 1 To gblInsCount

                If i = 1 Then
                    tempstr = ""
                End If
                If i > 1 Then
                    tempstr = "_000" & i
                End If

                If dictNVP.Item("COBPriority" & tempstr) = "1" Then
                    '20130522 - use the star version which includes iplancode2
                    tempPlanCode = Replace(dictNVP.Item("iplancode2" & tempstr), "'", "''") & Replace(dictNVP.Item("iplancode" & tempstr), "'", "''")

                End If
            Next
            If tempPlanCode <> "" Then
                sql = sql & "'" & tempPlanCode & "', "
            Else
                '20130526 - changed to null from 'NNN'
                sql = sql & "null, "
            End If
            '=================================================================================================================
            insertString(dictNVP("Allergy Description"))
            insertString(dictNVP("patient_height"))
            insertString(dictNVP("patient_weight"))
            insertString(dictNVP("Hospital Service"))



            If boolConsultingExists Then
                For j = 0 To UBound(consultingArray)
                    If j < 10 Then
                        insertNumber(consultingArray(j))
                    End If
                Next

                For j = (UBound(consultingArray) + 1) To 9
                    sql = sql & "NULL, "
                Next
            End If
            '20090920 ' add servicing facility
            insertString(dictNVP("Servicing Facility"))
            '20090920 - end

            sql = sql & "'" & DateTime.Now & "') "

            updatecommand.CommandText = sql
            myConnection.Open()
            updatecommand.ExecuteNonQuery()
            myConnection.Close()

        Catch ex As Exception
            functionError = True
            gblLogString = gblLogString & "Insert Visit Error" & vbCrLf
            gblLogString = gblLogString & ex.Message & vbCrLf
            'LogFile.Close()
            Exit Sub
        End Try

    End Sub
    Public Sub insertString(ByVal theString As String)
        '20130508 - modified for double quotes
        Try
            'If theString <> "" Then
            'sql = sql & "'" & Replace(theString, "'", "''") & "', "
            'Else
            'sql = sql & "NULL, "
            'End If

            Select Case theString
                Case ""
                    sql = sql & "NULL, "
                Case """"""
                    sql = sql & "NULL, "
                Case Else
                    sql = sql & "'" & Replace(theString, "'", "''") & " ', "


            End Select
        Catch ex As Exception
            functionError = True
            gblLogString = gblLogString & "Insert String Error" & vbCrLf
            gblLogString = gblLogString & ex.Message & vbCrLf
            'LogFile.Close()
            Exit Sub
        End Try
    End Sub
    Public Function STARupdateString(ByVal strVariableName As String, ByVal strVariableValue As String) As String
        STARupdateString = ""

        If Trim(strVariableValue) = """""" Or Trim(strVariableValue) = "" Then
            STARupdateString = ", " & strVariableName & " = NULL "
        Else
            STARupdateString = ", " & strVariableName & " = '" & Replace(strVariableValue, "'", "''") & "'"
        End If
    End Function
    Public Sub insertLastString(ByVal theString As String)
        '20130508 - changed process for double quotes
        Try
            'If theString <> "" Then
            'sql = sql & "'" & Replace(theString, "'", "''") & "') "
            'Else
            'sql = sql & "NULL) "
            'End If
            Select Case theString
                Case """"""
                    sql = sql & "NULL) "
                Case ""
                    sql = sql & "NULL) "
                Case Else
                    sql = sql & "'" & Replace(theString, "'", "''") & "') "
            End Select


        Catch ex As Exception
            functionError = True
            gblLogString = gblLogString & "Insert Last String Error" & vbCrLf
            gblLogString = gblLogString & ex.Message & vbCrLf
            'LogFile.Close()
            Exit Sub
        End Try
    End Sub

    Public Sub insertNumber(ByVal theString As String)
        '20130508 - mod the handle double quotes
        Try
            Select Case theString
                Case ""
                    sql = sql & "NULL, "
                Case """"""
                    sql = sql & "NULL, "
                Case Else
                    sql = sql & "'" & Replace(theString, "'", "''") & " ', "

            End Select
        Catch ex As Exception
            functionError = True
            gblLogString = gblLogString & "Insert Number Error" & vbCrLf
            gblLogString = gblLogString & ex.Message & vbCrLf
            'LogFile.Close()
            Exit Sub
        End Try
    End Sub


    Public Sub addContact(ByVal dictNVP As Hashtable)
        Dim myConnection As New SqlConnection(connectionString)
        Dim objCommand As New SqlCommand
        Dim updatecommand As New SqlCommand
        Dim dataReader As SqlDataReader
        'Dim theID As String
        Dim theDescriptor As String
        Dim paNumExists As Boolean

        Dim addit As Boolean
        Dim updateit As Boolean
        Dim contactExists As Boolean
        Dim strSQL As String
        Dim sql As String
        updatecommand.Connection = myConnection

        objCommand.Connection = myConnection
        Try
            contactExists = False
            paNumExists = False

            theDescriptor = dictNVP.Item("nok.code")
            addit = False
            updateit = False

            If dictNVP.Item("Event Type Code") = "A01" Then addit = True
            If dictNVP.Item("Event Type Code") = "A04" Then addit = True
            If dictNVP.Item("Event Type Code") = "A05" Then addit = False

            If Len(dictNVP.Item("03contact.lastName")) > 0 Then
                contactExists = True
            End If

            sql = "select panum from [03contact] where panum = '" & gblORIGINAL_PA_NUMBER & "'"
            objCommand.CommandText = sql
            myConnection.Open()
            dataReader = objCommand.ExecuteReader()
            If dataReader.HasRows Then
                'record found; contact exists

                paNumExists = True
            Else
                'contact does not exist
                paNumExists = False
            End If

            myConnection.Close()
            dataReader.Close()

            If ((addit) And (contactExists) And (Not paNumExists)) Then


                strSQL = "INSERT [03contact] (panum, name, relation, ph1, ph2, updated) "
                strSQL = strSQL & "VALUES ("
                strSQL = strSQL & "'" & gblORIGINAL_PA_NUMBER & "', "

                strSQL = strSQL & "'" & Replace(dictNVP.Item("03contact.firstName"), "'", "''") & " " & Replace(dictNVP.Item("03contact.lastName"), "'", "''") & "', "

                strSQL = strSQL & "'" & theDescriptor & "', "

                strSQL = strSQL & "'" & dictNVP.Item("nok.phone") & "', "
                strSQL = strSQL & "'" & dictNVP.Item("nok.businessphone") & "', "
                '========================================================================================
                strSQL = strSQL & "'" & DateTime.Now & "') "
                updatecommand.CommandText = strSQL
                myConnection.Open()
                updatecommand.ExecuteNonQuery()
                myConnection.Close()
            End If

            'LogFile.WriteLine("Contact Process Complete")
        Catch ex As Exception
            functionError = True
            gblLogString = gblLogString & "Contact Error" & vbCrLf
            gblLogString = gblLogString & ex.Message & vbCrLf
            'LogFile.Close()
            Exit Sub
        End Try
    End Sub
    Public Sub processError(ByVal dictNVP As Hashtable)
        Dim myConnection As New SqlConnection(connectionString)
        Dim objCommand As New SqlCommand
        objCommand.Connection = myConnection
        functionError = False
        dbError = False
        continueProcessing = True
        orphanFound = False
        Try
            Dim dataReader As SqlDataReader
            Dim checkThis As Boolean = True
            Dim sql As String = ""
            If dictNVP.Item("Event Type Code") = "A01" Then checkThis = False
            If dictNVP.Item("Event Type Code") = "A04" Then checkThis = False
            If dictNVP.Item("Event Type Code") = "A05" Then checkThis = False
            If dictNVP.Item("Event Type Code") = "A31" Then checkThis = False

            If checkThis And dictNVP.Item("panum") <> "" Then
                sql = "select panum from [01visit] where panum = '" & extractPanum(dictNVP.Item("panum")) & "'"
                objCommand.CommandText = sql
                myConnection.Open()
                dataReader = objCommand.ExecuteReader()
                If dataReader.HasRows Then
                    'do nothing the panum exists
                Else
                    'can't find the panum so send to orpan directory
                    '20150717 - don't process orphans leave orphan found as false
                    orphanFound = False ' True
                    continueProcessing = False
                    gblLogString = gblLogString & CStr(DateTime.Now) & " - Orphan Found. Panum = " & extractPanum(dictNVP.Item("panum")) & vbCrLf
                End If
                dataReader.Close()

            End If

        Catch ex As Exception
            continueProcessing = False
            If Err.Number = 5 Then
                dbError = True
                functionError = False
            Else
                functionError = True
            End If

            gblLogString = gblLogString & "Connection Error: " & Err.Number & vbCrLf
            gblLogString = gblLogString & ex.Message & vbCrLf
        Finally
            myConnection.Close()
        End Try
    End Sub

    Public Sub writeToLog2(ByVal logText As String, ByVal eventType As Integer)
        Dim myLog As New EventLog()
        Try
            ' check for the existence of the log that the user wants to create.
            ' Create the source, if it does not already exist.
            If Not EventLog.SourceExists("STAR_McareFeed") Then
                EventLog.CreateEventSource("STAR_McareFeed", "STAR_McareFeed")
            End If

            ' Create an EventLog instance and assign its source.

            myLog.Source = "STAR_McareFeed"

            ' Write an informational entry to the event log.
            If eventType = 1 Then
                myLog.WriteEntry(logText, EventLogEntryType.Error, 1)
            ElseIf eventType = 2 Then
                myLog.WriteEntry(logText, EventLogEntryType.Warning, 2)
            ElseIf eventType = 3 Then
                myLog.WriteEntry(logText, EventLogEntryType.Information, 3)
            End If

        Catch ex As Exception
            gblLogString = gblLogString + ex.Message
        Finally
            myLog.Close()
        End Try
    End Sub
    Public Sub writeTolog(ByVal strMsg As String, ByVal eventType As Integer)
        '20140205 - use a text file to log errors instead of the event log
        Dim file As System.IO.StreamWriter
        Dim tempLogFileName As String = strLogDirectory & "McareULHFeed_log.txt"
        file = My.Computer.FileSystem.OpenTextFileWriter(tempLogFileName, True)
        file.WriteLine(DateTime.Now & " : " & strMsg)
        file.Close()
    End Sub
    Public Sub checkZMI(ByVal dictNVP As Hashtable)
        '20100218 -  procedure to count multiple ZMI segments
        Dim ZMICount As Integer
        Try
            Dim myEnumerator As IDictionaryEnumerator = dictNVP.GetEnumerator()
            ZMICount = 0

            'For Each value In dictNVP.values
            'Console.WriteLine(value)

            'Next
            myEnumerator.Reset()
            While myEnumerator.MoveNext()
                If Left$(myEnumerator.Key, 5) = "ZMI_5" Then
                    gblZMICount = gblZMICount + 1
                    ZMICount = ZMICount + 1
                End If

            End While
            gblLogString = gblLogString & " ZMICount =  "
            gblLogString = gblLogString & ZMICount & vbCrLf
        Catch ex As Exception
            functionError = True
            gblLogString = gblLogString & "ZMI Enumeration Error" & vbCrLf
            gblLogString = gblLogString & ex.Message & vbCrLf
            'LogFile.Close()
            Exit Sub
        End Try
    End Sub
    Public Sub processZMI_old(ByVal dictNVP As Hashtable)
        '2018 procedure writes auditNotes field to 03Insurer table.
        Dim i As Integer = 0
        Dim tempstr As String = ""
        Dim tmpInsCode As String = ""
        Dim auditNotes As String = ""
        Dim updateIT As Boolean = False
        'Dim LogFile As StreamWriter = File.AppendText("c:\temp\Log.txt")
        Try
            Dim myConnection As New SqlConnection(connectionString)
            Dim updateCommand As New SqlCommand
            updateCommand.Connection = myConnection
            Dim result As New Text.StringBuilder

            If dictNVP.Item("Event Type Code") = "A01" Then updateIT = True
            If dictNVP.Item("Event Type Code") = "A04" Then updateIT = True
            If dictNVP.Item("Event Type Code") = "A05" Then updateIT = True
            If dictNVP.Item("Event Type Code") = "A08" Then updateIT = True

            If updateIT Then
                For i = 1 To gblZMICount
                    If i = 1 Then
                        tempstr = ""
                    End If
                    If i > 1 Then
                        tempstr = "_000" & i
                    End If

                    tmpInsCode = dictNVP.Item("ZMI_5" & tempstr)

                    'If dictNVP("ZMI_3" & tempstr) <> "" Then
                    'result.AppendLine("Priority: " + dictNVP("ZMI_3" & tempstr))
                    'End If
                    'If dictNVP("ZMI_4" & tempstr) <> "" Then
                    'result.AppendLine("Co. Name: " + dictNVP("ZMI_4" & tempstr))
                    'End If
                    If dictNVP("ZMI_5" & tempstr) <> "" Then
                        result.AppendLine("Co. Ins Code: " + dictNVP("ZMI_5" & tempstr))
                    End If

                    If dictNVP("ZMI_6" & tempstr) <> "" Then
                        result.AppendLine("Precert No.: " + dictNVP("ZMI_6" & tempstr))
                    End If
                    If dictNVP("ZMI_7" & tempstr) <> "" Then
                        result.AppendLine("Precert Contact: " + dictNVP("ZMI_7" & tempstr))
                    End If
                    If dictNVP("ZMI_8" & tempstr) <> "" Then
                        result.AppendLine("Precert Phone No. (with ext): " + dictNVP("ZMI_8" & tempstr))
                    End If
                    If dictNVP("ZMI_9" & tempstr) <> "" Then
                        result.AppendLine("Precert Fax No.: " + dictNVP("ZMI_9" & tempstr))
                    End If
                    If dictNVP("ZMI_10" & tempstr) <> "" Then
                        result.AppendLine("Services Authorized 1: " + dictNVP("ZMI_10" & tempstr))
                    End If
                    If dictNVP("ZMI_11" & tempstr) <> "" Then
                        result.AppendLine("Services Authorized 2: " + dictNVP("ZMI_11" & tempstr))
                    End If
                    If dictNVP("ZMI_12" & tempstr) <> "" Then
                        result.AppendLine("Precert Amount: " + dictNVP("ZMI_12" & tempstr))
                    End If
                    If dictNVP("ZMI_13" & tempstr) <> "" Then
                        result.AppendLine("Benefit Per: " + dictNVP("ZMI_13" & tempstr))
                    End If
                    If dictNVP("ZMI_14" & tempstr) <> "" Then
                        result.AppendLine("Benefit Effective Date: " + dictNVP("ZMI_14" & tempstr))
                    End If
                    If dictNVP("ZMI_15" & tempstr) <> "" Then
                        result.AppendLine("Benefit Ded: " + dictNVP("ZMI_15" & tempstr))
                    End If
                    If dictNVP("ZMI_16" & tempstr) <> "" Then
                        result.AppendLine("Benefit Phone No.: " + dictNVP("ZMI_16" & tempstr))
                    End If
                    If dictNVP("ZMI_17" & tempstr) <> "" Then
                        result.AppendLine("Benefit Met: " + dictNVP("ZMI_17" & tempstr))
                    End If
                    If dictNVP("ZMI_18" & tempstr) <> "" Then
                        result.AppendLine("Benefit Copay: " + dictNVP("ZMI_18" & tempstr))
                    End If
                    If dictNVP("ZMI_19" & tempstr) <> "" Then
                        result.AppendLine("Benefit Out of Pocket: " + dictNVP("ZMI_19" & tempstr))
                    End If
                    If dictNVP("ZMI_20" & tempstr) <> "" Then
                        result.AppendLine("Coverage 1: " + dictNVP("ZMI_20" & tempstr))
                    End If
                    If dictNVP("ZMI_21" & tempstr) <> "" Then
                        result.AppendLine("Coverage 2: " + dictNVP("ZMI_21" & tempstr))
                    End If
                    If dictNVP("ZMI_22" & tempstr) <> "" Then
                        result.AppendLine("Verified by: " + dictNVP("ZMI_22" & tempstr))
                    End If
                    If dictNVP("ZMI_23" & tempstr) <> "" Then
                        result.AppendLine("Verify Date: " + dictNVP("ZMI_23" & tempstr))
                    End If
                    If dictNVP("ZMI_24" & tempstr) <> "" Then
                        result.AppendLine("Limitation 1: " + dictNVP("ZMI_24" & tempstr))
                    End If
                    If dictNVP("ZMI_25" & tempstr) <> "" Then
                        result.AppendLine("Limitation 2: " + dictNVP("ZMI_25" & tempstr))
                    End If
                    If dictNVP("ZMI_26" & tempstr) <> "" Then
                        result.AppendLine("Limitation 3: " + dictNVP("ZMI_26" & tempstr))
                    End If
                    If dictNVP("ZMI_27" & tempstr) <> "" Then
                        result.AppendLine("Limitation 4: " + dictNVP("ZMI_27" & tempstr))
                    End If
                    If dictNVP("ZMI_28" & tempstr) <> "" Then
                        result.AppendLine("Benefit Comment 1: " + dictNVP("ZMI_28" & tempstr))
                    End If
                    If dictNVP("ZMI_29" & tempstr) <> "" Then
                        result.AppendLine("Benefit Comment 2: " + dictNVP("ZMI_29" & tempstr))
                    End If
                    If dictNVP("ZMI_30" & tempstr) <> "" Then
                        result.AppendLine("Benefit Comment 3: " + dictNVP("ZMI_30" & tempstr))
                    End If
                    If dictNVP("ZMI_31" & tempstr) <> "" Then
                        result.AppendLine("Benefit Comment 4: " + dictNVP("ZMI_31" & tempstr))
                    End If
                    If dictNVP("ZMI_32" & tempstr) <> "" Then
                        result.AppendLine("Benefit Comment 5: " + dictNVP("ZMI_32" & tempstr))
                    End If
                    If dictNVP("ZMI_33" & tempstr) <> "" Then
                        result.AppendLine("Benefit Comment 6: " + dictNVP("ZMI_33" & tempstr))
                    End If
                    If dictNVP("ZMI_34" & tempstr) <> "" Then
                        result.AppendLine("Benefit Comment 7: " + dictNVP("ZMI_34" & tempstr))
                    End If
                    If dictNVP("ZMI_35" & tempstr) <> "" Then
                        result.AppendLine("Benefit Comment 8: " + dictNVP("ZMI_35" & tempstr))
                    End If
                    If dictNVP("ZMI_36" & tempstr) <> "" Then
                        result.AppendLine("Benefit Comment 9: " + dictNVP("ZMI_36" & tempstr))
                    End If
                    If dictNVP("ZMI_37" & tempstr) <> "" Then
                        result.AppendLine("Benefit Comment 10: " + dictNVP("ZMI_37" & tempstr))
                    End If
                    If dictNVP("ZMI_38" & tempstr) <> "" Then
                        result.AppendLine("Benefit Comment 11: " + dictNVP("ZMI_38" & tempstr))
                    End If
                    'Set the stringbuilder to a string
                    auditNotes = result.ToString()

                    'LogFile.WriteLine(auditNotes)

                    sql = "UPDATE [03Insurer] "
                    sql = sql & "SET updated = '" & DateTime.Now & "'"
                    sql = sql & ", auditNotes = '" & auditNotes & "'"
                    '20170804 - Keep track of insurance changes
                    sql = sql & ", lastupdatedStar = '" & DateTime.Now & "' "

                    sql = sql & " Where panum = '" & gblORIGINAL_PA_NUMBER & "'"
                    sql = sql & " AND iplanCode = '" & tmpInsCode & "'"
                    updateCommand.CommandText = sql
                    myConnection.Open()
                    updateCommand.ExecuteNonQuery()
                    myConnection.Close()

                    'LogFile.WriteLine(sql)

                    result.Length = 0 'clear the string Builder by setting its length to zero.

                    'LogFile.Close()
                Next 'gblZMICount
            End If 'If updateIT

        Catch ex As Exception
            functionError = True
            gblLogString = gblLogString & "Process ZMI Error" & vbCrLf
            gblLogString = gblLogString & ex.Message & vbCrLf
            'LogFile.Close()
            Exit Sub

        End Try
    End Sub
    Public Sub processZMI(ByVal dictNVP As Hashtable)
        '2018 procedure writes auditNotes field to 03Insurer table.
        Dim i As Integer = 0
        Dim sql As String = ""
        Dim tempstr As String = ""
        Dim updateIT As Boolean = False
        Dim recordExists As Boolean = False

        Try
            Dim myConnection As New SqlConnection(connectionString)
            Dim objCommand As New SqlCommand
            objCommand.Connection = myConnection
            Dim dataReader As SqlDataReader

            If dictNVP.Item("Event Type Code") = "A01" Then updateIT = True
            If dictNVP.Item("Event Type Code") = "A04" Then updateIT = True
            If dictNVP.Item("Event Type Code") = "A05" Then updateIT = True
            If dictNVP.Item("Event Type Code") = "A08" Then updateIT = True

            If updateIT And gblZMICount > 0 Then
                For i = 1 To gblZMICount
                    If i = 1 Then
                        tempstr = ""
                    End If
                    If i > 1 Then
                        tempstr = "_000" & i
                    End If

                    sql = "select * from [03insurerSupplement] "
                    sql = sql & "where panum = '" & gblORIGINAL_PA_NUMBER & "' and iplancode = '" & dictNVP.Item("ZMI_5" & tempstr) & "'"
                    objCommand.CommandText = sql
                    myConnection.Open()
                    dataReader = objCommand.ExecuteReader()

                    If dataReader.HasRows Then
                        recordExists = True
                    Else
                        recordExists = False

                    End If


                    dataReader.Close()
                    myConnection.Close()
                    If recordExists Then
                        Call ZMIUpdate(dictNVP, tempstr)
                    Else
                        Call ZMIInsert(dictNVP, tempstr)
                    End If



                Next 'gblZMICount
            End If 'If updateIT

        Catch ex As Exception
            functionError = True
            gblLogString = gblLogString & "Process ZMI2 Error" & vbCrLf
            gblLogString = gblLogString & ex.Message & vbCrLf
            'LogFile.Close()
            Exit Sub

        End Try
    End Sub
    Public Sub ZMIInsert_old(ByVal dictNVP As Hashtable, ByVal tempStr As String)
        'Dim sql As String = ""
        Try
            Dim myConnection As New SqlConnection(connectionString)
            Dim objCommand As New SqlCommand
            objCommand.Connection = myConnection


            sql = "Insert [03InsurerSupplement] "
            sql = sql & "(panum, iplancode, precertNo, precertContact, precertPhone, precertFax, serviceAuth1, serviceAuth2, "
            sql = sql & "precertAmount, benefitPerson, benefitEffectiveDate, benefitDeduction, benefitPhone, benefitMet, "
            sql = sql & "benefitCoPay, benefitOutOfPocket, coverage1, coverage2, verifiedBy, verifyDate, limitations1, limitations2, "
            sql = sql & "limitations3, limitations4, "
            sql = sql & "comments1, comments2, comments3, comments4, comments5, comments6, comments7, comments8, "
            sql = sql & "comments9, comments10, "
            sql = sql & "comments11) "
            sql = sql & "VALUES ("
            insertString(gblORIGINAL_PA_NUMBER)
            insertString(dictNVP("ZMI_5" & tempStr)) 'iplancode
            insertString(dictNVP("ZMI_6" & tempStr)) ' precertNo
            insertString(dictNVP("ZMI_7" & tempStr)) ' precertContact
            insertString(dictNVP("ZMI_8" & tempStr)) ' precertPhone
            insertString(dictNVP("ZMI_9" & tempStr)) ' precertFax
            insertString(dictNVP("ZMI_10" & tempStr)) ' serviceAuth1
            insertString(dictNVP("ZMI_11" & tempStr)) ' serviceAuth2
            insertString(dictNVP("ZMI_12" & tempStr)) ' precertAmount
            insertString(dictNVP("ZMI_13" & tempStr)) ' benefitPerson
            insertString(dictNVP("ZMI_14" & tempStr)) ' benefitEffectiveDate
            insertString(dictNVP("ZMI_15" & tempStr)) ' benefitDeduction
            insertString(dictNVP("ZMI_16" & tempStr)) ' benefitPhone
            insertString(dictNVP("ZMI_17" & tempStr)) ' benefitMet
            insertString(dictNVP("ZMI_18" & tempStr)) ' benefitCoPay
            insertString(dictNVP("ZMI_19" & tempStr)) ' benefitOutOfPocket
            insertString(dictNVP("ZMI_20" & tempStr)) ' coverage1
            insertString(dictNVP("ZMI_21" & tempStr)) ' coverage2
            insertString(dictNVP("ZMI_22" & tempStr)) ' verifiedBy
            insertString(dictNVP("ZMI_23" & tempStr)) ' verifyDate
            insertString(dictNVP("ZMI_24" & tempStr)) ' limitations1
            insertString(dictNVP("ZMI_25" & tempStr)) ' limitations2
            insertString(dictNVP("ZMI_26" & tempStr)) ' limitations3
            insertString(dictNVP("ZMI_27" & tempStr)) ' limitations4
            insertString(dictNVP("ZMI_28" & tempStr)) ' comments1
            insertString(dictNVP("ZMI_29" & tempStr)) ' comments2
            insertString(dictNVP("ZMI_30" & tempStr)) ' comments3
            insertString(dictNVP("ZMI_31" & tempStr)) ' comments4
            insertString(dictNVP("ZMI_32" & tempStr)) ' comments5
            insertString(dictNVP("ZMI_33" & tempStr)) ' comments6
            insertString(dictNVP("ZMI_34" & tempStr)) ' comments7
            insertString(dictNVP("ZMI_35" & tempStr)) ' comments8
            insertString(dictNVP("ZMI_36" & tempStr)) ' comments9
            insertString(dictNVP("ZMI_37" & tempStr)) ' comments10

            insertLastString(dictNVP("ZMI_38" & tempStr))  'comments11
            objCommand.CommandText = sql
            myConnection.Open()

            objCommand.ExecuteNonQuery()
            myConnection.Close()


        Catch ex As Exception
            functionError = True
            gblLogString = gblLogString & "Process ZMI Insert Error" & vbCrLf
            gblLogString = gblLogString & ex.Message & vbCrLf
            'LogFile.Close()
            Exit Sub

        End Try
    End Sub
    Public Sub ZMIUpdate_old(ByVal dictNVP As Hashtable, ByVal tempStr As String)
        Dim sql As String = ""
        Try
            Dim myConnection As New SqlConnection(connectionString)
            Dim objCommand As New SqlCommand
            objCommand.Connection = myConnection

            sql = "UPDATE [03InsurerSupplement] "
            sql = sql & "SET updated = '" & DateTime.Now & "'"
            sql = sql & ", precertNo = '" & Replace(dictNVP("ZMI_6" & tempStr), "'", "''") & "'"
            sql = sql & ", precertContact = '" & Replace(dictNVP("ZMI_7" & tempStr), "'", "''") & "'"
            sql = sql & ", precertPhone = '" & Replace(dictNVP("ZMI_8" & tempStr), "'", "''") & "'"
            sql = sql & ", precertFax = '" & Replace(dictNVP("ZMI_9" & tempStr), "'", "''") & "'"
            sql = sql & ", serviceAuth1 = '" & Replace(dictNVP("ZMI_10" & tempStr), "'", "''") & "'"
            sql = sql & ", serviceAuth2 = '" & Replace(dictNVP("ZMI_11" & tempStr), "'", "''") & "'"
            sql = sql & ", precertAmount = '" & Replace(dictNVP("ZMI_12" & tempStr), "'", "''") & "'"
            sql = sql & ", benefitPerson = '" & Replace(dictNVP("ZMI_13" & tempStr), "'", "''") & "'"
            sql = sql & ", benefitEffectiveDate = '" & Replace(dictNVP("ZMI_14" & tempStr), "'", "''") & "'"
            sql = sql & ", benefitDeduction = '" & Replace(dictNVP("ZMI_15" & tempStr), "'", "''") & "'"
            sql = sql & ", benefitPhone = '" & Replace(dictNVP("ZMI_16" & tempStr), "'", "''") & "'"
            sql = sql & ", BenefitMet = '" & Replace(dictNVP("ZMI_17" & tempStr), "'", "''") & "'"
            sql = sql & ", benefitCoPay = '" & Replace(dictNVP("ZMI_18" & tempStr), "'", "''") & "'"
            sql = sql & ", benefitOutOfPocket = '" & Replace(dictNVP("ZMI_19" & tempStr), "'", "''") & "'"
            sql = sql & ", coverage1 = '" & Replace(dictNVP("ZMI_20" & tempStr), "'", "''") & "'"
            sql = sql & ", coverage2 = '" & Replace(dictNVP("ZMI_21" & tempStr), "'", "''") & "'"
            sql = sql & ", verifiedBy = '" & Replace(dictNVP("ZMI_22" & tempStr), "'", "''") & "'"
            sql = sql & ", verifyDate = '" & Replace(dictNVP("ZMI_23" & tempStr), "'", "''") & "'"
            sql = sql & ", limitations1 = '" & Replace(dictNVP("ZMI_24" & tempStr), "'", "''") & "'"
            sql = sql & ", limitations2 = '" & Replace(dictNVP("ZMI_25" & tempStr), "'", "''") & "'"
            sql = sql & ", limitations3 = '" & Replace(dictNVP("ZMI_26" & tempStr), "'", "''") & "'"
            sql = sql & ", limitations4 = '" & Replace(dictNVP("ZMI_27" & tempStr), "'", "''") & "'"
            sql = sql & ", comments1 = '" & Replace(dictNVP("ZMI_28" & tempStr), "'", "''") & "'"
            sql = sql & ", comments2 = '" & Replace(dictNVP("ZMI_29" & tempStr), "'", "''") & "'"
            sql = sql & ", comments3 = '" & Replace(dictNVP("ZMI_30" & tempStr), "'", "''") & "'"
            sql = sql & ", comments4 = '" & Replace(dictNVP("ZMI_31" & tempStr), "'", "''") & "'"
            sql = sql & ", comments5 = '" & Replace(dictNVP("ZMI_32" & tempStr), "'", "''") & "'"
            sql = sql & ", comments6 = '" & Replace(dictNVP("ZMI_33" & tempStr), "'", "''") & "'"
            sql = sql & ", comments7 = '" & Replace(dictNVP("ZMI_34" & tempStr), "'", "''") & "'"
            sql = sql & ", comments8 = '" & Replace(dictNVP("ZMI_35" & tempStr), "'", "''") & "'"
            sql = sql & ", comments9 = '" & Replace(dictNVP("ZMI_36" & tempStr), "'", "''") & "'"
            sql = sql & ", comments10 = '" & Replace(dictNVP("ZMI_37" & tempStr), "'", "''") & "'"
            sql = sql & ", comments11 = '" & Replace(dictNVP("ZMI_38" & tempStr), "'", "''") & "'"


            sql = sql & " Where panum = '" & gblORIGINAL_PA_NUMBER & "'"
            sql = sql & " and iplancode = '" & dictNVP("ZMI_5" & tempStr) & "'"

            objCommand.CommandText = sql
            myConnection.Open()
            objCommand.ExecuteNonQuery()
            myConnection.Close()



        Catch ex As Exception
            functionError = True
            gblLogString = gblLogString & "ZMIUpdate Error" & vbCrLf
            gblLogString = gblLogString & ex.Message & vbCrLf
            'LogFile.Close()
            Exit Sub

        End Try
    End Sub
    Public Sub ZMIInsert(ByVal dictNVP As Hashtable, ByVal tempStr As String)
        'Dim sql As String = ""
        Try
            Dim myConnection As New SqlConnection(connectionString)
            Dim objCommand As New SqlCommand
            objCommand.Connection = myConnection


            sql = "Insert [03InsurerSupplement] "
            sql = sql & "(panum, iplancode, precertPhone, coverage1, coverage2) "
            'sql = sql & "precertAmount, benefitPerson, benefitEffectiveDate, benefitDeduction, benefitPhone, benefitMet, "
            'sql = sql & "benefitCoPay, benefitOutOfPocket, coverage1, coverage2, verifiedBy, verifyDate, limitations1, limitations2, "
            'sql = sql & "limitations3, limitations4, "
            'sql = sql & "comments1, comments2, comments3, comments4, comments5, comments6, comments7, comments8, "
            'sql = sql & "comments9, comments10, "
            'sql = sql & "comments11) "
            sql = sql & "VALUES ("
            insertString(gblORIGINAL_PA_NUMBER)
            insertString(dictNVP("ZMI_5" & tempStr))
            insertString(dictNVP("ZMI_8" & tempStr))
            insertString(dictNVP("ZMI_20" & tempStr))
            insertLastString(dictNVP("ZMI_21" & tempStr))
            
            objCommand.CommandText = sql
            myConnection.Open()

            objCommand.ExecuteNonQuery()
            myConnection.Close()


        Catch ex As Exception
            functionError = True
            gblLogString = gblLogString & "Process ZMI Insert Error" & vbCrLf
            gblLogString = gblLogString & ex.Message & vbCrLf
            'LogFile.Close()
            Exit Sub

        End Try
    End Sub
    Public Sub ZMIUpdate(ByVal dictNVP As Hashtable, ByVal tempStr As String)
        Dim sql As String = ""
        Try
            Dim myConnection As New SqlConnection(connectionString)
            Dim objCommand As New SqlCommand
            objCommand.Connection = myConnection

            sql = "UPDATE [03InsurerSupplement] "
            sql = sql & "SET updated = '" & DateTime.Now & "'"

            sql = sql & ", precertPhone = '" & Replace(dictNVP("ZMI_8" & tempStr), "'", "''") & "'"
            
            sql = sql & ", coverage1 = '" & Replace(dictNVP("ZMI_20" & tempStr), "'", "''") & "'"
            sql = sql & ", coverage2 = '" & Replace(dictNVP("ZMI_21" & tempStr), "'", "''") & "'"
            
            sql = sql & " Where panum = '" & gblORIGINAL_PA_NUMBER & "'"
            sql = sql & " and iplancode = '" & dictNVP("ZMI_5" & tempStr) & "'"

            objCommand.CommandText = sql
            myConnection.Open()
            objCommand.ExecuteNonQuery()
            myConnection.Close()



        Catch ex As Exception
            functionError = True
            gblLogString = gblLogString & "ZMIUpdate Error" & vbCrLf
            gblLogString = gblLogString & ex.Message & vbCrLf
            'LogFile.Close()
            Exit Sub

        End Try
    End Sub
    Public Sub processAL1(ByVal dictNVP As Hashtable)
        '20140215 For A31. Note, ther is no panum in the A31 so we need to use the mrnum
        '20140321 - added use of extractMrnum to this function only.
        '20140915 - modified search criteria for processAL1
        '20160916 - capture all AL1 data.
        'Dim A31connectionString As String = "server=10.48.242.249,1433;database=PatientGlobal;uid=sysmax;pwd=Condor!"
        'Dim A31connectionString As String = "server=10.48.64.5\sqlexpress;Initial Catalog=PatientGlobal;User Id=sysmax;Password=Condor!"
        'connectionString = "server=HPLAPTOP;database=STAR_ITW;uid=sa;pwd=b436328"
        Dim A31connectionString As String = conIniFile.GetString("Strings", "MCAREULHAL1", "(none)")

        Dim myConnection As New SqlConnection(A31connectionString)
        Dim objCommand As New SqlCommand
        Dim updatecommand As New SqlCommand
        updatecommand.Connection = myConnection
        objCommand.Connection = myConnection


        Dim boolProsessThis As Boolean
        Dim tempstr As String
        Dim star_region As String = "" '20140514 added to capture region in global table also capture global corporate number
        Try

            Dim i As Integer
            boolProsessThis = False
            If dictNVP.Item("TriggerEventID") = "A31" Then boolProsessThis = True
            '21040514
            star_region = UCase(dictNVP("Sending Facility"))

            If boolProsessThis Then
                updatecommand.Connection = myConnection
                objCommand.Connection = myConnection
                sql = "delete from [Allergies] "
                sql = sql & "where CorporateNumber = " & gblCorporateNumber ' extractMrnum(dictNVP.Item("mrnum"))
                updatecommand.CommandText = sql
                myConnection.Open()
                updatecommand.ExecuteNonQuery()
                myConnection.Close()
                updatecommand.Connection = myConnection
                objCommand.Connection = myConnection

                tempstr = ""
                sql = ""
                i = 0

                For i = 1 To gblAL1Count
                    sql = ""
                    If i = 1 Then
                        tempstr = ""
                    End If

                    '20140918 - accept more than 9 AL1 segments
                    If i > 1 And i < 10 Then
                        tempstr = "_000" & i
                    End If

                    If i >= 10 And i < 100 Then
                        tempstr = "_00" & i
                    End If

                    'If Not IsDBNull(dictNVP("Allergy Code ID" & tempstr)) Then '20140915
                    ''If Trim(dictNVP("Allergy Code ID" & tempstr)) <> "" Then '20160916 - capture all AL1 data.
                    '20140514 - add star_region and gblCorporateNumber
                    sql = "Insert [Allergies] "
                    sql = sql & "(mrnum, region, CorporateNumber, type, code_id, description, coding_system, Severity, Reaction, IDDate, "
                    sql = sql & "added) "

                    sql = sql & "VALUES ("
                    sql = sql & extractMrnum(dictNVP.Item("mrnum")) & ", "
                    insertString(star_region)
                    sql = sql & gblCorporateNumber & ", "

                    insertString(dictNVP.Item("Allergy Type" & tempstr))
                    insertString(dictNVP("Allergy Code ID" & tempstr))
                    insertString(dictNVP.Item("Allergy Description" & tempstr))
                    insertString(dictNVP.Item("Allergy Coding System" & tempstr))

                    '20140303 - added severity, reaction and IDDate
                    insertString(dictNVP.Item("Severity" & tempstr))
                    insertString(dictNVP.Item("Reaction" & tempstr))
                    sql = sql & "'" & ConvertDate(dictNVP("IDDate" & tempstr)) & "', "

                    sql = sql & "'" & DateTime.Now & "') "

                    updatecommand.CommandText = sql
                    myConnection.Open()
                    updatecommand.ExecuteNonQuery()
                    myConnection.Close()
                    'End If '20160916 - capture all AL1 data.
                Next 'For i = 1 To gblAL1Count
            End If 'If boolProsessThis
        Catch ex As Exception
            functionError = True
            gblLogString = gblLogString & "AL1 Process Error" & vbCrLf
            gblLogString = gblLogString & ex.Message & vbCrLf
            'LogFile.Close()
            Exit Sub
        End Try
    End Sub

    Public Sub processAL1_Old(ByVal dictNVP As Hashtable)
        Dim myConnection As New SqlConnection(connectionString)
        Dim objCommand As New SqlCommand
        Dim updatecommand As New SqlCommand
        Dim boolProsessThis As Boolean
        Dim tempstr As String
        Dim varMrnum As String = ""
        '20140215 - modified routine to match STAR_ITWfeed
        '20140303 - added severity, reaction and IDDate
        '20140915 - modified search criteria for processAL1
        Try

            Dim i As Integer
            boolProsessThis = False
            'If dictNVP.Item("Event Type Code") = "A01" Then boolProsessThis = False
            'If dictNVP.Item("Event Type Code") = "A04" Then boolProsessThis = False
            'If dictNVP.Item("Event Type Code") = "A05" Then boolProsessThis = False
            'If dictNVP.Item("Event Type Code") = "A08" Then boolProsessThis = False

            If dictNVP.Item("TriggerEventID") = "A31" Then boolProsessThis = True

            If IsNumeric(Left(gblMrnum, 1)) Then
                varMrnum = gblMrnum
            Else
                varMrnum = Mid(gblMrnum, 2)
            End If

            If boolProsessThis Then
                updatecommand.Connection = myConnection
                objCommand.Connection = myConnection
                sql = "delete from [222AL1] "
                'sql = sql & "where panum = '" & gblORIGINAL_PA_NUMBER & "'"
                sql = sql & "where mrnum = " & varMrnum

                updatecommand.CommandText = sql
                myConnection.Open()
                updatecommand.ExecuteNonQuery()
                myConnection.Close()
                updatecommand.Connection = myConnection
                objCommand.Connection = myConnection

                tempstr = ""
                sql = ""
                i = 0

                For i = 1 To gblAL1Count
                    sql = ""
                    If i = 1 Then
                        tempstr = ""
                    End If
                    If i > 1 Then
                        tempstr = "_000" & i
                    End If
                    If Not IsDBNull(dictNVP("Allergy Code ID" & tempstr)) Then
                        sql = "Insert [222AL1] "
                        'sql = sql & "(panum, type, code_id, description, coding_system, "

                        '20140303 - added severity, reaction and IDDate
                        sql = sql & "(mrnum, type, code_id, description, coding_system, Severity, Reaction, IDDate, "
                        sql = sql & "added) "

                        sql = sql & "VALUES ("
                        'sql = sql & "'" & gblORIGINAL_PA_NUMBER & "', "
                        sql = sql & varMrnum & ", "
                        insertString(dictNVP.Item("Allergy Type" & tempstr))
                        insertString(dictNVP("Allergy Code ID" & tempstr))
                        insertString(dictNVP.Item("Allergy Description" & tempstr))
                        insertString(dictNVP.Item("Allergy Coding System" & tempstr))

                        '20140303 - added severity, reaction and IDDate
                        insertString(dictNVP.Item("Severity" & tempstr))
                        insertString(dictNVP.Item("Reaction" & tempstr))
                        sql = sql & "'" & ConvertDate(dictNVP("IDDate" & tempstr)) & "', "

                        sql = sql & "'" & DateTime.Now & "') "

                        updatecommand.CommandText = sql
                        myConnection.Open()
                        updatecommand.ExecuteNonQuery()
                        myConnection.Close()
                    End If
                Next 'For i = 1 To gblAL1Count
            End If 'If boolProsessThis
        Catch ex As Exception
            functionError = True
            gblLogString = gblLogString & "AL1_Old Process Error" & vbCrLf
            gblLogString = gblLogString & ex.Message & vbCrLf
            'LogFile.Close()
            Exit Sub
        End Try
    End Sub
    Public Function extractMrnum(ByVal varMrnum As String) As String
        extractMrnum = "0"
        If IsNumeric(Left(varMrnum, 1)) Then
            extractMrnum = varMrnum
        Else
            extractMrnum = Mid(varMrnum, 2)
        End If
    End Function

    Public Function extractPanum(ByVal varPanum As String) As String
        extractPanum = "0"
        If IsNumeric(Left(varPanum, 1)) Then
            extractPanum = varPanum
        Else
            extractPanum = Mid(varPanum, 2)
        End If
    End Function


    Public Function extractCorpNo(ByVal PID2 As String) As String
        extractCorpNo = "0"
        If IsNumeric(Left(PID2, 1)) Then
            extractCorpNo = PID2
        Else
            extractCorpNo = Mid(PID2, 2)
        End If
    End Function
    Public Function extractCorpNoCerner(ByVal pidString As String) As String
        '20140321 - use cerner version 
        Dim pidArray() As String
        Dim pidItemArray() As String
        Dim tempStr As String = ""
        Dim testData As String = ""
        Dim J As Integer = 0

        Try
            If Left(pidString, 1) = "~" Then
                pidString = Mid(pidString, 2)
            End If
            extractCorpNoCerner = "0"
            pidArray = Split(pidString, "~")
            For J = 0 To UBound(pidArray)
                tempStr = pidArray(J)
                tempStr = Trim(tempStr)
                'testData = Mid$(tempStr, Len(tempStr) - 1, 2)
                pidItemArray = Split(tempStr, "^")
                If pidItemArray(4) = "PI" Then
                    extractCorpNoCerner = pidItemArray(0)
                End If
                'If testData = "PI" Then
                'Dim pidSubArray() As String
                'pidSubArray = Split(tempStr, "^")
                'extractCorpNo = pidSubArray(0)
                'End If

            Next
        Catch ex As Exception
            extractCorpNoCerner = "0"

        End Try

    End Function

    Public Sub addOrphanPaNUm(ByVal dictNVP As Hashtable)
        '20130603 - routine added to add the panum for an orphan to catch up on previous records in the STAR ITW database
        Dim myConnection As New SqlConnection(connectionString)
        Dim updatecommand As New SqlCommand
        updatecommand.Connection = myConnection
        Dim star_region As String = UCase(dictNVP("Sending Facility"))
        Dim sql As String = ""
        sql = "insert into [01visit] (mrnum, corpNo, star_region, orphanAdded, panum) "
        If IsNumeric(dictNVP("mrnum")) Then
            sql = sql & "VALUES (" & dictNVP("mrnum") & ", "
        Else
            sql = sql & "VALUES (0, "
        End If
        sql = sql & gblCorporateNumber & ", "
        sql = sql & "'" & star_region & "', "
        sql = sql & "'" & DateTime.Now & "', "
        sql = sql & "'" & dictNVP("panum") & "') "
        updatecommand.CommandText = sql
        myConnection.Open()
        updatecommand.ExecuteNonQuery()
        myConnection.Close()
    End Sub

    Public Function xlateClass(ByVal dictNVP As Hashtable) As String
        '20130619 - translate the patient class from STAR
        'PV1_2 = "Patient Class"
        'PV1_18 = "Patient Type"
        'logic from StarToInvision Trimap 2/25/2013 version 1
        xlateClass = " "
        Select Case dictNVP("Patient Class")
            Case "R"
                xlateClass = "O"
            Case "B"
                xlateClass = "E"
            Case "P"
                If UCase(Left(dictNVP("Patient Type"), 2)) = "PZ" Or UCase(Left(dictNVP("Patient Type"), 2)) = "CZ" Then
                    xlateClass = "I"
                Else
                    xlateClass = "O"
                End If
            Case Else
                xlateClass = dictNVP("Patient Class")
        End Select
    End Function
    Public Function calcSTARStatus(ByVal dictNVP As Hashtable) As String

        '20140220
        Dim myConnection As New SqlConnection(connectionString)
        Dim objCommand As New SqlCommand
        Dim updatecommand As New SqlCommand
        updatecommand.Connection = myConnection
        Dim sql As String = ""
        objCommand.Connection = myConnection
        Dim strCurrentStatus As String = ""
        Dim strNewStatus As String = ""
        Dim dataReader As SqlDataReader

        calcSTARStatus = ""

        Try
            'Get the current status if the record exists in the [001episode] table.
            sql = "select Status from [01visit] where panum = '" & extractPanum(dictNVP.Item("panum")) & "'"
            objCommand.CommandText = sql
            myConnection.Open()
            dataReader = objCommand.ExecuteReader()
            If dataReader.HasRows Then
                While dataReader.Read()
                    strCurrentStatus = dataReader.GetString(0)

                End While
            Else
                strCurrentStatus = ""
            End If
            myConnection.Close()
            dataReader.Close()



            Select Case dictNVP.Item("TriggerEventID")
                Case "A01", "A04", "A05", "A08"
                    '20140220 - only do for IP, OP or new patient
                    '20140723 - added PRE to check
                    If strCurrentStatus = "IP" Or strCurrentStatus = "OP" Or strCurrentStatus = "" Or strCurrentStatus = "PRE" Then

                        If dictNVP.Item("Patient Class") = "P" Then strNewStatus = "PRE"
                        If dictNVP.Item("Patient Class") = "I" Then strNewStatus = "IP"
                        If dictNVP.Item("Patient Class") = "O" Then strNewStatus = "OP"
                        If dictNVP.Item("Patient Class") = "R" Then strNewStatus = "OP"
                        '20140319 - add E patient class
                        If dictNVP.Item("Patient Class") = "E" Then strNewStatus = "OP"
                        '20160203 - add B patient class
                        If dictNVP.Item("Patient Class") = "B" Then strNewStatus = "IP"
                    Else
                        strNewStatus = strCurrentStatus
                    End If 'If strCurrentStatus = "IP" Or strCurrentStatus = "OP" Or strCurrentStatus = ""

                Case "A03"
                    'change second character to a "D"
                    'strNewStatus = Left(strCurrentStatus, 1) & "D"

                Case "A11"
                    '20140317 - same as for A03   
                    'strNewStatus = Left(strCurrentStatus, 1) & "D"

                Case "A06"
                    '20140317 - new status for A06.
                    strNewStatus = "IP"
                Case "A07"
                    '20140317 - new status for A07.
                    strNewStatus = "OP"
            End Select

            calcSTARStatus = strNewStatus

        Catch ex As Exception

        End Try
    End Function
    Public Function calcPatStatus(ByVal dictNVP As Hashtable) As String
        '20140317
        '20140318 - added case selection for A06 and A07
        '20140403 - Only change the patStatus if the DCDate is null or it is an A13, cancel discharge.
        calcPatStatus = ""

        Try
            Dim myConnection As New SqlConnection(connectionString)
            Dim objCommand As New SqlCommand

            Dim sql As String = ""
            objCommand.Connection = myConnection
            Dim strCurrentPatStatus As String = ""
            Dim strNewStatus As String = ""
            Dim dataReader As SqlDataReader
            Dim boolDichDateHasValue As Boolean = False

            sql = "select DcDate from [01visit] where panum = '" & dictNVP.Item("panum") & "' "
            sql = sql & " and DcDate is not null "
            objCommand.CommandText = sql
            myConnection.Open()
            dataReader = objCommand.ExecuteReader()
            If dataReader.HasRows Then
                boolDichDateHasValue = True
            End If
            myConnection.Close()
            dataReader.Close()


            sql = "select patStatus from [01visit] where panum = '" & extractPanum(dictNVP.Item("panum")) & "'"
            objCommand.CommandText = sql
            myConnection.Open()
            dataReader = objCommand.ExecuteReader()
            If dataReader.HasRows Then
                While dataReader.Read()
                    strCurrentPatStatus = dataReader.GetString(0)

                End While
            Else
                strCurrentPatStatus = ""
            End If
            myConnection.Close()
            dataReader.Close()

            If Not boolDichDateHasValue Then
                Select Case dictNVP.Item("TriggerEventID")
                    Case "A06"
                        calcPatStatus = "IA"
                    Case "A07"
                        calcPatStatus = "OA"

                    Case "A03", "A11"
                        calcPatStatus = Left(strCurrentPatStatus, 1) & "D"

                    Case "A13" '20140319 - return patStatus to A
                        calcPatStatus = Left(strCurrentPatStatus, 1) & "A"
                    Case Else

                        If dictNVP.Item("Patient Class") = "P" Then calcPatStatus = "OP"
                        If dictNVP.Item("Patient Class") = "I" Then calcPatStatus = "IA"
                        If dictNVP.Item("Patient Class") = "O" Then calcPatStatus = "OA"
                        If dictNVP.Item("Patient Class") = "R" Then calcPatStatus = "OA"
                        '20140319 - added E patient class to calcPatStatus
                        If dictNVP.Item("Patient Class") = "E" Then calcPatStatus = "EA"
                        '20160421 - added B patient class to calcPatStatus
                        If dictNVP.Item("Patient Class") = "B" Then calcPatStatus = "IA"

                End Select
            Else
                calcPatStatus = strCurrentPatStatus
            End If ' If Not boolDichDateHasValue

            If dictNVP.Item("TriggerEventID") = "A13" Then
                calcPatStatus = Left(strCurrentPatStatus, 1) & "A"
            End If

        Catch ex As Exception

        End Try
    End Function
    Public Function calcStatus(ByVal dictNVP As Hashtable) As String
        Dim strFirstCharacter As String = ""
        Dim strSecondCharacter As String = ""
        '20130620 - calculate the status based on raw Patient Class data from STAR and criteria below

        'from AL:
        'If PV1-2 = "R", then set to "O"
        'Else if PV1-2 = "B", then set to "E"
        'Else if PV1-2 = “P”, and:
        '{
        'If PV1-18 starts with “PZ” or “CZ”, then set to"I";
        'Else set to “O” 
        '}
        'otherwise pass thru the Patient Class with no translation.

        'Capture the patient class and add to the 001episode table
        '[01visit].patstatus: 1st character is the Patient Class 
        '2nd character is “A” for A01, A04 & A13
        '2nd character is “P” for A05
        '2nd character is “D” for A03


        calcStatus = ""
        '1. Get the first character using the xlateClass Function
        strFirstCharacter = xlateClass(dictNVP)
        '2. Get the second character based on type for HL7 record

        Select Case dictNVP("Event Type Code")
            Case "A01", "A04", "A13", "A06", "A07"
                strSecondCharacter = "A"
            Case "A05"
                strSecondCharacter = "P"
            Case "A03"
                strSecondCharacter = "D"
        End Select

        '3. Combine the two calculated characters if the second character was calculated,
        'otherwise send nothing.
        If strSecondCharacter <> "" Then
            calcStatus = strFirstCharacter & strSecondCharacter
        Else
            calcStatus = ""
        End If


    End Function

    Public Sub processOEC(ByVal dictNVP As Hashtable)
        '20140220 - new procedure to process oec info. Also sets patStatus to OA.
        '20140403 - don't process OEC if there is a discharge date in DCDate.
        Dim myConnection As New SqlConnection(connectionString)
        Dim updatecommand As New SqlCommand
        updatecommand.Connection = myConnection
        Dim dataReader As SqlDataReader
        Dim sql As String = ""
        Dim boolDTChangedIsNULL As Boolean = True
        Dim boolDichDateHasValue As Boolean = False

        sql = "select DcDate from [01visit] where panum = '" & dictNVP.Item("panum") & "' "
        sql = sql & " and DcDate is not null "
        updatecommand.CommandText = sql
        myConnection.Open()
        dataReader = updatecommand.ExecuteReader()
        If dataReader.HasRows Then
            boolDichDateHasValue = True
        End If
        myConnection.Close()
        dataReader.Close()

        If Not boolDichDateHasValue Then

            sql = "select dtchanged from [01visit] where panum = '" & dictNVP.Item("panum") & "' "
            sql = sql & " and dtchanged is not null "
            updatecommand.CommandText = sql
            myConnection.Open()

            dataReader = updatecommand.ExecuteReader()
            If dataReader.HasRows Then
                boolDTChangedIsNULL = False
            End If
            myConnection.Close()
            dataReader.Close()

            '20160113 - change patient type criteria for McareULH Star to OVT only
            '20160203 - add Patient Type = "OMT"
            If (dictNVP.Item("Patient Type") = "OVT" Or dictNVP.Item("Patient Type") = "OMT") And dictNVP.Item("Patient Class") = "O" Then
                sql = "Update [01visit] set "
                sql = sql & "Status = 'OEC'"
                sql = sql & ", PatStatus = 'OA'"
                sql = sql & ", OECLimit = 23"

                If boolDTChangedIsNULL Then
                    sql = sql & " ,dtChanged = '" & ConvertDate(dictNVP.Item("01visit.AdminDate")) & "'"
                End If

                sql = sql & ", OECExpire = '" & DateAdd("h", 23, CDate(ConvertDate(dictNVP.Item("01visit.AdminDate")))) & "' "


                sql = sql & "WHERE panum = '" & dictNVP.Item("panum") & "'"

                updatecommand.CommandText = sql
                myConnection.Open()
                updatecommand.ExecuteNonQuery()
                myConnection.Close()
            End If
        End If ' If Not boolDichDateHasValue
    End Sub

    Public Sub ProcessA34(ByVal dictNVP As Hashtable)
        '20141021 - write A34 Information to PatientGlobal database, table = A34Queue
        '20150112 - added old corpNo also
        'Dim A34connectionString As String = "server=10.48.242.249,1433;database=PatientGlobal;uid=sysmax;pwd=Condor!"
        'Dim A34connectionString As String = "server=10.48.64.5\sqlexpress;database=PatientGlobal;uid=sysmax;pwd=Condor!"
        Dim A34connectionString As String = conIniFile.GetString("Strings", "MCAREULHA34", "(none)")

        Dim myConnection As New SqlConnection(A34connectionString)
        Dim updatecommand As New SqlCommand
        updatecommand.Connection = myConnection
        Dim sql As String = ""
        Dim strCorpNo As String = dictNVP.Item("CORPNO")
        Dim strMrNum As String = dictNVP.Item("mrnum")
        Dim strOldMRNUM As String = dictNVP.Item("oldmrnum")
        Dim strRegion As String = UCase(dictNVP("Sending Facility"))
        Dim strOldCorpNo As String = dictNVP.Item("oldcorpno")
        Try
            If strCorpNo <> "" And strMrNum <> "" And strOldMRNUM <> "" Then
                sql = "Insert [A34Queue] "
                sql = sql & "(facility, originalMRN, NewMRN, originalCorp, NewCorp, RequestedDate) "
                sql = sql & "VALUES ("
                sql = sql & "'" & strRegion & "', " & strOldMRNUM & ", " & strMrNum & ", " & strOldCorpNo & ", " & strCorpNo & ", "
                sql = sql & "'" & DateTime.Now & "') "
                updatecommand.CommandText = sql
                myConnection.Open()
                updatecommand.ExecuteNonQuery()
                myConnection.Close()
            End If
        Catch ex As Exception
            functionError = True
            gblLogString = gblLogString & "A34 Process Error" & vbCrLf
            gblLogString = gblLogString & ex.Message & vbCrLf
            'LogFile.Close()
            Exit Sub
        End Try

    End Sub

    Public Sub ProcessIN1_14(ByVal dictNVP As Hashtable, ByVal tempstr As String)

        'Dim Int As Integer
        'Dim tempStr As String = ""
        'For Int = 1 To gblInsCount
        'If Int = 1 Then
        'tempStr = ""
        'End If
        'If Int > 1 Then
        'tempstr = "_000" & Int
        'End If

        'check to see if records exist
        'Call dbo.smc_InsAuthSelect
        'result contains rows update else insert
        Dim objDBCommand As New SqlCommand
        Dim objDBCommand2 As New SqlCommand
        Dim objDBCommand3 As New SqlCommand
        Dim objDBCommand4 As New SqlCommand
        Dim dreader As SqlDataReader
        'Dim dreader2 As SqlDataReader
        Dim sql As String = ""
        Dim IN114array()
        'STAR_Plancode = Trim(Replace(dictNVP("iplancode2" & tempStr), "'", "''")) & Trim(Replace(dictNVP("iplancode" & tempStr), "'", "''"))

        Dim plancode As String = dictNVP.Item("iplancode" & tempstr)
        Dim plancode2 As String = dictNVP.Item("iplancode2" & tempstr)
        Dim fullcode As String = plancode2 & plancode
        Dim panum As String = dictNVP.Item("panum")
        Dim I As String = ""
        Dim ID As Integer
        Dim insert As Boolean = True

        Using conn As New SqlConnection(connectionString)
            With objDBCommand

                .Connection = conn
                .Connection.Open()

                sql = "Select patconum "
                sql += " FROM [03Insurer] i "
                sql += " WHERE i.iplancode = '" & fullcode & "' "
                sql += " and i.panum = '" & panum & "'"
                .CommandText = sql
                dreader = objDBCommand.ExecuteReader()
                While dreader.Read
                    I = dreader("patconum")
                End While

            End With
        End Using
        If I <> "" Then
            Using conn As New SqlConnection(connectionString)
                ID = Convert.ToInt32(I)
                With objDBCommand2
                    .Connection = conn
                    .Connection.Open()

                    sql += "DELETE FROM [03InsAuthReceive] "
                    sql += " WHERE INSID = '" & ID & "' "
                    .CommandText = sql
                    objDBCommand2.ExecuteNonQuery()
                End With
            End Using

            Using conn As New SqlConnection(connectionString)
                With objDBCommand3
                    .Connection = conn
                    .Connection.Open()

                    'Dim position As Integer
                    Dim fromDate = String.Empty
                    Dim toDate = String.Empty
                    Dim Authcode = String.Empty
                    Dim IN114 As String = dictNVP("AuthNum" & tempstr)
                    IN114array = IN114.Split("^")

                    If IN114array(0) <> "" Then
                        Authcode = IN114array(0)
                    End If

                    .CommandText = "dbo.smc_InsAuthUpdateReceive"
                    .CommandType = CommandType.StoredProcedure
                    .Parameters.Clear()
                    .Parameters.AddWithValue("@InsID", ID)
                    .Parameters.AddWithValue("@positionNum", 1)
                    .Parameters.AddWithValue("@AuthCode", Authcode)
                    .Parameters.AddWithValue("@fromDate", fromDate)
                    .Parameters.AddWithValue("@toDate", toDate)
                    .Parameters.AddWithValue("@insert", True)

                    objDBCommand3.ExecuteNonQuery()

                End With

            End Using


            Using conn As New SqlConnection(connectionString)
                With objDBCommand4
                    .Connection = conn
                    .Connection.Open()

                    Dim ZGI2 As String = dictNVP("Additional Auths" & tempstr)
                    If Not ZGI2 Is Nothing Then
                        Dim ZGI() = ZGI2.Split("~")
                        Dim position As Integer = 1

                        Dim valueArray() As String
                        For Each value As String In ZGI
                            Dim fromDate = String.Empty
                            Dim toDate = String.Empty
                            Dim Authcode = String.Empty

                            position += 1

                            If value <> "" Then

                                Dim cnt As Integer = 0
                                For Each c As Char In value
                                    If c = "^" Then
                                        cnt += 1
                                    End If
                                Next


                                If cnt = 2 Then
                                    valueArray = value.Split("^")
                                    If valueArray(0) <> "" Then
                                        Authcode = valueArray(0)
                                    End If
                                    If valueArray(1) <> "" Then
                                        fromDate = valueArray(1)
                                    End If
                                    If valueArray(2) <> "" Then
                                        toDate = valueArray(2)
                                    End If
                                ElseIf cnt = 1 Then
                                    valueArray = value.Split("^")
                                    If valueArray(0) <> "" Then
                                        Authcode = valueArray(0)
                                    End If
                                    If valueArray(1) <> "" Then
                                        fromDate = valueArray(1)
                                    End If
                                ElseIf cnt = 0 Then
                                    valueArray = value.Split("^")
                                    If valueArray(0) <> "" Then
                                        Authcode = valueArray(0)
                                    End If

                                End If

                            End If


                            .CommandText = "dbo.smc_InsAuthUpdateReceive"
                            .CommandType = CommandType.StoredProcedure
                            .Parameters.Clear()
                            .Parameters.AddWithValue("@InsID", ID)
                            .Parameters.AddWithValue("@positionNum", position)
                            .Parameters.AddWithValue("@AuthCode", Authcode)
                            .Parameters.AddWithValue("@fromDate", fromDate)
                            .Parameters.AddWithValue("@toDate", toDate)
                            .Parameters.AddWithValue("@insert", True)

                            objDBCommand4.ExecuteNonQuery()

                        Next
                    End If
                End With

            End Using

            'Using conn As New SqlConnection(connectionString)
            '    With objDBCommand3
            '        .Connection = conn
            '        .Connection.Open()

            '        Dim IN114 As String = dictNVP("AuthNum")
            '        Dim IN114array() = IN114.Split("~")
            '        Dim position As Integer
            '        Dim fromDate = DBNull.Value
            '        Dim toDate = DBNull.Value
            '        Dim Authcode = DBNull.Value
            '        Dim valueArray() As String
            '        For Each value As String In IN114array
            '            position += 1


            '            If position = 1 Then
            '                Authcode = Replace(value, "^", "")

            '            Else
            '                valueArray = value.Split("^")
            '                If valueArray(0) <> "" Then
            '                    Authcode = valueArray(0)
            '                End If
            '                If value(1) <> "" Then
            '                    fromDate = valueArray(1)
            '                End If
            '                If value(2) <> "" Then
            '                    toDate = valueArray(2)
            '                End If
            '            End If
            '            .CommandText = "dbo.smc_InsAuthUpdateReceive"
            '            .CommandType = CommandType.StoredProcedure
            '            .Parameters.Clear()
            '            .Parameters.AddWithValue("@InsID", ID)
            '            .Parameters.AddWithValue("@positionNum", position)
            '            .Parameters.AddWithValue("@AuthCode", Authcode)
            '            .Parameters.AddWithValue("@fromDate", fromDate)
            '            .Parameters.AddWithValue("@toDate", toDate)
            '            .Parameters.AddWithValue("@insert", True)

            '            objDBCommand3.ExecuteNonQuery()

            '        Next
            '    End With

            'End Using
        End If
        'Next

    End Sub

    Private Sub checkZ47(theFile As FileInfo, myfile As StreamReader)
        Dim newProblemDir As String = strOutputDirectory & "Z47\"

        If Not Directory.Exists(newProblemDir) Then
            Directory.CreateDirectory(newProblemDir)
        End If

        'make copy in the problems directory delete any previous ones with same name
        Dim fi2 As FileInfo = New FileInfo(strOutputDirectory & "Z47\" & theFile.Name)
        fi2.Delete()
        theFile.CopyTo(strOutputDirectory & "Z47\" & theFile.Name)

        'get rid of the file so it doesn't mess up the next run.
        myfile.Close()
        If theFile.Exists Then
            theFile.Delete()
        End If
    End Sub

End Module
