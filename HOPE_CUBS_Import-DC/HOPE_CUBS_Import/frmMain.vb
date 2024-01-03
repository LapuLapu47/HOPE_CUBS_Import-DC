Imports System.IO

Public Class frmMain

    Public blnProd As Boolean = True
    Public strDB As String

    Public strFileName As String
    Public strTmpFileName As String
    Public intC As Integer
    Public blnSTOP As Boolean = False
    Public swTotal As New Stopwatch
    Public strQueryValue As String
    Public intQueryType As String

    Private Sub Form1_Shown(sender As Object, e As EventArgs) _
     Handles MyBase.Shown

        MyBase.Refresh()

        For intC = 1 To 11

            ' *************************************************************************************************
            ' ************************************** IMPORTANT ************************************************
            ' THIS PROGRAM WAS MODELED AFTER THE HOPE IMPORT FOR CUBS, BUT ALL REFERENCES TO CPT/KDOR HAVE BEEN 
            ' REMOVED AS THEY ARE NO LONGER NEEDED
            '
            ' IMPORTANT CHANGES: 
            '       IMPORT PATH FOR THE FILES HAS BEEN CHANGED TO \\hh-fileserver01\tempul2\DC\coll-fig\IMPORT\
            '       strFileName CHANGED TO 
            '
            ' *************************************************************************************************
            ' *************************************************************************************************

            DetermineDB()

            intQueryType = 1                                       
            ' intQueryType Values
            ' 0 = No SQL Query Executed
            ' 1 = Stored Procedure - *** DEFAULT ***


            ' STEPS TO LOAD THE DATA FILES FROM CUBS INTO THE DATABASE
            '
            '     STEP 1 - Validate all files exist in import location 
            '
            '     STEP 2 - Load Payment_Manage_Export_file_mmddyyyy.txt into temp_trans and update 
            '        PT 1 - Import Payment_Manage_Export_file_mmddyyyy.txt into temp_trans
            '        PT 2 - Verify transactions imported are > 0
            '        PT 3 - Set ImportDate = Today for all temp_trans records
            '        PT 4 - Update the Commision Rate & Amount for all temp_trans records
            '        PT 5 - Insert any new packeted debtors into the reference table
            '
            '     STEP 3 - Remove duplicate transactions from temp_trans, then push transactions into appropriate tables from temp_trans
            '        PT 1 - Remove transactions already in live tables, then push remaining transactions over from temp_trans            
            '
            '     STEP 4 - Cleanup Steps
            '        PT 1 - Move any transactions with HHH or XXX from tbltransactions to tblnocredit
            '        PT 2 - Move transactions older than 60 days from tblmanagerreview to tblnocredit
            '        PT 3 - Archive qualifying transactions from paymentmanager to hope_archive databases
            '        PT 4 - Optional daily export of tbltransactions to hh-mp-pmgr as needed. Can be enabled or disabled from the SP in MySQL

            If blnSTOP = False Then
                Select Case intC
                                            
                    Case 1
                        
                        intQueryType = 0              ' 0 = No SQL Query Executed
                        ValidateFilesExist()

                    Case 2  ' STEP 2 - PT 1 - Import DC_Payment_Manage_Export_file_mmddyyyy.txt into temp_trans

                        strQueryValue = "DC-" & strFileName
                        intQueryType = 2    ' Standard Import
                    
                    Case 3  ' STEP 2 - PT 2 - Verify transactions imported are > 0

                        intQueryType = 0                           ' 0 = No SQL Query Executed
                        VerifyImport()

                    Case 4  ' STEP 2 - PT 3 - Set ImportDate = Today for all temp_trans records

                        strQueryValue = "IMPORT_DATE"

                    Case 5  ' STEP 2 - PT 4 - Update the Commision Rate & Amount for all temp_trans records

                        strQueryValue = "UPDATE_BASE_RATE"

                    Case 6  ' STEP 2 - PT 5 - Insert any new packeted debtors into the reference table

                        strQueryValue = "UPDATE_PACKET_TABLE"

                    Case 7  ' STEP 3 - PT 1 - Remove transactions already in live tables, then push remaining transactions over from temp_trans

                        strQueryValue = "MOVE_TRANSACTIONS"

                    Case 8  ' STEP 4 - PT 1 - Move any transactions with HHH or XXX from tbltransactions to tblnocredit

                        strQueryValue = "TX_TO_NOCREDIT"

                    Case 9  ' STEP 4 - PT 2 - Move transactions older than 60 days from tblmanagerreview to tblnocredit

                        strQueryValue = "ARCHIVE_MRREVIEW"

                    Case 10 ' STEP 4 - PT 3 - Archive qualifying transactions from paymentmanager to hope_archive databases

                        strQueryValue = "ARCHIVE_FROM_PRODUCTION"

                    Case 11 ' STEP 4 - PT 4 - Optional daily export of tbltransactions to hh-mp-pmgr as needed. Can be enabled or disabled from the SP in MySQL

                        strQueryValue = "EXPORT_TRANSACTION_DETAIL"

                End Select

                If intQueryType > 0 Then

                    ExecuteQuery(intQueryType, strQueryValue)
                End If

                UpdateInterface()

            End If

        Next intC

        swTotal.Stop()
        MyBase.Refresh()

        If blnSTOP = False Then
            intQueryType = 9
            strQueryValue = ""
            _ADD_LOG(intQueryType, strQueryValue)
            Me.Close()
        End If

    End Sub

        Public Sub ValidateFilesExist()

        Dim sMMDDYYYY As String = Format(Now(), "MMddyyyy")
        Dim sSource As String = "\\hh-fileserver01\tempul2\dc\coll-fig\import\"
        'Dim sSource As String = "\\cubs_files\pcshare\tempul\export\coll-fig\"
        'Dim sTarget As String = "\\hh-fileserver01\tempul2\coll-fig\"
        strFileName = "Payment_Manage_Export_file_" & sMMDDYYYY & ".txt"

        For f = 1 To 1
            Select Case f
                Case 1
                    strTmpFileName = "DC-" & strFileName
                'Case 2
                '    strTmpFileName = "KDOR_" & strFileName
                'Case 3
                '    strTmpFileName = "CPT_" & strFileName
                'Case 4
                '    strTmpFileName = "CPT_CRREF_" & strFileName
            End Select
            If Not File.Exists(sSource & strTmpFileName) Then
                blnSTOP = True
                If lblError.Text = "" Then
                    lblError.Text = "MISSING: " & strTmpFileName
                Else
                    lblError.Text = lblError.Text & ", " & strTmpFileName
                End If
            'Else
                'If Not File.Exists(sTarget & strTmpFileName) Then
                    'File.Copy(sSource & strTmpFileName, sTarget & strTmpFileName, True)
                'End If
            End If
        Next f

    End Sub

    Public Sub UpdateInterface()

        CheckedListBox1.SetItemCheckState(intC-1, CheckState.Checked)

        lblRunTime2.Text = Int((swTotal.ElapsedMilliseconds / 1000) / 60).ToString("00") & ":" & Int((swTotal.ElapsedMilliseconds / 1000) Mod 60).ToString("00")

        MyBase.Refresh()

    End Sub

    Private Sub FormLoad(sender As Object, e As EventArgs) Handles MyBase.Load

        CheckedListBox1.Items.Add("File Check")
        CheckedListBox1.Items.Add("Import Standard Files")
        CheckedListBox1.Items.Add("-- Validate Import")
        CheckedListBox1.Items.Add("-- Update Import Date")
        CheckedListBox1.Items.Add("-- Update Base Rate & Credit")
        CheckedListBox1.Items.Add("-- Update Packet Table")
        CheckedListBox1.Items.Add("Move Transactions")
        CheckedListBox1.Items.Add("House Payments to NoCredit")
        CheckedListBox1.Items.Add("Old MR Review TX to NoCredit")
        CheckedListBox1.Items.Add("Archive transactions in from Production")
        CheckedListBox1.Items.Add("Export Detail Report")


        swTotal.Start()

    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick

        lblRunTime2.Text = Int((swTotal.ElapsedMilliseconds / 1000) / 60).ToString("00") & ":" & Int((swTotal.ElapsedMilliseconds / 1000) Mod 60).ToString("00")

        MyBase.Refresh()

    End Sub

End Class

