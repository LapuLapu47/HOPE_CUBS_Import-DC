Module modQueries

    Public strTempFile As String
    Public intErrorCount As Integer

    Public Sub ExecuteQuery(intQueryType As Integer, strQueryVal As String)

        Select Case intQueryType
            Case 1
                ExecuteQuery_StoredProcedure(strQueryVal)

            Case Else
                ImportFile(intQueryType, strQueryVal)
        End Select

        _ADD_LOG(intQueryType, strQueryVal)

    End Sub

    Public Sub ImportFile(intImportType As Integer, strQueryVal As String)

        ' NOTE: strQueryVal contains the file name to import at this point
        strTempFile = strQueryVal

        Select Case intImportType

            Case 2 ' STANDARD IMPORT
                strQueryVal = "LOAD Data " &
                "LOCAL INFILE '\\\\hh-fileserver01\\tempul2\\DC\\coll-fig\\IMPORT\\" & strQueryVal & "' INTO TABLE temp_trans " &
                "COLUMNS TERMINATED BY '|'  ## This should be your delimiter " &
                "LINES TERMINATED BY '\n' " &
                    "(@TransactionID,@DatePosted,@PortNumber,@Debtor,@TransCode,'0','0',@CollInitials,@DeskType,@PayAmt,@TransDate,@PostDate,@StatusCode,@ClientID,@ClientName,@Desk,@PayAdj,@FieldToUpdate,@NSF,@PymtType,@Shared,@Description,@TransSpredSeq,@Batch,@Overpayment,@Packet,@ActionCode,@ActionCodeDate,@ClientDays,@DateReported,@Logon,@OrigTran,@Reason,@Credited,@ReviewedBy) " &
                "SET TransactionID = @TransactionID, " &
                "DatePosted = @DatePosted, " &
                "PortNumber = nullif(@PortNumber,''), " &
                "Debtor = If(CHAR_LENGTH(@Debtor) > 8, Left(@Debtor, 8),@Debtor), " &
                "TransCode = nullif(@TransCode,''), " &
                "CommRate = 0.00, " &
                "CommAmt = 0.00, " &
                "CollInitials = If(CHAR_LENGTH(@CollInitials) > 4,'XXX',@CollInitials), " &
                "DeskType = nullif(@DeskType,''), " &
                "PayAmt = nullif(@PayAmt,''), " &
                "TransDate = @TransDate, " &
                "PostDate = @PostDate, " &
                "StatusCode = nullif(@StatusCode,''), " &
                "ClientID = nullif(@ClientID,''), " &
                "ClientName = nullif(@ClientName,''), " &
                "Desk = nullif(@Desk,''), " &
                "PayAdj = nullif(@PayAdj,''), " &
                "FieldToUpdate = nullif(@FieldToUpdate,''), " &
                "NSF = nullif(@NSF,''), " &
                "PymtType = nullif(@PymtType,''), " &
                "Shared = nullif(@Shared,''), " &
                "Description = nullif(@Description,''), " &
                "TransSpredSeq = nullif(@TransSpredSeq,''), " &
                "Batch = nullif(@Batch,''), " &
                "Overpayment = nullif(@Overpayment,''), " &
                "Packet = nullif(@Packet,''), " &
                "ActionCode = nullif(@ActionCode,''), " &
                "ActionCodeDate = @ActionCodeDate, " &
                "ClientDays = If(@ClientDays = '',45, @ClientDays), " &
                "DateReported = @DateReported, " &
                "Logon = nullif(@Logon,''), " &
                "Reason = nullif(@Reason,''), " &
                "OrigTrans = nullif(@OrigTran,''), " &
                "Credited = nullif(@Credited,''), " &
                "ReviewedBy = @ReviewedBy;"

            Case Else
                frmMain.lblError.Text = "Invalid File Import Type: " & Str(intImportType)
                frmMain.blnSTOP = True
                Exit Sub

        End Select

        ExecuteQuery_ImportQuery(strQueryVal)

    End Sub

    Public Sub _ADD_LOG(_intQueryType As Integer, _strQueryVal As String)

        Dim _blnLogFlag As Boolean = True

        Select Case _intQueryType

            Case 1

                Select Case _strQueryVal
                    Case "REFRESH_FROM_PRODUCTION"
                        strLogDesc = "Updated transaction tables in hopebackup"
                    Case "IMPORT_DATE"
                        strLogDesc = " - Import Date Updated"
                    Case "UPDATE_BASE_RATE"
                        strLogDesc = " - Base Rate Set"
                    Case "UPDATE_PACKET_TABLE"
                        strLogDesc = " - Packet Table Updated"
                    Case "DROP_CREATE_KDOR"
                        strLogDesc = ""
                        _blnLogFlag = False
                    Case "KDOR_UPDATE"
                        strLogDesc = " - KDOR 2nds Rates Updated"
                    Case "MOVE_TRANSACTIONS"
                        strLogDesc = " - Transactions moved"
                    Case "CREATE_TMPXREF"
                        strLogDesc = ""
                        _blnLogFlag = False
                    Case "REFRESH_CPTXREF"
                        strLogDesc = ""
                        _blnLogFlag = False
                    Case "UPDATE_CPT"
                        strLogDesc = ""
                        _blnLogFlag = False
                    Case "CUBS_TX_TO_CPT"
                        strLogDesc = "CUBS_TX_TO_CPT - Moved linked CUBS TX"
                    Case "TX_TO_NOCREDIT"
                        strLogDesc = "TX_TO_NOCREDIT - Moved XXX & HHH TX to tblnocredit"
                    Case "ARCHIVE_MRREVIEW"
                        strLogDesc = "ARCHIVE_MRREVIEW - TX older than 60 days moved from mrreview to nocredit"
                    Case "ARCHIVE_FROM_PRODUCTION"
                        strLogDesc = "Ran Archive Process"
                    Case "EXPORT_TRANSACTION_DETAIL"
                        strLogDesc = "Exported Detail Report"
                    Case Else
                        frmMain.blnSTOP = True
                        frmMain.lblError.Text = "Unable to Execute Stored Procedure - " & _strQueryVal
                        Exit Sub
                End Select

                strLogDesc = "Stored Procedure " & strLogDesc
            Case 9
                strLogDesc = "Import completed in " & frmMain.lblRunTime2.Text
            Case Else
                strLogDesc = "File Import - Completed Import for " & strTempFile

        End Select

        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        If _blnLogFlag = True Then

            ExecuteQuery_StoredProcedure("ADD_LOG")

        End If

    End Sub


    Public Sub VerifyImport()

        Dim intIMPCNT As Int32

        intIMPCNT = ExecuteQuery_ReturnInt32("IMPORT_COUNT", intIMPCNT)

        If intIMPCNT < 1 Then
            ExecuteQuery_StoredProcedure("CLEAR_TABLE")
            strLogDesc = "No Records Imported"
            ExecuteQuery_StoredProcedure("ADD_LOG")
        Else
            strLogDesc = intIMPCNT & " Records Imported"
            ExecuteQuery_StoredProcedure("ADD_LOG")
        End If

    End Sub
    Public Sub UpdateLog_Error()

        frmMain.blnSTOP = True
        frmMain.lblError.Text = strErrorMessage
        strLogDesc = strErrorMessage

        If intErrorCount = 0 Then
            intErrorCount += 1
            ExecuteQuery_StoredProcedure("ADD_LOG")
        Else
            Exit Sub
        End If

        MsgBox(strErrorMessage)

    End Sub

End Module
