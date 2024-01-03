Imports MySql.Data.MySqlClient

Module modConnection

    Public strLogDesc As String
    Public strInitials As String = "BIZ"
    Public strLogUpdateType As String = "IMPORT"
    Public strSQLConn As String
    Public strErrorMessage As String

    Public Sub DetermineDB()
        If frmMain.blnProd Then
            frmMain.strDB = "paymentmanager"
        Else
            frmMain.strDB = "louie"
        End If
    End Sub
    Public Sub SetConnection()

        strSQLConn = "Server=hh-mp-pmgr;user=PMUser;password=$h0wM3Th3M0n3y$;database=" & frmMain.strDB & ";default command timeout=360"

        DetermineDB()

    End Sub

    Public Function ExecuteQuery_ReturnInt32(_strQueryValue, _intRetVal)

        SetConnection()

        Using cn As New MySqlConnection(strSQLConn)
            Using sqlCommand As New MySqlCommand

                With sqlCommand
                    .Connection = cn
                    .CommandText = _strQueryValue
                    .CommandType = CommandType.StoredProcedure
                    SetQueryParameters(sqlCommand, _strQueryValue)
                End With

                Try
                    If cn.State <> ConnectionState.Open Then cn.Open()
                    _intRetVal = sqlCommand.ExecuteScalar()
                Catch ex As Exception
                    strErrorMessage = "Import Error - " & ex.Message
                    UpdateLog_Error()
                End Try

            End Using
        End Using

        Return _intRetVal
    End Function

    Public Sub ExecuteQuery_StoredProcedure(strProcName As String)

        SetConnection()

        Using cn As New MySqlConnection(strSQLConn)
            Using sqlCommand As New MySqlCommand

                With sqlCommand
                    .Connection = cn
                    .CommandText = strProcName
                    .CommandType = CommandType.StoredProcedure
                    SetQueryParameters(sqlCommand, strProcName)
                End With

                Try
                    If cn.State <> ConnectionState.Open Then cn.Open()
                    sqlCommand.ExecuteNonQuery()
                Catch ex As Exception
                    strErrorMessage = "Import Error - " & ex.Message
                    UpdateLog_Error()
                End Try

            End Using
        End Using

    End Sub

    Public Sub ExecuteQuery_ImportQuery(strImportString)

        SetConnection()

        Using cn As New MySqlConnection(strSQLConn)
            Using sqlCommand As New MySqlCommand

                With sqlCommand
                    .CommandText = strImportString
                    .Connection = cn
                    .CommandType = CommandType.Text
                End With

                Try
                    If cn.State <> ConnectionState.Open Then cn.Open()
                    sqlCommand.ExecuteNonQuery()
                Catch ex As Exception
                    strErrorMessage = "Import Error - " & ex.Message
                    UpdateLog_Error()
                End Try

            End Using
        End Using

    End Sub

    Public Sub SetQueryParameters(sqlcmd As MySqlCommand, strProc As String)

        Select Case strProc

            Case "ADD_LOG"
                Dim parLT As MySqlParameter = New MySqlParameter("LT", strLogUpdateType)
                parLT.Direction = ParameterDirection.Input
                parLT.DbType = DbType.String
                sqlcmd.Parameters.Add(parLT)

                Dim parCM As MySqlParameter = New MySqlParameter("CM", strLogDesc)
                parCM.Direction = ParameterDirection.Input
                parCM.DbType = DbType.String
                sqlcmd.Parameters.Add(parCM)

                Dim parINIT As MySqlParameter = New MySqlParameter("INIT", strInitials)
                parINIT.Direction = ParameterDirection.Input
                parINIT.DbType = DbType.String
                sqlcmd.Parameters.Add(parINIT)

            Case "REFRESH_FROM_PRODUCTION"
                ' Do Nothing

            Case "CUBS_TX_TO_CPT"
                ' Do nothing

            Case "TX_TO_NOCREDIT"
                ' Do nothing

            Case "ARCHIVE_MRREVIEW"
                ' Do nothing

            Case Else
                Dim parTBL As MySqlParameter = New MySqlParameter("TBL", "temp_trans")
                parTBL.Direction = ParameterDirection.Input
                parTBL.DbType = DbType.String
                sqlcmd.Parameters.Add(parTBL)

        End Select

    End Sub

End Module
