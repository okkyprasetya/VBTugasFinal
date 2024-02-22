Imports System.Data.SqlClient
Imports studentAdmission.BO.studentAdmission.BO

Namespace DAL
    Public Class verificatorDAL
        Private Const strConn As String = "Server=.\BSISqlExpress;Database=finalProject;Trusted_Connection=True;"

        Public Function GetAll() As IEnumerable(Of Users)
            Dim users As New List(Of Users)()
            Using conn As New SqlConnection(strConn)
                Dim strSql As String = "SELECT * FROM dbo.UserData WHERE RoleID = 1"
                Dim cmd As New SqlCommand(strSql, conn)
                conn.Open()
                Dim dr As SqlDataReader = cmd.ExecuteReader()

                If dr.HasRows Then
                    While dr.Read()
                        users.Add(New Users() With {
                            .userId = Convert.ToInt32(dr("UserID")),
                            .uFName = dr("FirstName").ToString(),
                            .uMName = dr("MiddleName").ToString(),
                            .uLName = dr("LastName").ToString(),
                            .userEmail = dr("UserEmail").ToString(),
                            .roleID = Convert.ToInt32(dr("RoleID"))
                        })
                    End While
                End If

                dr.Close()
                cmd.Dispose()
            End Using

            Return users
        End Function

        Public Sub AddUser(firstName As String, middleName As String, lastName As String, emailAddress As String, password As String)
            Using conn As New SqlConnection(strConn)
                Dim cmd As New SqlCommand("dbo.addVerificator", conn)
                cmd.CommandType = System.Data.CommandType.StoredProcedure

                ' Add parameters
                cmd.Parameters.AddWithValue("@fname", firstName)
                cmd.Parameters.AddWithValue("@midname", middleName)
                cmd.Parameters.AddWithValue("@lname", lastName)
                cmd.Parameters.AddWithValue("@password", password)
                cmd.Parameters.AddWithValue("@email", emailAddress)

                conn.Open()
                cmd.ExecuteNonQuery()
            End Using
        End Sub

        Public Sub EditUser(firstName As String, middleName As String, lastName As String, uid As Integer)
            Using conn As New SqlConnection(strConn)
                Dim cmd As New SqlCommand("dbo.editVerificatorData", conn)
                cmd.CommandType = System.Data.CommandType.StoredProcedure

                ' Add parameters
                cmd.Parameters.AddWithValue("@fname", firstName)
                cmd.Parameters.AddWithValue("@mname", middleName)
                cmd.Parameters.AddWithValue("@lname", lastName)
                cmd.Parameters.AddWithValue("@uid", uid)

                conn.Open()
                cmd.ExecuteNonQuery()
            End Using
        End Sub

        Public Sub DeleteUser(uid As Integer)
            Using conn As New SqlConnection(strConn)
                Dim cmd As New SqlCommand("dbo.deleteVerificator", conn)
                cmd.CommandType = System.Data.CommandType.StoredProcedure

                ' Add parameters
                cmd.Parameters.AddWithValue("@uid", uid)

                conn.Open()
                cmd.ExecuteNonQuery()
            End Using
        End Sub

        Public Function GetDataByID(ID As Integer) As IEnumerable(Of Users)
            Dim users As New List(Of Users)()
            Using conn As New SqlConnection(strConn)
                Dim strSql As String = "SELECT * FROM dbo.UserData where userID = @ID"
                Dim cmd As New SqlCommand(strSql, conn)
                cmd.Parameters.AddWithValue("@ID", ID)

                conn.Open()
                Dim dr As SqlDataReader = cmd.ExecuteReader()

                If dr.HasRows Then
                    While dr.Read()
                        users.Add(New Users() With {
                            .userId = Convert.ToInt32(dr("UserID")),
                            .uFName = dr("FirstName").ToString(),
                            .uMName = dr("MiddleName").ToString(),
                            .uLName = dr("LastName").ToString(),
                            .userEmail = dr("UserEmail").ToString(),
                            .roleID = Convert.ToInt32(dr("RoleID"))
                        })
                    End While
                End If

                dr.Close()
                cmd.Dispose()
            End Using

            Return users
        End Function
    End Class
End Namespace

