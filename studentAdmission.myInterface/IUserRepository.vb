

Imports studentAdmission.BO.studentAdmission.BO

Namespace MyInterfaces.DAL
    Public Interface IUserRepository
        Function GetAll() As IEnumerable(Of Users)
        Sub AddUser(firstName As String, middleName As String, lastName As String, emailAddress As String, password As String)
        Sub EditUser(firstName As String, middleName As String, lastName As String, uid As Integer)
        Sub DeleteUser(uid As Integer)
        Function GetDataByID(ID As Integer) As IEnumerable(Of Users)
    End Interface
End Namespace
