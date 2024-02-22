Imports studentAdmission.BO.studentAdmission.BO
Imports studentAdmission.DAL.DAL
Imports YourProjectNamespace.DAL

Module Module1
    Sub Main(args As String())
        'DOTNET FRAMEWORK VB'
        ' Instantiate repositories
        Dim verificatorRepository As New verificatorDAL()
        Dim studentRepository As New studentDAL()

        'Retrieve all verificator only
        'Console.WriteLine("Retrieving all Verificator:")
        'Dim allUsers = verificatorRepository.GetAll()
        'For Each user In allUsers
        '    Console.WriteLine($"UserID: {user.userId}, Name: {user.uFName} {user.uMName} {user.uLName}, Email: {user.userEmail}, RoleID: {user.roleID}")
        'Next

        ' Add a new verificator
        'Console.WriteLine(vbCrLf & "Adding a new user:")
        'verificatorRepository.AddUser("John", "Doe", "john.doe@example.com", "password123", "password123")

        ' Update verificator
        'Console.WriteLine(vbCrLf & "Updating an existing user:")
        'verificatorRepository.EditUser("Jane", "Smith", "jane.smith@example.com", 30)

        '' Delete a verificator
        'Console.WriteLine(vbCrLf & "Deleting a user:")
        'verificatorRepository.DeleteUser(30)
        'verificatorRepository.DeleteUser(31)


        'Retrieve all students
        Console.WriteLine("Retrieve all students: ")
        Dim allStudent = studentRepository.GetAll()
        For Each student In allStudent
            Console.WriteLine($"UserID: {student.userId}, Name: {student.uFName} {student.uMName} {student.uLName}, Email: {student.userEmail}, RoleID: {student.roleID}")
        Next

        'Add a new student/ register student'
        'Console.WriteLine("Register student")
        'studentRepository.AddUser("Michael", "Liu", ".S", "okky@gmail.com", "1234")

        ''Complete data by student (only general data)'
        'Console.WriteLine("Completing data: ")
        'Console.WriteLine(vbCrLf & "Updating an existing user:")
        'studentRepository.EditUser(20, "1234567", "1999-10-01", True, 2)


        Console.WriteLine(vbCrLf & "Press any key to exit...")
        Console.ReadKey()
    End Sub
End Module
