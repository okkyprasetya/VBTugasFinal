Imports studentAdmission.BO
Imports studentAdmission.DAL
Imports studentAdmission.DAL.DAL

Module Module1
    Sub Main(args As String())
        ' Instantiate userDAL
        Dim userRepository As New userDAL()

        ' Retrieve all users
        Console.WriteLine("Retrieving all users:")
        Dim allUsers = userRepository.GetAll()
        For Each user In allUsers
            Console.WriteLine($"UserID: {user.userId}, Name: {user.uFName} {user.uMName} {user.uLName}, Email: {user.userEmail}, RoleID: {user.roleID}")
        Next

        '' Add a new user
        'Console.WriteLine(vbCrLf & "Adding a new user:")
        'userRepository.AddVerificator("John", "Doe", "john.doe@example.com", "password123", "password123")

        '' Retrieve all users after addition
        'Console.WriteLine(vbCrLf & "Retrieving all users after addition:")
        'allUsers = userRepository.GetAll()
        'For Each user In allUsers
        '    Console.WriteLine($"UserID: {user.userId}, Name: {user.uFName} {user.uMName} {user.uLName}, Email: {user.userEmail}, RoleID: {user.roleID}")
        'Next

        '' Update an existing user
        'Console.WriteLine(vbCrLf & "Updating an existing user:")
        'userRepository.EditVerificator("Jane", "Smith", "jane.smith@example.com", 1)

        '' Retrieve all users after update
        'Console.WriteLine(vbCrLf & "Retrieving all users after update:")
        'allUsers = userRepository.GetAll()
        'For Each user In allUsers
        '    Console.WriteLine($"UserID: {user.userId}, Name: {user.uFName} {user.uMName} {user.uLName}, Email: {user.userEmail}, RoleID: {user.roleID}")
        'Next

        '' Delete a user
        'Console.WriteLine(vbCrLf & "Deleting a user:")
        'userRepository.DeleteVerificator(1)

        '' Retrieve all users after deletion
        'Console.WriteLine(vbCrLf & "Retrieving all users after deletion:")
        'allUsers = userRepository.GetAll()
        'For Each user In allUsers
        '    Console.WriteLine($"UserID: {user.userId}, Name: {user.uFName} {user.uMName} {user.uLName}, Email: {user.userEmail}, RoleID: {user.roleID}")
        'Next

        'Console.WriteLine(vbCrLf & "Press any key to exit...")
        'Console.ReadKey()
    End Sub
End Module

