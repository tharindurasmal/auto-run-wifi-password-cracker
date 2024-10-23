Dim shell, outputFile, fso, wifiProfiles, profileName, commandResult

' Create a FileSystemObject
Set fso = CreateObject("Scripting.FileSystemObject")

' Create a shell object to run system commands
Set shell = CreateObject("WScript.Shell")

' Create a text file to save the WiFi names and passwords
Set outputFile = fso.CreateTextFile("wifi_passwords.txt", True)

' Run command to get all Wi-Fi profiles
Set wifiProfiles = shell.Exec("netsh wlan show profiles")
Do While Not wifiProfiles.StdOut.AtEndOfStream
    profileName = wifiProfiles.StdOut.ReadLine
    ' Look for lines that contain "All User Profile"
    If InStr(profileName, "All User Profile") Then
        ' Extract the Wi-Fi profile name
        profileName = Trim(Split(profileName, ":")(1))
        
        ' Get the Wi-Fi password for this profile
        Set commandResult = shell.Exec("netsh wlan show profile name=""" & profileName & """ key=clear")
        
        ' Write profile name to text file
        outputFile.WriteLine("Wi-Fi Profile: " & profileName)
        
        ' Variable to check if a password was found
        Dim passwordFound
        passwordFound = False
        
        ' Read the output for each profile's password
        Do While Not commandResult.StdOut.AtEndOfStream
            Dim line
            line = commandResult.StdOut.ReadLine
            ' Look for the key content (password)
            If InStr(line, "Key Content") Then
                outputFile.WriteLine("Password: " & Trim(Split(line, ":")(1)))
                passwordFound = True
                Exit Do
            End If
        Loop
        
        ' If no password found, indicate that
        If Not passwordFound Then
            outputFile.WriteLine("Password: None")
        End If
        
        outputFile.WriteLine "-----------------------------------"
    End If
Loop

' Close the text file
outputFile.Close

