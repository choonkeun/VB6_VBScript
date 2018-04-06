<!DOCTYPE html>
<html>
<body>
<%
'http://Home/Test_VBScript.asp

    Dim arr(5)
    arr(0) = "1"           'Number as String
    arr(1) = "VBScript"    'String
    arr(2) = 100                            'Number
    arr(3) = 2.45                          'Decimal Number
    arr(4) = #10/07/2013#  'Date
    arr(5) = #12.45 PM#    'Time

    a = 5 \ 2
    response.write("a = 5 \ 2 -> " & CStr(a) & "<br />")

    b = 5 / 2
    response.write("a = 5 / 2 -> " & CStr(b) & "<br />")

    response.write("123.999 Int Value : " & CStr(Int(123.999)) & "<br />")
    response.write("123.999 Fix Value : " & CStr(Fix(123.999)) & "<br /><br />")
    
    'array filter
    response.write("display weeknames that has NOT 'S' <br />")
    a=Array("Sunday","Monday","Tuesday","Wednesday","Thursday","Friday","Saturday")
    b=Filter(a,"S",False)   'does not have "S"
    for each x in b
        response.write(x & "<br />")
    next

    response.write("display weeknames that has 'S' <br />")
    a=Array("Sunday","Monday","Tuesday","Wednesday","Thursday","Friday","Saturday")
    b=Filter(a,"S")   'have "S"
    for each x in b
        response.write(x & "<br />")
    next

    response.write("display weeknames that has 's' <br />")
    a=Array("Sunday","Monday","Tuesday","Wednesday","Thursday","Friday","Saturday")
    b=Filter(a,"s")   'have "s"
    for each x in b
        response.write(x & "<br />")
    next

    On Error Resume Next
    Err.Raise 6 ' Raise an overflow error.
    response.write("Error # " & CStr(Err.Number) & " " & Err.Description & "<br />")
    Err.Clear ' Clear the error

    'regEx
    Dim regEx, Match, Matches, RetStr, srExp 'Create variable.
    strExp = "Is1 is2 IS3 is4"  'Input string
    Set regEx = New RegExp      'Instantiate RegExp object
    regEx.Pattern = "is."       'Set pattern.
    regEx.IgnoreCase = True     'Set case insensitivity.
    regEx.Global = True         'Set global applicability.
    Set Matches = regEx.Execute(strExp) 'Execute search.
    RetStr = ""                 'Zero out string
    For Each Match in Matches   'Iterate Matches collection.
        RetStr = RetStr & "Match found at position "
        RetStr = RetStr & Match.FirstIndex & ". Match Value is '"
        RetStr = RetStr & Match.Value & "'." & "<br />"
    Next
    'MsgBox RetSt
    response.write("<br />")
    response.write("Data: " + strExp + "<br />")
    response.write("Pattern: " + regEx.Pattern + "<br />")
    response.write("Count: " + CStr(Matches.Count) + "<br />")
    response.write("Found: " + RetStr + "<br />")

    Set oMatch = Matches(0)
    MatStr = "subMatch(0): " + oMatch.SubMatches(0)
    MatStr = MatStr & ", subMatch(1): " + oMatch.SubMatches(1) & "<br />"
    response.write("SubMatches: " + MatStr + "<br />")

    Set fso = CreateObject("Scripting.FileSystemObject")
    WinFolder = fso.GetSpecialFolder(0) & "\" 'Result is "C:\Windows\"
    SysFolder = fso.GetSpecialFolder(1) & "\" 'Result is "C:\Windows\system32\"
    response.write("WinFolder: " + WinFolder + "<br />")
    response.write("SysFolder: " + SysFolder + "<br />")

    Set fso = CreateObject("Scripting.FileSystemObject") 'Instantiate the FSO object
    drvPath = fso.GetSpecialFolder(1)
    Set f = fso.GetFolder(drvPath)  'Instantiate the parent folder object
    Set fc = f.SubFolders           'Return the subfolder Folders collection
    response.write("<br />")
    response.write("Recursive SubFolders: " + CStr(fc.Count) + "<br />")
    ShowSubfolders FSO.GetFolder(drvPath), 4        '4 is sub level depth

    myItem = fc.Item ("appmgmt")    'displays the entire path to the Web subfolder
    response.write("appmgmt path: " + myItem + "<br />")

    Set fso = CreateObject("Scripting.FileSystemObject") 'Instantiate the FSO object
    drvPath = fso.GetSpecialFolder(0)
    Set f = fso.GetFolder(drvPath)
    fc = f.Subfolders                       'Returns collection of (sub)folders
    s = ""
    For each f1 in fc
        s = s & "SubFolders: " &f1.path & "<br />"
    Next
    response.write("Recursive SubFolders: " + s + "<br />")

    
    Sub ShowSubFolders(Folder, Depth)
        If Depth > 0 then
            For Each Subfolder in Folder.SubFolders
                'Wscript.Echo Subfolder.Path
                response.write("SubFolders: " + Subfolder.Path + "<br />")
                ShowSubFolders Subfolder, Depth -1
            Next
        End if
    End Sub


%>  
</body>
</html>