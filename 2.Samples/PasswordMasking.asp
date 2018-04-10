<!DOCTYPE html>
<html>
<body>
<%
'http://Home/PasswordMasking.asp

    tmpStr = "Provider=SQLNCLI10.1;User ID=batch;Password=aa""@#$;Data Source=NorthWind;Initial Catalog=facetsxc;gggg=bbbb;dddd=bbggggbb;"
    outValue = PasswordMasking(tmpStr)
    response.write("1" + outValue + "<br />")

    tmpStr = "Password=aa""@#$;Provider=SQLNCLI10.1;User ID=batch;Data Source=NorthWind;Initial Catalog=facetsxc;gggg=bbbb;dddd=bbggggbb;"
    outValue = PasswordMasking(tmpStr)
    response.write("2" + outValue + "<br />")

    tmpStr = "Provider=SQLNCLI10.1;User ID=batch;Data Source=NorthWind;Initial Catalog=facetsxc;gggg=bbbb;dddd=bbggggbb;Password=aa""@#$;"
    outValue = PasswordMasking(tmpStr)
    response.write("3" + outValue + "<br />")

    tmpStr = "Provider=SQLNCLI10.1;User ID=batch;Data Source=NorthWind;Initial Catalog=facetsxc;gggg=bbbb;dddd=bbggggbb;Password=aa""@#$"
    outValue = PasswordMasking(tmpStr)
    response.write("4" + outValue + "<br />")

    'if exist("Password=")
    '    If exist as ("Password=;")
    '        Copy input to Output or Masking Password
    '    Else     ("=XXXXXXX;")
    '        Masking Password
    'Else
    '    Copy input to Output
    '
    Function PasswordMasking(inputStr)
        On Error Resume Next
        Dim strArray, i
        PasswordMasking = inputStr
        strArray = Split(inputStr, ";")
        For i=Lbound(strArray) to Ubound(strArray)
            If InStr(LCASE(strArray(i)), "password=") <> 0 Then
                strArray(i) = "Password=XXXXXX"
                PasswordMasking = Join(strArray, ";")
                Exit Function
            End If
        Next
    End Function

    'Function PasswordMasking(inputStr)
    '    On Error Resume Next
    '    Dim iStart, iEnd, iPassword
    '    PasswordMasking = inputStr
    '    iStart = InStr(1,LCase(inputStr),"password=") + 9
    '    if iStart <= 9 Then Exit Function
    '    iEnd = InStr(iStart,inputStr,";")
    '    if iEnd < 1 Then iEnd = Len(inputStr) + 1
    '    iPassword = Mid(inputStr,iStart,iEnd - iStart)
    '    PasswordMasking = replace(inputStr,iPassword,"XXXXXX")
    'End Function

    'Function MaskConnString(stringToMask)
    '    On Error Resume Next
    '    Dim regexMid, regexEnd, objMatchesMid, objMatchesEnd, lastChar, outMaskedStr
    '
    '    Set regexMid = new RegExp
    '    regexMid.IgnoreCase = true
    '    regexMid.Pattern = "Password=(.+?)[;\s]"
    '
    '    Set regexEnd = new RegExp
    '    regexEnd.IgnoreCase = true
    '    regexEnd.Pattern = "Password=(.+?)$"
    '
    '    if (regexMid.Test(stringToMask)) Then
    '          'wscript.echo "found in front or mid string"
    '          set objMatchesMid = regexMid.Execute(stringToMask)
    '          outMaskedStr = Replace(stringToMask, objMatchesMid(0).SubMatches(0), "XXXXXX")
    '    elseif (regexEnd.Test(stringToMask)) Then
    '          'wscript.echo "found at end of string"
    '          set objMatchesEnd = regexEnd.Execute(stringToMask)
    '          outMaskedStr = Replace(stringToMask, objMatchesEnd(0).SubMatches(0), "XXXXXX")
    '    else
    '          'wscript.echo "not found"
    '          outMaskedStr = stringToMask
    '    end if
    '    MaskConnString = outMaskedStr
    'End Function

%>
</body>
</html>