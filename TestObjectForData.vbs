Function TestObjectForData(ByVal objToCheck)
    'region TestObjectForDataFunctionMetadata
    '##########################################################################################
    ' Checks an object or variable to see if it "has data".
    ' If any of the following are true, then objToCheck is regarded as NOT having data:
    '   VarType(objToCheck) = 10
    '   objToCheck Is Nothing
    '   IsEmpty(objToCheck)
    '   IsNull(objToCheck)
    '   objToCheck = ""
    ' In any of these cases, the function returns False. Otherwise, it returns True.

    ' Version: 1.0.20201222.0
    '##########################################################################################
    'endregion TestObjectForDataFunctionMetadata

    'region License
    '##########################################################################################
    ' Copyright 2020 Frank Lesniak
    '
    ' Permission is hereby granted, free of charge, to any person obtaining a copy of this
    ' software and associated documentation files (the "Software"), to deal in the Software
    ' without restriction, including without limitation the rights to use, copy, modify, merge,
    ' publish, distribute, sublicense, and/or sell copies of the Software, and to permit
    ' persons to whom the Software is furnished to do so, subject to the following conditions:
    '
    ' The above copyright notice and this permission notice shall be included in all copies or
    ' substantial portions of the Software.
    '
    ' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED,
    ' INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR
    ' PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE
    ' FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR
    ' OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER
    ' DEALINGS IN THE SOFTWARE.
    '##########################################################################################
    'endregion License

    'region DownloadLocationNotice
    '##########################################################################################
    ' The most up-to-date version of this script can be found on the author's GitHub repository
    ' at https://github.com/franklesniak/Test_Object_For_Data
    '##########################################################################################
    'endregion DownloadLocationNotice

    Dim boolTestResult
    Dim boolFunctionReturn

    boolFunctionReturn = True

    'Check VarType(objToCheck) = 10
    On Error Resume Next
    boolTestResult = (VarType(objToCheck) = 10)
    If Err Then
        'Error occurred
        Err.Clear
        On Error GoTo 0
    Else
        'No Error
        On Error GoTo 0
        If boolTestResult = True Then
            'No data
            boolFunctionReturn = False
        End If
    End If

    'Check to see if objToCheck Is Nothing
    If boolFunctionReturn = True Then
        On Error Resume Next
        boolTestResult = (objToCheck Is Nothing)
        If Err Then
            'Error occurred
            Err.Clear
            On Error GoTo 0
        Else
            'No Error
            On Error GoTo 0
            If boolTestResult = True Then
                'No data
                boolFunctionReturn = False
            End If
        End If
    End If

    'Check IsEmpty(objToCheck)
    If boolFunctionReturn = True Then
        On Error Resume Next
        boolTestResult = IsEmpty(objToCheck)
        If Err Then
            'Error occurred
            Err.Clear
            On Error GoTo 0
        Else
            'No Error
            On Error GoTo 0
            If boolTestResult = True Then
                'No data
                boolFunctionReturn = False
            End If
        End If
    End If

    'Check IsNull(objToCheck)
    If boolFunctionReturn = True Then
        On Error Resume Next
        boolTestResult = IsNull(objToCheck)
        If Err Then
            'Error occurred
            Err.Clear
            On Error GoTo 0
        Else
            'No Error
            On Error GoTo 0
            If boolTestResult = True Then
                'No data
                boolFunctionReturn = False
            End If
        End If
    End If

    'Check objToCheck = ""
    If boolFunctionReturn = True Then
        On Error Resume Next
        boolTestResult = (objToCheck = "")
        If Err Then
            'Error occurred
            Err.Clear
            On Error GoTo 0
        Else
            'No Error
            On Error GoTo 0
            If boolTestResult = True Then
                'No data
                boolFunctionReturn = False
            End If
        End If
    End If

    TestObjectForData = boolFunctionReturn
End Function
