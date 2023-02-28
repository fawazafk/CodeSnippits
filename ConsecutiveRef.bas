Attribute VB_Name = "NewMacros"
Sub ConsecutiveRef()

    '

    ' Fawaz: Reorganize consecutive numbers with hyphens

    '
    Dim regexOne, regexMatch                    As Object
    Dim theMatches                              As Object
    Dim Match                                   As Object
    Set regexOne = CreateObject("VBScript.RegExp")
    Set regexMatch = CreateObject("VBScript.RegExp")
    Application.ScreenUpdating = False
    Dim i                                       As Long
    Dim foundText, newText                      As String
    Dim begining_ind, ending_ind, unfinished      As Boolean
    Dim begining, ending, begining_i, ending_i  As Integer
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    regexOne.Pattern = "([0-9]{1,}, )+[0-9]{1,}"
    regexOne.Global = True
    regexOne.IgnoreCase = True
    regexMatch.Global = True
    regexMatch.IgnoreCase = True
    unfinished = True
     
    ' Looping untill no changes are needed on the string
    While unfinished
    Set theMatches = regexOne.Execute(Selection.Text)
        For Each Match In theMatches
            foundText = Replace(Match.Value, " ", "")
            foundText = Replace(Replace(foundText, Chr(10), ""), Chr(13), "")
            myArray = Split(foundText, ",")
            
            ' Reset Values
            begining_i = 0
            ending_i = 0
            begining_ind = False
            ending_ind = False
            newText = ""
                 
 
            ' Loop inside the array to find begining/ending
            For i = 0 To UBound(myArray) - 1
                If CInt(myArray(i)) = CInt(myArray(i + 1)) - 1 And begining_ind = False Then        'Check for beining number
                    begining_ind = True
                    ending_ind = False
                    begining_i = i
                    begining = CInt(myArray(i))
                    If (i + 1 = UBound(myArray)) Then        'Check if we reach the end of array if array size is 2
                        begining_ind = False
                        ending_ind = True
                        ending_i = i + 1
                        ending = CInt(myArray(i + 1))
                    End If
                Else
                    'Check for ending number in middle of array
                    If (CInt(myArray(i)) <> CInt(myArray(i + 1)) - 1) And ending_ind = False And begining_ind = True Then
                        begining_ind = False
                        ending_ind = True
                        ending_i = i
                        ending = CInt(myArray(i))
                    'Check if we reach the end of array and set ending as last number in array
                    ElseIf (i + 1 = UBound(myArray)) And ending_ind = False And begining_ind = True Then
                        begining_ind = False
                        ending_ind = True
                        ending_i = i + 1
                        ending = CInt(myArray(i + 1))
                    End If
                End If
            Next i
            
            'Reform the array to string: newText
            If ending_i > 0 Then
                For i = 0 To begining_i - 1
                    newText = newText & myArray(i) & ", "
                Next i
                newText = newText & begining & ChrW(8211) & ending & ", "
            Else
                newText = newText & myArray(0) & ", "
            End If
            If ending_i < UBound(myArray) Then
                For i = ending_i + 1 To UBound(myArray)
                    newText = newText & myArray(i) & ", "
                Next i
            End If
     
            ' Replace within Selection.Text using a second pattern
            newText2 = Left(newText, Len(newText) - 2)
            regexMatch.Pattern = Match
            Selection.Text = regexMatch.Replace(Selection.Text, newText2)
        Next Match
       
        If newText2 = newText3 Then
            unfinished = False
        End If
        newText3 = newText2
    Wend
    
End Sub
