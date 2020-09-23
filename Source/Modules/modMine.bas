Attribute VB_Name = "modMine"

Option Explicit


Public rel_path As String  'relative path to our executable


Public Function GetFuncName(strDecl As String) As String
    Dim pos As Integer
    
    If InStr(strDecl, "(") > 0 Then
        pos = InStr(strDecl, "(")
        GetFuncName = Mid(strDecl, 1, pos - 1)
    Else
        GetFuncName = strDecl
    End If
End Function


Public Sub FindPublics(strFilename As String)
    Dim fnum As Integer, pos As Integer, pos2 As Integer
    Dim one_line As String, additional_line As String
    Dim head As String, body As String, tail As String
    Dim ItmX As ListItem
    Dim continued As Boolean

    fnum = FreeFile
    Open strFilename For Input As fnum

    Do While Not EOF(fnum)
        ' Read a line.
        Line Input #fnum, one_line

        ' check line for the prefix
        If Left$(LTrim(one_line), 6) = "Public" _
            Or Left$(LTrim(one_line), 3) = "Sub" And Mid$(LTrim(one_line), 4, 1) = " " _
            Or Left$(LTrim(one_line), 8) = "Function" Then
            ' Positive. Check if multi-lines...
            If Trim$(Right$(one_line, 1)) = "_" Then
                continued = True
                additional_line = LTrim$(Left$(one_line, Len(one_line) - 1))
            Else 'if not, we have a complete line
                Set ItmX = Form1.lvwPublic.ListItems.Add()
                'remove "Public" prefixes
                one_line = IIf(Left$(LTrim(one_line), 6) = "Public", _
                    Mid(one_line, Len("Public ") + 1, Len(one_line)), _
                    IIf(Left$(LTrim(one_line), 3) = "Sub", Mid(one_line, Len("Sub ") + 1), _
                    Mid(one_line, Len("Function ") + 1)))
                ItmX.Text = RemovePrefx(one_line)
                'set plural labels as required...
                If Form1.lvwPublic.ListItems.Count > 1 Then _
                    Form1.Label3.Caption = "Publics"
                continued = False
            End If
        Else 'we append the continuation here...
            If continued = True Then 'there's more...
                If Trim$(Right$(one_line, 1)) = "_" Then
                    'cat this line and get more
                    additional_line = additional_line _
                    & LTrim$(Left$(one_line, Len(one_line) - 1))
                Else 'we have the last remaining line now
                    additional_line = additional_line _
                    & LTrim$(Left$(one_line, Len(one_line)))
                    
                    'we remove ugly spaces after "(" and before ")"
                    'which vb doesnt like
                    If InStr(additional_line, "(") > 0 Then
                        pos = InStr(additional_line, "(")
                        head = Left(additional_line, pos)
                        
                        'remove "Public" prefixes
                        head = Mid(head, Len("Public ") + 1, Len(head))
                    End If
                    
                    If InStr(additional_line, ")") > 0 Then
                        pos2 = InStr(additional_line, ")")
                        tail = Mid(additional_line, pos2, (Len(additional_line) - pos2) + 1)
                    End If
                    
                    body = Mid(additional_line, pos + 2, _
                    Len(additional_line) - (Len(tail) + pos + 2))
                    
                    Set ItmX = Form1.lvwPublic.ListItems.Add()
                    ItmX.Text = RemovePrefx(head & body & tail)
                    If Form1.lvwPublic.ListItems.Count > 1 Then _
                        Form1.Label3.Caption = "Publics"
                    continued = False
                End If
            End If
        End If
    Loop

    Close fnum
End Sub


Public Sub FindModules(strFilename As String)
    Dim fnum As Integer
    Dim one_line As String
    Dim pos As Integer, pos2 As Integer
    Dim ItmX As ListItem

    fnum = FreeFile
    Open strFilename For Input As fnum

    Do While Not EOF(fnum)
        ' Read a line.
        Line Input #fnum, one_line

        ' See what it is.
        If InStr(one_line, "=") > 0 Then
            pos = InStr(one_line, "=")
            If (Left$(UCase$(one_line), pos - 1) = "MODULE") Then
                If (Right$(UCase$(one_line), 3) = "BAS") Then
                    If InStrRev(one_line, "\", -1, vbTextCompare) >= 0 Then
                        pos = InStrRev(one_line, "\")
                    
                        'save parent folder, if any, relative to executable path
                        pos2 = InStr(one_line, ";")
                        
                        If pos > 0 Then
                            rel_path = Mid(one_line, pos2 + 2, (pos - 1) - pos2)
                        Else
                            rel_path = ""
                        End If
                    
                        Set ItmX = Form1.lvwModules.ListItems.Add(1, , , , 1)
                        If pos > 0 Then
                            ItmX.Text = (Mid$(one_line, pos + 1))
                        Else
                            ItmX.Text = (Mid$(one_line, pos2 + 2, Len(one_line)))
                        End If
                        'save parent folder
                        ItmX.Tag = rel_path
                        
                        'We don't need the ext here, so remove it
                        ItmX.Text = Left(ItmX.Text, Len(ItmX.Text) - 4)
                        If Form1.lvwModules.ListItems.Count > 1 Then _
                            Form1.Label2.Caption = "Modules"
                    End If
                End If
            End If
        End If
    Loop

    Close fnum
End Sub


Public Function prjTitle(strFilename As String) As String
    Dim fnum As Integer
    Dim one_line As String
    Dim pos As Integer

    fnum = FreeFile
    Open strFilename For Input As fnum

    Do While Not EOF(fnum)
        ' Read a line.
        Line Input #fnum, one_line

        ' See what it is.
        If InStr(one_line, "=") > 0 Then
            pos = InStr(one_line, "=")
            If (Left$(UCase$(one_line), pos - 1) = "NAME") Then
                one_line = Mid$(one_line, pos + 2)
                prjTitle = Left$(one_line, Len(one_line) - 1)
                Exit Do
            End If
        End If
    Loop

    Close fnum
End Function


Public Function ProjectDesc(strFilename As String) As String
    Dim fnum As Integer
    Dim one_line As String
    Dim pos As Integer

    fnum = FreeFile
    Open strFilename For Input As fnum

    Do While Not EOF(fnum)
        ' Read a line.
        Line Input #fnum, one_line

        ' See what it is.
        If InStr(one_line, "=") > 0 Then
            pos = InStr(one_line, "=")
            If (Left$(UCase$(one_line), pos - 1) = "DESCRIPTION") Then _
                ProjectDesc = Mid$(one_line, pos + 1, Len(one_line) - (pos))
        End If
    Loop

    Close fnum
End Function


Public Function ProjectType(strFilename As String) As Boolean
    Dim fnum As Integer
    Dim one_line As String
    Dim pos As Integer

    fnum = FreeFile
    Open strFilename For Input As fnum

    Do While Not EOF(fnum)
        ' Read a line.
        Line Input #fnum, one_line

        ' See what it is.
        If InStr(one_line, "=") > 0 Then
            pos = InStr(one_line, "=")
            If (Left$(UCase$(one_line), pos - 1) = "TYPE") Then
                If (Mid$(UCase$(one_line), pos + 1) = "OLEDLL") Then
                    ProjectType = True
                    Exit Do
                End If
            Else
                PopupBalloon Form1, "Unable to load file. Please verify that it is a" & _
                vbCrLf & "valid ActiveX DLL project.", _
                Form1.Caption, tipIconError
                
                ProjectType = False
                Exit Do
            End If
        End If
    Loop

    Close fnum
End Function


Private Function RemovePrefx(strString As String) As String
    If Left(strString, Len("Sub")) = "Sub" Then
        RemovePrefx = Mid(strString, Len("Sub") + 2, Len(strString) - (Len("Sub") + 1))
    ElseIf Left(strString, Len("Const")) = "Const" Then
        RemovePrefx = Mid(strString, Len("Const") + 2, Len(strString) - (Len("Const") + 1))
    ElseIf Left(strString, Len("Type")) = "Type" Then
        RemovePrefx = Mid(strString, Len("Type") + 2, Len(strString) - (Len("Type") + 1))
    ElseIf Left(strString, Len("Enum")) = "Enum" Then
        RemovePrefx = Mid(strString, Len("Enum") + 2, Len(strString) - (Len("Enum") + 1))
    ElseIf Left(strString, Len("Function")) = "Function" Then
        RemovePrefx = Mid(strString, Len("Function") + 2, Len(strString) - (Len("Function") + 1))
    ElseIf Left(strString, Len("Property Get")) = "Property Get" Then
        RemovePrefx = Mid(strString, Len("Property Get") + 2, Len(strString) - (Len("Property Get") + 1))
    ElseIf Left(strString, Len("Property Let")) = "Property Let" Then
        RemovePrefx = Mid(strString, Len("Property Let") + 2, Len(strString) - (Len("Property Let") + 1))
    ElseIf Left(strString, Len("Property Set")) = "Property Set" Then
        RemovePrefx = Mid(strString, Len("Property Set") + 2, Len(strString) - (Len("Property Set") + 1))
    Else
        RemovePrefx = strString
    End If
End Function


Public Function ReverseString(ByVal InputString As String) As String
    Dim lLen As Long, lCtr As Long
    Dim sChar As String
    Dim sAns As String

    lLen = Len(InputString)
    For lCtr = lLen To 1 Step -1
        sChar = Mid(InputString, lCtr, 1)
        sAns = sAns & sChar
    Next

    ReverseString = sAns
End Function

