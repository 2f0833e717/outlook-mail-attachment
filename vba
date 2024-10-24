
'----------------------------------------------------------------'
'自動貼付ファイル保存マクロ'
'----------------------------------------------------------------'
Private Sub Application_NewMailEx(ByVal EntryIDCollection As String)

    Debug.Print ""
    Debug.Print "自動貼付ファイル保存マクロ start " & Format(Now, "yyyy/mm/dd hh:mm:ss")
    On Error GoTo ErrorHandler

    '添付ファイルを保存する親フォルダパスを指定'
    Dim folderPath As String
    folderPath = "C:Users\xxxx\Downloads\"

    '7-zipのパスを指定'
    Dim str7ZipPath As String
    str7ZipPath = """C:\Program Files\7-Zip\7z.exe"""

    '特定の件名を読み取る'
    Dim fsoSubject As Object
    Set fsoSubject = CreateObject("Scripting.FileSystemObject")
    Dim fileSubject As Object
    Set fileSubject = fsoSubject.OpenTextFile("C:Users\xxxx\outlook-attach-settings-subjects.txt",1)
    Dim subjects() As String
    Dim subjectsLineCount As Integer
    subjectsLineCount = 0
    Do Until fileSubject.AtEndOfStream
      ReDim Preserve subjects(subjectsLineCount)
      subjects(subjectsLineCount) = fileSubject.ReadLine
      subjectsLineCount = subjectsLineCount + 1
    Loop
    fileSubject.Close

    '特定の宛先を読み取る'
    Dim fsoRecipient As Object
    Set fsoRecipient = CreateObject("Scripting.FileSystemObject")
    Dim fileRecipient As Object
    Set fileRecipient = fsoRecipient.OpenTextFile("C:\Users\xxxx\outlook-attach-settings-recipients.txt",1)
    Dim recipients() As String
    Dim recipientsLineCount As Integer
    recipientsLineCount = 0
    Do Until fileRecipient.AtEndOfStream
      ReDim Preserve recipients(recipientsLineCount)
      recipients(recipientsLineCount) = fileRecipient.ReadLine
      recipientsLineCount = recipientsLineCount + 1
    Loop
    fileRecipient.Close

    'パスワードをpassword.txtから読み取る'
    Dim strPassword As String
    strPassword = CreateObject("Scripting.FileSystemObject").OpenTextFile("C:\Users\xxxx\outlook-attach-settings-password.txt").ReadLine
  
    'error.txtとsuccess.txtのパスを指定'
    Dim objFSO As Object
    Dim objTextFile As Object
    Dim strErrorLog As String
    Dim strSuccessLog As String
    strErrorLog = "C:\Users\xxxx\outlook-attach-log-error.txt"
    strSuccessLog = "C:\Users\xxxx\outlook-attach-log-success.txt"

    '複数のメールを同時受信した際は、複数のIDがカンマ区切りで渡されるため、カンマ区切りでIDを配列に格納'
    Dim entryIDs As Variant
    entryIDs = Split(EntryIDCollection, ",")

    Dim i As Long
    For i = 0 To UBound(entryIDs)
        '受信したメールを取得'
        Dim objMsg As Object
        Set objMsg = Application.Session.GetItemFromID(entryIDs)
        Debug.Print "objMsg.Subject :" & objMsg.Subject
        Debug.Print "objMsg.SenderName :" & objMsg.SenderName

        Dim j As Integer
        For j = 0 To UBound(recipients)
            '特定の宛先からのメールだけにマクロを適用する場合'
            If InStr(objMsg.SenderName, recipients(j)) > 0 Then
                Debug.Print recipients(j)
                Dim k As Integer
                For k = 0 To UBound(subjects)
                    '特定の懸命が含まれるメールだけにマクロを適用する場合'
                    'If (InStr(objMsg.Subject, strSubject1) > 0 _'
                    'Or InStr(objMsg.Subject, strSubject2) > 0 _'
                    'Or InStr(objMsg.Subject, strSubject3) > 0) Then'
                    'If InStr(objMsg.Subject, subjects(k)) > 0 Then'
                    If 0 = 0 Then
                        'Debug.Print subjects(k)'
                        Debug.Print "subjects success"
          
                        '受信日時を取得してタイトルように文字列変換'
                        Dim recTime As String
                        recTime = Format(objMsg.ReceivedTime, "yyyymmdd-hhmm_")

                        '受信日時yyyymmddでサブフォルダ名を用意'
                        Dim subfolder As String
                        subfolder = Format(objMsg.ReceivedTime, "yyyymmdd")

                        'Dir関数で親フォルダの中にサブフォルダの有無を確認'
                        Dim frag As String
                        frag = Dir(folderPath & subfolder, vbDirectory)

                        '親フォルダの中に不サブフォルダがない場合（すでにサブフォルダがある場合は何もしない）'
                        If frag = "" Then
                            'サブフォルダを作成'
                            MkDir folderPath & subfolder
                        End If

                        '受信したメールに添付されたファイルごとに処理'
                        Dim objAttach As attachment
                        For Each objAttach In objMsg.Attachments
                            Dim attachFileName As String
                            attachFileName = folderPath & subfolder & "\" & recTime & objAttach.fileName
                          
                            '添付ファイルを保存'
                            objAttach.SaveAsFile attachFileName

                            '添付ファイルがzipの場合解凍'
                            Debug.Print "自動貼付ファイル保存マクロ zip解凍"
                            If Right(attachFileName, 4) = ".zip" Then
                                Dim objShell As Object
                                Set objShell = CreateObject("WScript.Shell")

                                '7-Zipを使用してパスワード付きのzipファイルを解凍'
                                Dim intReturn As Integer
                                Debug.Print str7ZipPath & _
                                "x " & _
                                folderPath & _
                                subfolder & _
                                " -p " & _
                                strPassword & _
                                " " & _
                                attachFileName

                                intReturn = objShell.Run(str7ZipPath & _
                                "x " & _
                                folderPath & _
                                subfolder & _
                                " -p " & _
                                strPassword & _
                                " " & _
                                attachFileName, 0, True)

                                '処理が成功した場合は、特定のPathにあるsuccess.txt二処理した結果の概要を出力'
                                If intReturn = 0 Then
                                    Set objFSO = CreateObject("Scripting.FileSystemObject")
                                    Set objTextFile = objFSO.OpenTextFile(strSuccessLog, 8, True)
                                    objTextFile.WriteLine Format(Now, "yyyymmdd-hh:mm:ss") & "-" & objMsg.Subject & ", " & objAttach.fileName & ", " & objMsg.SenderName & ": The file " & attachFileName & " was successfully extracted."
                                    objTextFile.Close
                                    Set objFSO = Nothing
                                    Set objTextFile = Nothing
                                '処理が失敗した場合は、特定のPathにあるerror.txtにエラーを出力'
                                Else
                                    Set objFSO = CreateObject("Scripting.FileSystemObject")
                                    Set objTextFile = objFSO.OpenTextFile(strErrorLog, 8, True)
                                    objTextFile.WriteLine Format(Now, "yyyymmdd-hh:mm:ss") & "-" & objMsg.Subject & ", " & objAttach.fileName & ", " & objMsg.SenderName & ": The file " & attachFileName & " could not be extracted."
                                    objTextFile.Close
                                    Set objFSO = Nothing
                                    Set objTextFile = Nothing
                                End If

                                Set objShell = Nothing
                            End If
                        Next
                    End If
                Next k
            End If
        Next j
    Next i
    
    Set objMsg = Nothing

    'Outlookマクロの実行結果をsuccess.txtに出力'
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set objTextFile = objFSO.OpenTextFile(strSuccessLog, 8, True)
    objTextFile.WriteLine Format(Now, "yyyymmdd-hh:mm:ss") & ": Outlook macro executed successfully."
    objTextFile.Close
    Set objFSO = Nothing
    Set objTextFile = Nothing

    Exit Sub

ErrorHandler:
    'エラーが発生した場合は、特定のPathにあるerror.txtにエラーを出力'
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set objTextFile = objFSO.OpenTextFile(strErrorLog, 8, True)
    objTextFile.WriteLine Format(Now, "yyyymmdd-hh:mm:ss") & ": An error occurred: " & err.Description
    objTextFile.Close
    Set objFSO = Nothing
    Set objTextFile = Nothing
End Sub

'----------------------------------------------------------------'
'C:Users\xxxx\outlook-attach-settings-subjects.txt等のファイルは全て改行区切り'
'例)'
'1111@gmail.com'
'2222@gmail.com'
'3333@gmail.com'

'----------------------------------------------------------------'
