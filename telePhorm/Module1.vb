Imports System.IO
Imports System.Xml

Module Module1

    '-dir d:\priority -user tabula -pass Tabula! -clid 0876120183 -xml http://mail.atp.ie:8080

    Private arg As clArg
    Private doc As XmlDocument

    Sub Main(ByVal args As String())

        doc = New XmlDocument
        arg = New clArg(args)

        Dim clid As String = String.Empty

        Try
            With arg.Keys
                If .Contains("?") Or .Count = 0 Then
                    Console.Write(My.Resources.help.Replace("$build$", _
                        String.Format("{0}.{1}.{2}", _
                            My.Application.Info.Version.Major, _
                            My.Application.Info.Version.Minor, _
                            My.Application.Info.Version.Build _
                        ) _
                    ) _
                )

                Else
                    For Each a As String In arg.Keys
                        Select Case a.ToLower
                            Case "user", "usr", "u"
                                My.Settings.PRIUSER = arg(a)
                            Case "password", "pass", "pas", "pw", "pwd", "p"
                                My.Settings.PRIPASSWORD = arg(a)
                            Case "feed", "srv", "svr", "serv", "server", "xml", "x"
                                My.Settings.FEEDURL = arg(a)
                            Case "clid", "clsid", "callerid", "callid", "id", "num", "c"
                                For i As Integer = 0 To arg(a).Length - 1
                                    Select Case arg(a).Substring(i, 1)
                                        Case "1", "2", "3", "4", "5", "6", "7", "8", "9", "0"
                                            clid &= arg(a).Substring(i, 1)
                                    End Select
                                Next
                        End Select
                    Next
                    My.Settings.Save()

                    If clid.Length > 0 Then

                        If My.Settings.FEEDURL.Length = 0 Then Throw New Exception(String.Format("Server not specified.", ""))
                        If My.Settings.PRIUSER.Length = 0 Then Throw New Exception(String.Format("Priority User not specified.", ""))
                        If My.Settings.PRIPASSWORD.Length = 0 Then Throw New Exception(String.Format("Priority Password not specified.", ""))

                        doc.Load( _
                            String.Format("{0}/phone.ashx?num={1}", _
                                My.Settings.FEEDURL, _
                                clid _
                            ) _
                        )

                        Dim result As XmlNode = doc.SelectSingleNode("phone/result")
                        If IsNothing(result) Then
                            MsgBox( _
                                String.Format( _
                                    "Unknown caller: {0}", _
                                    clid _
                                ) _
                            )
                        Else
                            RunCommand( _
                                result.SelectSingleNode("ENV").InnerText, _
                                result.SelectSingleNode("FORM").InnerText, _
                                result.SelectSingleNode("ID").InnerText _
                            )
                        End If

                    End If

                End If

            End With

        Catch EX As Exception
            Console.WriteLine(EX.Message)
        End Try

    End Sub

    Private Sub RunCommand(ByVal ENV As String, ByVal Form As String, ByVal ID As String)

        Dim sOutput As String = ""
        Dim sErrs As String = ""
        Dim myProcess As Process = New Process()
        Dim cmd As String = ""

        Try
            With myProcess
                With .StartInfo
                    .FileName = "cmd.exe"
                    .UseShellExecute = False
                    .CreateNoWindow = True
                    .RedirectStandardInput = True
                    .RedirectStandardOutput = True
                    .RedirectStandardError = True
                End With
                .Start()

                Dim sIn As StreamWriter = myProcess.StandardInput
                Dim sOut As StreamReader = myProcess.StandardOutput
                Dim sErr As StreamReader = myProcess.StandardError

                Dim path As New DirectoryInfo(System.IO.Path.GetDirectoryName( _
                            System.Reflection.Assembly.GetExecutingAssembly().CodeBase _
                        ).Replace("file:\", "") _
                    )
                With sIn
                    cmd = String.Format( _
                        "{0}{1}\WINRUN.exe{0} {0}{0} {3} {4} {0}{7}\system\prep{0} {5} WINFORM {2} {0}{0} {0}{6}{0} {0}{0} 2", _
                        Chr(34), _
                        path.FullName, _
                        Form, _
                        My.Settings.PRIUSER, _
                        My.Settings.PRIPASSWORD, _
                        ENV, _
                        ID, _
                        path.Parent.FullName _
                    )

                    cmd = Replace(cmd, "\\", "\")

                    'Console.WriteLine(cmd)

                    .AutoFlush = True
                    .Write(cmd & _
                        System.Environment.NewLine)
                    .Write("exit" & _
                        System.Environment.NewLine)
                    .Close()

                End With

                Dim l As Integer = 0
                Do Until l = 100
                    If sOut.Peek <> 0 Then
                        sOutput = sOutput + sOut.ReadLine
                    End If
                    l = l + 1
                    Threading.Thread.Sleep(1)
                Loop

                If Len(sOutput) > 0 Then
                    sOutput = sOut.ReadToEnd
                    sOut.Close()
                    sErrs = sErr.ReadToEnd()
                    sErr.Close()
                End If

                If Not myProcess.HasExited Then
                    myProcess.Kill()
                End If

                .Close()

                If sErrs.Length > 0 Then
                    MsgBox(sErrs & cmd)
                End If

            End With
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

End Module
