Option Explicit On
Imports System.Net
Imports Microsoft.Win32.SafeHandles

Public Class Form1
    Dim skylanderBytes(1023) As Byte
    Dim outRepoBytes(32) As Byte
    Dim inRepoBytes(32) As Byte
    Dim skyInfo(2) As String
    Dim skylanderID As Integer
    Dim skylanderGold As Long
    Dim skylanderEXP As Long
    Dim skylanderHat As Integer
    Dim skylanderSkill As Integer
    Dim skylanderSkillbin As String
    Dim skylanderFilepath As String
    Dim exepath As String
    Dim indexVal As Integer
    Dim indexValB As Integer
    Dim moneyA, expA, expB, expC, hatA, hatB, hatC, hatD, skill As Integer
    Dim X
    Dim autom As Boolean
    Dim portalHandle As SafeFileHandle



    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'initialize file dialogs, database connection, hat combobox and CRC calculator table
        OpenFileDialog1.FileName = ""
        SaveFileDialog1.FileName = ""

        exepath = Application.StartupPath & "\"

        hatsimulator(exepath)
        ComboBox1.SelectedIndex = 0
        TextBox3.Text = ComboBox1.Text
        CRCcalculator.inittable()
        lockControls()
        lockPortalControls()
        autom = False
    End Sub

    Private Sub ReadSkylanderToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ReadSkylanderToolStripMenuItem.Click
        'reads skylander data from the portal
        Dim timeout As Integer
        Dim readBlock As Integer

        'reset portal
        outRepoBytes(1) = &H52
        outputReport(portalHandle, outRepoBytes)
        System.Threading.Thread.Sleep(50)

        outRepoBytes(1) = &H41
        outRepoBytes(2) = 1
        outputReport(portalHandle, outRepoBytes)
        System.Threading.Thread.Sleep(500)

        'set to "read first skylander" mode
        outRepoBytes(1) = &H51
        outRepoBytes(2) = &H20
        readBlock = 0
        Do
            'send report and flush hid queue
            outRepoBytes(3) = readBlock
            outputReport(portalHandle, outRepoBytes)
            flushHid(portalHandle)
            timeout = 0
            Do
                'read the reply from the portal, the portal replies between 1 and 2 reports later
                inputReport(portalHandle, inRepoBytes)
                timeout = timeout + 1
            Loop Until inRepoBytes(1) <> &H53 Or timeout = 4

            If timeout <> 4 Then
                'if we didn't time out we copy the the bytes into the array
                Array.Copy(inRepoBytes, 4, skylanderBytes, readBlock * 16, 16)
                readBlock = readBlock + 1
            End If

        Loop While readBlock <= &H3F
        Array.Clear(outRepoBytes, 0, 33)
        If checkBlankSkylander(skylanderBytes) Then
            readskylandData(skylanderBytes, True)
            ToolStripStatusLabel1.Text = "This Skylander is blank/new"
        Else
            skylanderBytes = decryptSkylander(skylanderBytes)
            readskylandData(skylanderBytes, False)
        End If
    End Sub

    Private Sub ReadSwapperOtherHalfToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

        'same as read from portal, but reads the second position skylander (usually the Top half of a swapper)
        Dim timeout As Integer
        Dim readBlock As Integer


        outRepoBytes(1) = &H52
        outputReport(portalHandle, outRepoBytes)
        System.Threading.Thread.Sleep(50)

        outRepoBytes(1) = &H41
        outRepoBytes(2) = 1
        outputReport(portalHandle, outRepoBytes)
        System.Threading.Thread.Sleep(500)

        outRepoBytes(1) = &H51
        outRepoBytes(2) = &H21
        readBlock = 0
        Do
            outRepoBytes(3) = readBlock
            outputReport(portalHandle, outRepoBytes)
            flushHid(portalHandle)
            timeout = 0
            Do
                inputReport(portalHandle, inRepoBytes)
                timeout = timeout + 1
            Loop Until inRepoBytes(1) <> &H53 Or timeout = 4

            If timeout <> 4 Then
                Array.Copy(inRepoBytes, 4, skylanderBytes, readBlock * 16, 16)
                readBlock = readBlock + 1
            End If

        Loop While readBlock <= &H3F
        Array.Clear(outRepoBytes, 0, 33)
        If checkBlankSkylander(skylanderBytes) Then
            readskylandData(skylanderBytes, True)
            ToolStripStatusLabel1.Text = "This Skylander is blank/new"
        Else
            skylanderBytes = decryptSkylander(skylanderBytes)
            readskylandData(skylanderBytes, False)
        End If
    End Sub

    Sub readskylandData(ByRef skylanderData As Byte(), ByVal blank As Boolean)

        'get skylander ID
        skylanderID = SwapEndianness(skylanderData(16), skylanderData(17))

        skyInfo = searchXML(exepath, skylanderID)

        Label10.Text = "ID: " & skylanderID

        If skyInfo(0) = "" Then
            'if figure ID is not in database
            Label7.Text = "Unknown ID: " & skylanderID
            Label8.Text = "- -"
            PictureBox1.Image = Nothing
            lockControls()
            ToolStripStatusLabel1.Text = "Unknown Skylander read"
            Exit Sub

        Else
            'get figure name and element
            Label7.Text = skyInfo(0)
            Label8.Text = skyInfo(1)
            'load skylander picture if available
            PictureBox1.ImageLocation = exepath & "images\" & Trim(Str(skylanderID)) & ".png"
            'lock controls if a swapper bottom half, item or stage is opened
            If Label8.Text = "Item" Or Label8.Text = "Stage" Or blank Then
                lockControls()
                ToolStripStatusLabel1.Text = "Finished Reading"
                Exit Sub
            End If
        End If
        'get high index and fill the variable sections
        If skylanderData(137) > skylanderData(585) Then
            moneyA = 131
            expA = 128
            hatA = 148
            skill = 144
            indexVal = 1
        ElseIf skylanderData(137) < skylanderData(585) Then
            moneyA = 579
            expA = 576
            hatA = 596
            skill = 592
            indexVal = 2
        End If

        If skylanderData(274) > skylanderData(722) Then
            expB = 275
            expC = 280
            hatB = 277
            hatC = 284
            hatD = 286
            indexValB = 1
        ElseIf skylanderData(274) < skylanderData(722) Then
            expB = 723
            expC = 728
            hatB = 725
            hatC = 732
            hatD = 734
            indexValB = 2
        End If

        If InStr(skyInfo(1), "Vehicle") Then
            skylanderGold = SwapEndianness(skylanderData(expC), skylanderData(expC + 1))
            TextBox2.Text = skylanderGold
            TextBox1.Text = 0
            ComboBox1.SelectedIndex = 0
            TextBox3.Text = ComboBox1.Text
            Exit Sub
        End If

        'get skylander hat
        skylanderHat = HatControl.gethat(skylanderData(hatA), skylanderData(hatB), skylanderData(hatC), skylanderData(hatD))
        ComboBox1.SelectedIndex = skylanderHat
        TextBox3.Text = ComboBox1.Text
        'get skylander gold
        skylanderGold = SwapEndianness(skylanderData(moneyA), skylanderData(moneyA + 1))
        TextBox2.Text = skylanderGold
        'get skylander xp
        skylanderEXP = SwapEndianness(skylanderData(expA), skylanderData(expA + 1)) + SwapEndianness(skylanderData(expB), skylanderData(expB + 1)) + SwapEndianness(skylanderData(expC), skylanderData(expC + 1))
        If skylanderData(expC + 2) = 1 Then
            skylanderEXP = skylanderEXP + 65536
        End If
        TextBox1.Text = skylanderEXP
        'get skylander skills
        skylanderSkill = SwapEndianness(skylanderData(skill), skylanderData(skill + 1))
        autom = True
        If skylanderSkill And 2 Then
            RadioButton1.Enabled = True
            RadioButton1.Checked = True
            RadioButton2.Enabled = False
        ElseIf (skylanderSkill And 1) And Not (skylanderSkill And 2) Then
            RadioButton2.Enabled = True
            RadioButton2.Checked = True
            RadioButton1.Enabled = False
        Else
            RadioButton1.Enabled = False
            RadioButton2.Enabled = False
        End If
        autom = False
        unlockControls()
        ToolStripStatusLabel1.Text = "Finished Reading"

        If skyInfo(0) IsNot "" Then
            Dim postReq As HttpWebRequest = DirectCast(WebRequest.Create("http://localhost:5000/SendSkylander"), HttpWebRequest)
            postReq.Method = "GET"
            postReq.Headers("name") = Label7.Text
            postReq.Headers("element") = Label8.Text
            postReq.Headers("hat") = TextBox3.Text
            If RadioButton2.Checked Then
                postReq.Headers("path") = "A"
            ElseIf RadioButton1.Checked Then
                postReq.Headers("path") = "B"
            End If
            postReq.Headers("coins") = TextBox2.Text
            postReq.Headers("level") = Label9.Text.Replace("Level: ", "")
            postReq.GetResponse()
        End If
    End Sub

    'we make sure only numeric values are allowed in our textboxes
    Private Sub TextBox1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox1.KeyPress
        If Not Char.IsNumber(e.KeyChar) AndAlso Not Char.IsControl(e.KeyChar) Then
            e.Handled = True
        End If
    End Sub
    Private Sub TextBox2_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox2.KeyPress
        If Not Char.IsNumber(e.KeyChar) AndAlso Not Char.IsControl(e.KeyChar) Then
            e.Handled = True
        End If
    End Sub

    'level 20 is the max and 197500 points is needed for it, we change the level label to match each level threshold
    Private Sub TextBox1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox1.TextChanged
        If TextBox1.Text <> "" Then
            If Int(TextBox1.Text) > 197500 Then
                TextBox1.Text = 197500
            ElseIf Int(TextBox1.Text) = 197500 Then
                Label9.Text = "Level: 20"
            ElseIf Int(TextBox1.Text) >= 174300 And Int(TextBox1.Text) < 197500 Then
                Label9.Text = "Level: 19"
            ElseIf Int(TextBox1.Text) >= 152600 And Int(TextBox1.Text) < 174300 Then
                Label9.Text = "Level: 18"
            ElseIf Int(TextBox1.Text) >= 132400 And Int(TextBox1.Text) < 152600 Then
                Label9.Text = "Level: 17"
            ElseIf Int(TextBox1.Text) >= 113700 And Int(TextBox1.Text) < 132400 Then
                Label9.Text = "Level: 16"
            ElseIf Int(TextBox1.Text) >= 96500 And Int(TextBox1.Text) < 113700 Then
                Label9.Text = "Level: 15"
            ElseIf Int(TextBox1.Text) >= 85800 And Int(TextBox1.Text) < 96500 Then
                Label9.Text = "Level: 14"
            ElseIf Int(TextBox1.Text) >= 69200 And Int(TextBox1.Text) < 85800 Then
                Label9.Text = "Level: 13"
            ElseIf Int(TextBox1.Text) >= 55000 And Int(TextBox1.Text) < 69200 Then
                Label9.Text = "Level: 12"
            ElseIf Int(TextBox1.Text) >= 43000 And Int(TextBox1.Text) < 55000 Then
                Label9.Text = "Level: 11"
            ElseIf Int(TextBox1.Text) >= 33000 And Int(TextBox1.Text) < 43000 Then
                Label9.Text = "Level: 10"
            ElseIf Int(TextBox1.Text) >= 24800 And Int(TextBox1.Text) < 33000 Then
                Label9.Text = "Level: 9"
            ElseIf Int(TextBox1.Text) >= 18200 And Int(TextBox1.Text) < 24800 Then
                Label9.Text = "Level: 8"
            ElseIf Int(TextBox1.Text) >= 13000 And Int(TextBox1.Text) < 18200 Then
                Label9.Text = "Level: 7"
            ElseIf Int(TextBox1.Text) >= 9000 And Int(TextBox1.Text) < 13000 Then
                Label9.Text = "Level: 6"
            ElseIf Int(TextBox1.Text) >= 6000 And Int(TextBox1.Text) < 9000 Then
                Label9.Text = "Level: 5"
            ElseIf Int(TextBox1.Text) >= 3800 And Int(TextBox1.Text) < 6000 Then
                Label9.Text = "Level: 4"
            ElseIf Int(TextBox1.Text) >= 2200 And Int(TextBox1.Text) < 3800 Then
                Label9.Text = "Level: 3"
            ElseIf Int(TextBox1.Text) >= 1000 And Int(TextBox1.Text) < 2200 Then
                Label9.Text = "Level: 2"
            ElseIf Int(TextBox1.Text) >= 0 And Int(TextBox1.Text) < 1000 Then
                Label9.Text = "Level: 1"
            End If
        End If
    End Sub

    Private Sub TextBox2_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox2.TextChanged
        'max money is 65000
        If TextBox2.Text <> "" Then
            If Int(TextBox2.Text) > 65000 Then
                TextBox2.Text = 65000
            End If
        End If
    End Sub


    'this is to catch if the portal is removed
    Protected Overrides Sub WndProc(ByRef m As Message)

        If m.Msg = DeviceManagement.WM_DEVICECHANGE Then
            If (m.WParam.ToInt32 = DeviceManagement.DBT_DEVICEREMOVECOMPLETE) Then

                ' If WParam contains DBT_DEVICEREMOVAL, a device has been removed.
                ' Find out if it's the device we're communicating with.

                If checkDevice(m) Then
                    lockPortalControls()
                    ToolStripStatusLabel1.Text = "Portal Removed!"
                End If

            End If
        End If
        MyBase.WndProc(m)

    End Sub

    Private Sub ConnectToolStripMenuItem_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub lockControls()
        ComboBox1.SelectedIndex = 0
        TextBox3.Text = ComboBox1.Text
        TextBox1.Text = 0
        TextBox2.Text = 0
        ComboBox1.Enabled = False
        TextBox1.Enabled = False
        TextBox2.Enabled = False
    End Sub

    Private Sub unlockControls()
        ComboBox1.Enabled = True
        TextBox1.Enabled = True
        TextBox2.Enabled = True
    End Sub

    Public Sub lockPortalControls()
        ReadSkylanderToolStripMenuItem.Enabled = False
    End Sub

    Public Sub unlockPortalControls()
        ReadSkylanderToolStripMenuItem.Enabled = True
    End Sub


    Private Sub AboutToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        AboutBox1.ShowDialog()
    End Sub

    'fill the metadata window and show it
    Private Sub ShowMetadataToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dialog1.mapData(skylanderBytes)
        Dialog1.ShowDialog()
    End Sub

    'cleanup by closing the database and closing the hid handles to the portal
    Private Sub Form1_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        CloseCommunications(portalHandle)
    End Sub

    'get the portal handle to work with
    Private Sub ConnectToPortalToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ConnectToPortalToolStripMenuItem.Click
        portalHandle = FindThePortal()
    End Sub

End Class
