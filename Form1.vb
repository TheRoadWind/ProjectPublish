
#Disable Warning
Public Class Form1
    ' 窗体控件声明
    Private WithEvents TxtProjectFile As TextBox
    Private WithEvents TxtPublishDir As TextBox
    Private WithEvents TxtLog As TextBox
    Private WithEvents TxtMsbuild As TextBox
    Private WithEvents BtnBrowseProject As Button
    Private WithEvents BtnBrowsePublish As Button
    Private WithEvents BtnBrowseMsbuild As Button
    Private WithEvents BtnPublish As Button
    Private WithEvents BtnCancel As Button
    Private WithEvents CboRuntime As ComboBox
    Private WithEvents ChkSelfContained As CheckBox
    Private WithEvents ChkSingleFile As CheckBox
    Private WithEvents ChkCreatExe As CheckBox
    Private lblProject As Label
    Private lblPublish As Label
    Private lblRuntime As Label
    Private LbAd As Label

    Private WithEvents TxtCustomExeName As TextBox
    Private WithEvents lblCustomExeName As Label

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        InitializeForm()

        '加载上次路径
        If Not String.IsNullOrEmpty(My.Settings.LastProjectFile) Then
            Dim lp As String = My.Settings.LastProjectFile
            If System.IO.File.Exists(lp) Then
                TxtProjectFile.Text = lp
            End If
        End If

        If Not String.IsNullOrEmpty(My.Settings.LastPublishDir) Then
            Dim lp As String = My.Settings.LastPublishDir
            If System.IO.Directory.Exists(lp) Then
                TxtPublishDir.Text = lp
            End If
        End If

        If Not String.IsNullOrEmpty(My.Settings.LastMsBuildFile) Then
            Dim lp As String = My.Settings.LastMsBuildFile
            If System.IO.File.Exists(lp) Then
                TxtMsbuild.Text = lp
            End If
        End If
    End Sub
    Private Sub SaveLastParts()
        Try
            If Not String.IsNullOrEmpty(TxtProjectFile.Text) Then
                My.Settings.LastProjectFile = TxtProjectFile.Text
            End If

            If Not String.IsNullOrEmpty(TxtPublishDir.Text) Then
                My.Settings.LastPublishDir = TxtPublishDir.Text
            End If

            If Not String.IsNullOrEmpty(TxtMsbuild.Text) Then
                My.Settings.LastMsBuildFile = TxtMsbuild.Text
            End If
            My.Settings.Save()
        Catch ex As Exception
        End Try
    End Sub
    Private Sub InitializeForm()
        ' 设置窗体属性
        Me.Text = "项目发布工具"
        Me.StartPosition = FormStartPosition.CenterScreen
        Me.MinimumSize = New Size(600, 600)
        Me.AllowDrop = True

        ' 创建控件
        CreateControls()

        ' 设置默认值
        CboRuntime.SelectedIndex = 0 ' win-x86
        ChkSelfContained.Checked = True
        ChkSingleFile.Checked = True

        ' 允许通过命令行参数传入项目路径
        Dim args As String() = Environment.GetCommandLineArgs()
        If args.Length > 1 AndAlso System.IO.File.Exists(args(1)) Then
            TxtProjectFile.Text = args(1)
        End If
    End Sub

    Private Sub CreateControls()
        Dim yPos As Integer = 20
        ' 项目文件选择
        lblProject = New Label With {
            .Text = "选择项目文件：",
            .Location = New Point(20, yPos),
            .Size = New Size(120, 25)
        }

        TxtProjectFile = New TextBox With {
            .Location = New Point(140, yPos),
            .Size = New Size(350, 25),
            .ReadOnly = True
        }

        BtnBrowseProject = New Button With {
            .Text = "浏览...",
            .Location = New Point(500, yPos),
            .Size = New Size(80, 25)
        }
        yPos += 40
        ' 发布目录选择
        lblPublish = New Label With {
            .Text = "选择发布目录：",
            .Location = New Point(20, yPos),
            .Size = New Size(120, 25)
        }

        TxtPublishDir = New TextBox With {
            .Location = New Point(140, yPos),
            .Size = New Size(350, 25),
            .ReadOnly = True
        }

        BtnBrowsePublish = New Button With {
            .Text = "浏览...",
            .Location = New Point(500, yPos),
            .Size = New Size(80, 25)
        }

        yPos += 40
        'Msbuild目录
        LbAd = New Label With {
            .Text = "MSbuild目录：",
            .Location = New Point(20, yPos),
            .Size = New Size(120, 25)
        }
        TxtMsbuild = New TextBox With {
            .Location = New Point(140, yPos),
            .Size = New Size(350, 25),
            .ReadOnly = True
        }

        BtnBrowseMsbuild = New Button With {
            .Text = "浏览...",
            .Location = New Point(500, yPos),
            .Size = New Size(80, 25)
        }

        yPos += 40

        ' 运行时标识符
        lblRuntime = New Label With {
            .Text = "运行时标识符：",
            .Location = New Point(20, yPos),
            .Size = New Size(120, 25)
        }

        CboRuntime = New ComboBox With {
            .Location = New Point(140, yPos),
            .Size = New Size(150, 25),
            .DropDownStyle = ComboBoxStyle.DropDownList
        }
        CboRuntime.Items.AddRange({"win-x86", "win-x64", "linux-x64", "osx-x64"})

        ' 复选框选项
        ChkSelfContained = New CheckBox With {
            .Text = "自包含",
            .Location = New Point(300, yPos),
            .Size = New Size(80, 25),
            .Checked = True
        }

        ChkSingleFile = New CheckBox With {
            .Text = "单文件",
            .Location = New Point(380, yPos),
            .Size = New Size(80, 25),
            .Checked = True
        }
        ChkCreatExe = New CheckBox With {
            .Text = "生成压缩包",
            .Location = New Point(460, yPos),
            .Size = New Size(100, 25),
            .Checked = False}
        ' 添加最简单的安装包选项
        yPos += 40

        ' --- 自定义程序名称输入框 ---
        lblCustomExeName = New Label With {
            .Text = "设置版本号：",
            .Location = New Point(20, yPos),
            .Size = New Size(120, 25),
            .Enabled = ChkSingleFile.Checked ' 初始状态禁用，跟随“单文件”复选框
        }

        TxtCustomExeName = New TextBox With {
            .Location = New Point(140, yPos),
            .Size = New Size(200, 25),
            .Enabled = ChkSingleFile.Checked ' 初始状态禁用
        }
        ' 提示文本
        Dim lblTip As New Label With {
            .Text = "(仅当勾选‘单文件’时生效，无需后缀,格式例如:1.0.2)",
            .Location = New Point(350, yPos),
            .AutoSize = True,
            .ForeColor = Color.Gray,
            .Font = New Font("微软雅黑", 8)
        }
        yPos += 40


        ' 发布按钮
        BtnPublish = New Button With {
            .Text = "开始发布",
            .Location = New Point(140, yPos),
            .Size = New Size(120, 35),
            .BackColor = Color.FromArgb(0, 123, 255),
            .ForeColor = Color.White,
            .Font = New Font("微软雅黑", 10, FontStyle.Bold)
        }

        BtnCancel = New Button With {
            .Text = "取消",
            .Location = New Point(270, yPos),
            .Size = New Size(120, 35)
        }
        Dim CreateZip As New Button With {
            .Text = "制作压缩包",
            .Location = New Point(400, yPos),
            .Size = New Size(150, 35),
            .Enabled = False,
            .BackColor = Color.LightGreen
        }
        yPos += 40
        ' 日志文本框
        TxtLog = New TextBox With {
            .Multiline = True,
            .ScrollBars = ScrollBars.Vertical,
            .Location = New Point(0, yPos),
            .Size = New Size(Me.Width, 260),
            .ReadOnly = True,
            .Font = New Font("楷体", 10),
            .BackColor = Color.Black,
            .ForeColor = Color.White,
            .Visible = False
        }


        ' 添加到窗体
        Me.Controls.AddRange({
            lblProject, TxtProjectFile, BtnBrowseProject,
            lblPublish, TxtPublishDir, BtnBrowsePublish,
            LbAd, TxtMsbuild, BtnBrowseMsbuild,
            lblRuntime, CboRuntime, ChkSelfContained, ChkSingleFile,
            BtnPublish, BtnCancel, TxtLog, ChkCreatExe, CreateZip, lblCustomExeName, TxtCustomExeName, lblTip
        })
        AddHandler ChkCreatExe.CheckedChanged, Sub()
                                                   CreateZip.Enabled = ChkCreatExe.Checked
                                               End Sub
        AddHandler CreateZip.Click, Sub()
                                        If ChkCreatExe.Checked Then
                                            TxtLog.Visible = True
                                            Dim cur = DateTime.Now
                                            Dim publishDir As String = TxtPublishDir.Text
                                            Dim projectName As String = InputBox("请输入制作成压缩包(.ZIP)的名称", "压缩包命名提示")
                                            CreateDebugOutputPackage(projectName, publishDir)
                                            LogMessage($"打包状态: 成功 ✓")
                                            LogMessage($"总耗时:{DateDiff("s", cur, Now)}s")
                                        End If
                                    End Sub
        ' --- 绑定“单文件”复选框与自定义名称输入框的启用状态 ---
        AddHandler ChkSingleFile.CheckedChanged, Sub()
                                                     Dim isEnabled As Boolean = ChkSingleFile.Checked
                                                     lblCustomExeName.Enabled = isEnabled
                                                     TxtCustomExeName.Enabled = isEnabled
                                                     ' 如果禁用，清空输入框内容
                                                     If Not isEnabled Then
                                                         TxtCustomExeName.Text = ""
                                                     End If
                                                 End Sub
    End Sub

    Private Sub BtnBrowseProject_Click(sender As Object, e As EventArgs) Handles BtnBrowseProject.Click
        Using dialog As New OpenFileDialog()
            dialog.Filter = "项目文件 (*.csproj;*.vbproj;*.fsproj)|*.csproj;*.vbproj;*.fsproj|所有文件 (*.*)|*.*"
            dialog.Title = "选择项目文件"

            If Not String.IsNullOrEmpty(My.Settings.LastProjectFile) Then
                Dim lp = System.IO.Path.GetDirectoryName(My.Settings.LastProjectFile)
                If System.IO.Directory.Exists(lp) Then
                    dialog.InitialDirectory = lp
                End If
            End If

            If dialog.ShowDialog() = DialogResult.OK Then
                TxtProjectFile.Text = dialog.FileName

                SaveLastParts()
            End If
        End Using
    End Sub
    Private Sub BtnBrowseMsbuild_Click(sender As Object, e As EventArgs) Handles BtnBrowseMsbuild.Click
        Using dialog As New OpenFileDialog()
            dialog.Filter = "MsBuild程序 (*.exe)|*.exe|所有文件 (*.*)|*.*"
            dialog.Title = "选择MSbuild程序文件"

            If Not String.IsNullOrEmpty(My.Settings.LastMsBuildFile) Then
                Dim lp = System.IO.Path.GetDirectoryName(My.Settings.LastMsBuildFile)
                If System.IO.Directory.Exists(lp) Then
                    dialog.InitialDirectory = lp
                End If
            End If

            If dialog.ShowDialog() = DialogResult.OK Then
                TxtMsbuild.Text = dialog.FileName

                SaveLastParts()
            End If
        End Using
    End Sub
    Private Sub BtnBrowsePublish_Click(sender As Object, e As EventArgs) Handles BtnBrowsePublish.Click
        Using dialog As New FolderBrowserDialog()
            dialog.Description = "选择发布目录"
            dialog.ShowNewFolderButton = True

            If Not String.IsNullOrEmpty(My.Settings.LastPublishDir) Then
                If System.IO.Directory.Exists(My.Settings.LastPublishDir) Then
                    dialog.SelectedPath = My.Settings.LastPublishDir
                End If
            End If

            If dialog.ShowDialog() = DialogResult.OK Then
                TxtPublishDir.Text = dialog.SelectedPath

                SaveLastParts()
            End If
        End Using
    End Sub
    Private Sub WriteFileData(iPath As String, fileContent As Object, Optional Encod As System.Text.Encoding = Nothing)
        Try
            If Encod Is Nothing Then
                Encod = System.Text.Encoding.UTF8
            End If
            If TypeOf (fileContent) Is String Then
                System.IO.File.WriteAllText(iPath, fileContent, Encod)
            ElseIf TypeOf (fileContent) Is List(Of String) Then
                System.IO.File.WriteAllLines(iPath, fileContent, Encod)
            End If

            Return
        Catch ex As Exception
            MessageBox.Show("保存文件时出错: " & ex.Message & vbCrLf & "文件路径: " & iPath)
            Return
        End Try
    End Sub
    Private Sub UpdataForms(iPath As String)
        ' 创建一个小型通知窗口
        Dim updateForm As New Form()
        With updateForm
            .Text = "新版本内容填写界面"
            .Width = 400
            .Height = 400
            .StartPosition = FormStartPosition.CenterScreen
            .FormBorderStyle = FormBorderStyle.FixedDialog
            .MaximizeBox = False
            .MinimizeBox = False
        End With

        ' 创建标签
        Dim lblMessage As New TextBox With {
        .Location = New Point(20, 20),
        .Size = New Size(350, 200),
        .Font = New Font("微软雅黑", 12), .Multiline = True
    }

        ' 创建按钮
        Dim btnUpdateNow As New Button With {
        .Text = "确定",
        .Location = New Point(150, 240),
        .Size = New Size(100, 30),
        .BackColor = Color.LightGreen
    }

        updateForm.Controls.AddRange(btnUpdateNow, lblMessage)


        AddHandler btnUpdateNow.Click, Sub()
                                           Dim txt = lblMessage.Text
                                           If Not String.IsNullOrEmpty(txt) Then
                                               WriteFileData(iPath, txt)
                                           End If

                                           updateForm.DialogResult = DialogResult.Yes
                                       End Sub
        updateForm.ShowDialog()

        If updateForm.DialogResult = DialogResult.Yes Then
            ProjectPulish()
        End If

    End Sub
    Private Sub BtnPublish_Click(sender As Object, e As EventArgs) Handles BtnPublish.Click

        SaveLastParts()

        ' 验证输入
        If String.IsNullOrWhiteSpace(TxtProjectFile.Text) Then
            MessageBox.Show("请选择项目文件", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End If

        If String.IsNullOrWhiteSpace(TxtPublishDir.Text) Then
            MessageBox.Show("请选择发布目录", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End If

        If Not System.IO.File.Exists(TxtProjectFile.Text) Then
            MessageBox.Show("项目文件不存在", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End If
        Dim IsCon As Boolean = TxtProjectFile.Text.Contains("Diff.vbproj")
        If IsCon Then
            Dim re As DialogResult = MessageBox.Show("是否添加新版本新增内容？", "新版本新增内容提示", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
            If re = DialogResult.Yes Then
                '更新文件路径
                Dim Upath As String = "\\10.17.28.28\文件中转站\Program\updata.txt"

                Try
                    If System.IO.File.Exists(Upath) Then

                        '写入更新内容
                        UpdataForms(Upath)
                        Return
                    Else
                        GoTo NextStep
                    End If
                Catch ex As Exception
                    GoTo NextStep
                End Try
            Else
                GoTo NextStep
            End If
        End If

NextStep:
        ProjectPulish()
    End Sub
    Private Sub ProjectPulish()
        ' 准备发布参数
        Dim projectFile As String = TxtProjectFile.Text
        Dim publishDir As String = TxtPublishDir.Text
        Dim runtimeId As String = If(CboRuntime.SelectedItem?.ToString(), "win-x86")
        Dim projectName As String = System.IO.Path.GetFileName(projectFile)



        ' 构建 MSBuild 参数
        Dim arguments As String = $"""{projectFile}"" " &
                          "/t:Restore;Publish " &
                          "/p:Configuration=Release " &
                          $"/p:SelfContained={If(ChkSelfContained.Checked, "true", "false")} " &
                          $"/p:PublishSingleFile={If(ChkSingleFile.Checked, "true", "false")} " &
                          $"/p:RuntimeIdentifier={runtimeId} " &
                          "/p:IncludeNativeLibrariesForSelfExtract=true " &
                          "/p:SignManifests=false " &          ' 新增：禁用清单签名
                          "/p:SignAssembly=false " &           ' 新增：禁用程序集签名
                          "/p:GenerateClickOnceManifests=false " & ' 新增：完全禁用 ClickOnce
                          "/p:BootstrapperEnabled=false " &     ' ★ 新增：禁用引导程序生成
                          $"/p:PublishDir=""{publishDir}"" " &
                          "/verbosity:minimal " &
                          "/nologo"

        ' --- 新增：如果勾选了“单文件”并填写了自定义名称，则添加 AssemblyName 参数 ---
        If ChkSingleFile.Checked Then
            ' 移除用户可能输入的后缀，并确保名称合法（这里做简单处理）
            Dim customName As String = TxtCustomExeName.Text.Trim()
            Dim Vnum As String = ""
            If customName <> "" Then
                Vnum = $"_v{customName}"
            End If
            arguments = $"/p:AssemblyName={projectName.Substring(0, projectName.Length - 7)}{Vnum} " & arguments
        End If

        ' 查找 MSBuild
        Dim msbuildPath As String = TxtMsbuild.Text
        If String.IsNullOrEmpty(msbuildPath) Then
            MessageBox.Show("请重新选择Msbuild程序路径", "错误", MessageBoxButtons.OK)
            Return
        End If

        ' 显示确认对话框
        Dim confirm As DialogResult = MessageBox.Show(
            $"即将发布项目：{projectName}" & vbCrLf & vbCrLf &
            $"发布路径：{publishDir}" & vbCrLf &
            $"运行时：{runtimeId}" & vbCrLf & vbCrLf &
            $"确认开始发布吗？",
            "确认发布",
            MessageBoxButtons.YesNo,
            MessageBoxIcon.Question)

        If confirm <> DialogResult.Yes Then
            Return
        End If
        ' ========== 清理旧版本 ==========
        Try
            Dim baseExeName As String = System.IO.Path.GetFileNameWithoutExtension(projectFile)   ' "Diff"
            Dim newVersionSuffix As String = ""
            If ChkSingleFile.Checked AndAlso Not String.IsNullOrWhiteSpace(TxtCustomExeName.Text) Then
                newVersionSuffix = $"_v{TxtCustomExeName.Text.Trim()}"
            End If
            Dim newExeFullName As String = $"{baseExeName}{newVersionSuffix}.exe"

            Dim oldExes As String() = System.IO.Directory.GetFiles(publishDir, $"{baseExeName}*.exe", System.IO.SearchOption.TopDirectoryOnly)
            For Each oldExe In oldExes
                Dim oldFileName As String = System.IO.Path.GetFileName(oldExe)
                If String.Equals(oldFileName, newExeFullName, StringComparison.OrdinalIgnoreCase) Then
                    Continue For
                End If
                System.IO.File.Delete(oldExe)
                LogMessage($"已删除旧版本: {oldFileName}")
            Next
        Catch ex As Exception
            LogMessage($"清理旧版本时出错: {ex.Message}")
        End Try

        TxtLog.Visible = True

        Dim cur As Date = DateTime.Now
        ' 开始发布
        BtnPublish.Enabled = False
        BtnCancel.Enabled = False
        TxtLog.Clear()

        Try
            LogMessage("=== 开始发布 ===")
            LogMessage($"项目文件: {projectFile}")
            LogMessage($"发布目录: {publishDir}")
            LogMessage($"运行时: {runtimeId}")
            LogMessage($"MSBuild路径: {msbuildPath}")
            LogMessage("")

            Dim startInfo As New ProcessStartInfo With {
                .FileName = msbuildPath,
                .Arguments = arguments,
                .WorkingDirectory = System.IO.Path.GetDirectoryName(projectFile),
                .UseShellExecute = False,
                .RedirectStandardOutput = True,
                .RedirectStandardError = True,
                .CreateNoWindow = True,
                .StandardErrorEncoding = System.Text.Encoding.UTF8,
                .StandardOutputEncoding = System.Text.Encoding.UTF8
            }

            Using process As New Process()
                process.StartInfo = startInfo

                ' 实时输出处理
                AddHandler process.OutputDataReceived,
                    Sub(s, evt)
                        If Not String.IsNullOrEmpty(evt.Data) Then
                            Me.Invoke(Sub() LogMessage(evt.Data))
                        End If
                    End Sub

                AddHandler process.ErrorDataReceived,
                    Sub(s, evt)
                        If Not String.IsNullOrEmpty(evt.Data) Then
                            Me.Invoke(Sub() LogMessage($"[错误] {evt.Data}"))
                        End If
                    End Sub

                process.Start()
                process.BeginOutputReadLine()
                process.BeginErrorReadLine()

                ' 等待完成
                While Not process.HasExited
                    Application.DoEvents()
                    Threading.Thread.Sleep(100)
                End While

                LogMessage("")
                LogMessage($"=== 发布完成 ===")
                LogMessage($"退出代码: {process.ExitCode}")

                ' 检查目标 exe 是否生成
                Dim targetExePattern As String = System.IO.Path.GetFileNameWithoutExtension(projectFile) & "_v*.exe"
                If ChkSingleFile.Checked AndAlso Not String.IsNullOrWhiteSpace(TxtCustomExeName.Text) Then
                    targetExePattern = System.IO.Path.GetFileNameWithoutExtension(projectFile) & $"_v{TxtCustomExeName.Text.Trim()}.exe"
                Else
                    targetExePattern = System.IO.Path.GetFileNameWithoutExtension(projectFile) & ".exe"
                End If

                Dim generatedExe As String = System.IO.Directory.GetFiles(publishDir, targetExePattern).FirstOrDefault()

                If Not String.IsNullOrEmpty(generatedExe) Then
                    LogMessage($"状态: 成功 ✓ (主程序已生成: {System.IO.Path.GetFileName(generatedExe)})")
                    LogMessage($"总耗时:{DateDiff("s", cur, Now)}s")

                    ' 询问是否创建桌面快捷方式
                    Dim createShortcutResult As DialogResult = MessageBox.Show(
                        "发布成功！是否在桌面创建快捷方式？",
                        "创建快捷方式",
                        MessageBoxButtons.YesNo,
                        MessageBoxIcon.Question)

                    If createShortcutResult = DialogResult.Yes Then
                        Try
                            ' 1. 获取桌面路径
                            Dim desktopPath As String = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)

                            ' 2. 确定目标程序路径
                            Dim targetPath As String = ""
                            Dim projectBaseName As String = System.IO.Path.GetFileNameWithoutExtension(projectFile)
                            Dim expectedExePath As String = System.IO.Path.Combine(publishDir, projectBaseName & ".exe")

                            If System.IO.File.Exists(expectedExePath) Then
                                targetPath = expectedExePath
                            Else
                                Dim exeFiles As String() = System.IO.Directory.GetFiles(publishDir, "*.exe", System.IO.SearchOption.TopDirectoryOnly)
                                If exeFiles.Length > 0 Then
                                    targetPath = exeFiles(0)
                                Else
                                    MessageBox.Show($"在发布目录中未找到可执行文件 (.exe)。", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                                    MessageBox.Show("发布成功！", "完成", MessageBoxButtons.OK, MessageBoxIcon.Information)
                                    Return
                                End If
                            End If

                            ' 3. 获取快捷方式名称（包含后缀）
                            Dim defaultName As String = projectName.Substring(0, projectName.Length - 7）
                            Dim shortcutName As String = InputBox("请输入桌面快捷方式的名称（包含.lnk后缀）：", "快捷方式命名", defaultName & ".lnk")

                            ' 如果用户取消或输入为空，使用默认名称
                            If String.IsNullOrWhiteSpace(shortcutName) Then
                                shortcutName = defaultName & ".lnk"
                            End If

                            ' 确保名称包含.lnk后缀
                            If Not shortcutName.EndsWith(".lnk", StringComparison.OrdinalIgnoreCase) Then
                                shortcutName &= ".lnk"
                            End If

                            ' 4. 构建完整的快捷方式路径
                            Dim shortcutPath As String = System.IO.Path.Combine(desktopPath, shortcutName)

                            ' 5. 使用后期绑定创建快捷方式
                            Dim shellType As Type = Type.GetTypeFromProgID("WScript.Shell")
                            Dim shell As Object = Activator.CreateInstance(shellType)
                            Dim shortcut As Object = shellType.InvokeMember("CreateShortcut",
                                                                          System.Reflection.BindingFlags.InvokeMethod,
                                                                          Nothing,
                                                                          shell,
                                                                          New Object() {shortcutPath})

                            ' 设置快捷方式属性
                            shortcut.GetType().InvokeMember("TargetPath",
                                                           System.Reflection.BindingFlags.SetProperty,
                                                           Nothing,
                                                           shortcut,
                                                           New Object() {targetPath})

                            shortcut.GetType().InvokeMember("WorkingDirectory",
                                                           System.Reflection.BindingFlags.SetProperty,
                                                           Nothing,
                                                           shortcut,
                                                           New Object() {System.IO.Path.GetDirectoryName(targetPath)})

                            shortcut.GetType().InvokeMember("Description",
                                                           System.Reflection.BindingFlags.SetProperty,
                                                           Nothing,
                                                           shortcut,
                                                           New Object() {$"快捷方式到 {System.IO.Path.GetFileName(targetPath)}"})

                            shortcut.GetType().InvokeMember("IconLocation",
                                                           System.Reflection.BindingFlags.SetProperty,
                                                           Nothing,
                                                           shortcut,
                                                           New Object() {targetPath & ",0"})

                            ' 保存快捷方式
                            shortcut.GetType().InvokeMember("Save",
                                                           System.Reflection.BindingFlags.InvokeMethod,
                                                           Nothing,
                                                           shortcut,
                                                           Nothing)

                            LogMessage($"桌面快捷方式创建成功: {shortcutName}")
                            MessageBox.Show($"已在桌面创建快捷方式：{shortcutName}", "成功", MessageBoxButtons.OK, MessageBoxIcon.Information)

                        Catch ex As Exception
                            LogMessage($"[快捷方式创建失败] {ex.Message}")
                            MessageBox.Show($"创建桌面快捷方式时出错：{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        End Try
                    End If
                    MessageBox.Show("发布成功！", "完成", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Else
                    LogMessage($"状态: 失败 ✗ (未找到生成的主程序)")
                    MessageBox.Show($"发布失败，未找到主程序文件。", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End If

            End Using

        Catch ex As Exception
            LogMessage($"[异常] {ex.Message}")
            MessageBox.Show($"发布过程中出现异常：{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            ' 在发布完成后（无论退出代码是否为 0），清理不需要的文件
            Try
                ' 删除 setup.exe、.application、.manifest 等 ClickOnce 文件
                Dim extraFiles As String() = System.IO.Directory.GetFiles(publishDir, "setup.exe") _
                        .Concat(System.IO.Directory.GetFiles(publishDir, "*.application")) _
                        .Concat(System.IO.Directory.GetFiles(publishDir, "*.manifest")) _
                        .Concat(System.IO.Directory.GetFiles(publishDir, "*.pdb")) _
                        .Concat(System.IO.Directory.GetFiles(publishDir, "*.tmp")) _
                        .ToArray()
                For Each f In extraFiles
                    If System.IO.File.Exists(f) Then
                        System.IO.File.Delete(f)
                        LogMessage($"已删除多余文件: {System.IO.Path.GetFileName(f)}")
                    End If
                Next

                ' 删除 Application Files 文件夹
                Dim appFilesDir As String = System.IO.Path.Combine(publishDir, "Application Files")
                If System.IO.Directory.Exists(appFilesDir) Then
                    System.IO.Directory.Delete(appFilesDir, True)
                    LogMessage("已删除文件夹: Application Files")
                End If

            Catch ex As Exception
                LogMessage($"清理多余文件时出错: {ex.Message}")
            End Try

            BtnPublish.Enabled = True
            BtnCancel.Enabled = True
        End Try
    End Sub



    Private Sub LogMessage(message As String)
        TxtLog.AppendText($"[{DateTime.Now:HH:mm:ss}] {message}" & vbCrLf)
        TxtLog.ScrollToCaret()
    End Sub

    Private Sub BtnCancel_Click(sender As Object, e As EventArgs) Handles BtnCancel.Click
        Me.Close()
    End Sub

    ' 支持拖放项目文件到窗体
    Private Sub Form1_DragEnter(sender As Object, e As DragEventArgs) Handles Me.DragEnter
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            e.Effect = DragDropEffects.Copy
        End If
    End Sub

    Private Sub Form1_DragDrop(sender As Object, e As DragEventArgs) Handles Me.DragDrop
        Dim files As String() = CType(e.Data.GetData(DataFormats.FileDrop), String())
        If files.Length > 0 Then
            Dim file As String = files(0)
            If file.EndsWith(".csproj") OrElse file.EndsWith(".vbproj") OrElse file.EndsWith(".fsproj") Then
                TxtProjectFile.Text = file
            End If
        End If
    End Sub

    Private Sub TxtProjectFile_DragEnter(sender As Object, e As DragEventArgs) Handles TxtProjectFile.DragEnter
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            e.Effect = DragDropEffects.Copy
        End If
    End Sub

    Private Sub TxtProjectFile_DragDrop(sender As Object, e As DragEventArgs) Handles TxtProjectFile.DragDrop
        Dim files As String() = CType(e.Data.GetData(DataFormats.FileDrop), String())
        If files.Length > 0 Then
            Dim file As String = files(0)
            If file.EndsWith(".csproj") OrElse file.EndsWith(".vbproj") OrElse file.EndsWith(".fsproj") Then
                TxtProjectFile.Text = file
            End If
        End If
    End Sub

    Private Sub CreateDebugOutputPackage(projectName As String, publishDir As String)
        ' 检查 publishDir 是否存在
        If Not (System.IO.Directory.Exists(publishDir) OrElse System.IO.File.Exists(publishDir)) Then
            LogMessage($"[错误] 发布路径不存在: {publishDir}")
            Return
        End If

        ' 确定输出目录和 ZIP 文件路径
        Dim outputDir As String = System.IO.Path.GetDirectoryName(publishDir)
        Dim zipFile As String = System.IO.Path.Combine(outputDir, $"{projectName}.zip")

        ' 如果 ZIP 文件已存在，则删除旧文件
        If System.IO.File.Exists(zipFile) Then
            System.IO.File.Delete(zipFile)
            LogMessage($"删除旧文件: {zipFile}")
        End If

        ' 判断 publishDir 是文件夹还是文件
        If System.IO.Directory.Exists(publishDir) Then
            ' publishDir 是文件夹，打包整个文件夹
            LogMessage($"正在打包文件夹: {publishDir}")
            Try
                System.IO.Compression.ZipFile.CreateFromDirectory(publishDir, zipFile, System.IO.Compression.CompressionLevel.Fastest, False)
                Dim info As New System.IO.FileInfo(zipFile)
                LogMessage($"✓ 打包完成: {info.Length / 1024 / 1024:F1} MB")
            Catch ex As Exception
                LogMessage($"[错误] {ex.Message}")
                LogMessage($"[错误堆栈跟踪] {ex.StackTrace}")
            End Try
        Else
            ' publishDir 是文件，直接将文件打包成 ZIP
            If System.IO.File.Exists(publishDir) Then
                LogMessage($"正在打包文件: {publishDir}")

                ' 创建 ZIP 文件
                Try
                    Using zipStream As New System.IO.FileStream(zipFile, System.IO.FileMode.Create)
                        Using zipArchive As System.IO.Compression.ZipArchive = New System.IO.Compression.ZipArchive(zipStream, System.IO.Compression.ZipArchiveMode.Create)
                            ' 将文件添加到 ZIP 文件中
                            Dim entryName As String = System.IO.Path.GetFileName(publishDir)
                            System.IO.Compression.ZipFileExtensions.CreateEntryFromFile(zipArchive, publishDir, entryName, System.IO.Compression.CompressionLevel.Fastest)
                        End Using
                    End Using

                    Dim info As New System.IO.FileInfo(zipFile)
                    LogMessage($"✓ 打包完成: {info.Length / 1024 / 1024:F1} MB")
                Catch ex As Exception
                    LogMessage($"[错误] {ex.Message}")
                    LogMessage($"[错误堆栈跟踪] {ex.StackTrace}")
                End Try
            Else
                LogMessage($"[错误] publishDir 既不是文件也不是文件夹: {publishDir}")
                Return
            End If
        End If
    End Sub



End Class

