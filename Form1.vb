#Disable Warning
Public Class Form1
    ' ==================== 依赖扫描相关常量 ====================
    ''' <summary>
    ''' 扫描项目目录时需要排除的文件/文件夹名称（不区分大小写）。
    ''' </summary>
    Private Shared ReadOnly DependencyIgnoreNames As String() = {
        ".git", ".gitignore", ".gitattributes", ".gitmodules", ".github",
        ".vs", "My Project", "bin", "obj",
        "README.md", "readme.md", "ReadMe.md", "README.txt", "README"
    }

    ''' <summary>
    ''' 扫描项目目录时需要排除的文件扩展名（不区分大小写）。
    ''' 这些是源码/项目/IDE 相关文件，不属于运行时依赖。
    ''' </summary>
    Private Shared ReadOnly DependencyIgnoreExtensions As String() = {
        ".vb", ".cs", ".fs",
        ".resx",
        ".config",
        ".sln", ".slnx", ".suo", ".user",
        ".vbproj", ".csproj", ".fsproj",
        ".pdb", ".xml",
        ".tmp", ".cache"
    }

    ' ==================== 窗体控件声明 ====================
    ' 项目文件路径输入框
    Private WithEvents TxtProjectFile As TextBox
    ' 发布目录输入框
    Private WithEvents TxtPublishDir As TextBox
    ' 日志输出框
    Private WithEvents TxtLog As TextBox
    ' MSBuild 程序路径输入框
    Private WithEvents TxtMsbuild As TextBox
    ' 浏览项目文件按钮
    Private WithEvents BtnBrowseProject As Button
    ' 浏览发布目录按钮
    Private WithEvents BtnBrowsePublish As Button
    ' 浏览 MSBuild 按钮
    Private WithEvents BtnBrowseMsbuild As Button
    ' 开始发布按钮
    Private WithEvents BtnPublish As Button
    ' 取消按钮
    Private WithEvents BtnCancel As Button
    ' 运行时标识符下拉框
    Private WithEvents CboRuntime As ComboBox
    ' 自包含复选框
    Private WithEvents ChkSelfContained As CheckBox
    ' 单文件复选框
    Private WithEvents ChkSingleFile As CheckBox
    ' 生成压缩包复选框
    Private WithEvents ChkCreatExe As CheckBox
    ' 项目文件标签
    Private lblProject As Label
    ' 发布目录标签
    Private lblPublish As Label
    ' 运行时标识符标签
    Private lblRuntime As Label
    ' MSBuild 目录标签
    Private LbAd As Label

    ' 设置版本号输入框
    Private WithEvents TxtCustomExeName As TextBox
    ' 设置版本号标签
    Private WithEvents lblCustomExeName As Label
    ' 制作压缩包按钮
    Private WithEvents BtnCreateZip As Button

    ' 依赖勾选列表框（列出扫描到的依赖，默认全部勾选）
    Private WithEvents ClbDependencies As CheckedListBox
    ' 依赖列表框标签
    Private lblDependencies As Label

    ' ==================== 窗体加载 ====================
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        InitializeForm()

        ' 加载上次项目文件路径
        If Not String.IsNullOrEmpty(My.Settings.LastProjectFile) Then
            Dim lp As String = My.Settings.LastProjectFile
            If System.IO.File.Exists(lp) Then
                TxtProjectFile.Text = lp
                LoadVersionFromProject(lp)
                RefreshDependencyList(lp)
            End If
        End If

        ' 加载上次发布目录
        If Not String.IsNullOrEmpty(My.Settings.LastPublishDir) Then
            Dim lp As String = My.Settings.LastPublishDir
            If System.IO.Directory.Exists(lp) Then
                TxtPublishDir.Text = lp
            End If
        End If

        ' 加载上次 MSBuild 路径；若没有则自动查找
        If Not String.IsNullOrEmpty(My.Settings.LastMsBuildFile) Then
            Dim lp As String = My.Settings.LastMsBuildFile
            If System.IO.File.Exists(lp) Then
                TxtMsbuild.Text = lp
            End If
        End If

        ' 若 MSBuild 仍为空，则尝试自动获取
        If String.IsNullOrWhiteSpace(TxtMsbuild.Text) Then
            Dim autoMsbuild As String = FindMsBuildPath()
            If Not String.IsNullOrEmpty(autoMsbuild) Then
                TxtMsbuild.Text = autoMsbuild
                SaveLastParts()
            End If
        End If
    End Sub

    ' ==================== 保存上次使用路径 ====================
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
            ' 忽略保存异常
        End Try
    End Sub

    ' ==================== 初始化窗体 ====================
    Private Sub InitializeForm()
        ' 设置窗体属性
        Me.Text = "项目发布工具"
        Me.StartPosition = FormStartPosition.CenterScreen
        Me.MinimumSize = New Size(720, 700)
        Me.AllowDrop = True
        ' 让日志先 Dock，避免遮挡
        Me.Padding = New Padding(0)

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
            LoadVersionFromProject(args(1))
            RefreshDependencyList(args(1))
        End If
    End Sub

    ' ==================== 创建控件 ====================
    Private Sub CreateControls()
        ' 起始位置适当上移
        Dim yPos As Integer = 10
        Dim rowGap As Integer = 35

        ' 项目文件选择
        lblProject = New Label With {
            .Text = "选择项目文件：",
            .Location = New Point(20, yPos),
            .Size = New Size(120, 25)
        }

        TxtProjectFile = New TextBox With {
            .Location = New Point(140, yPos),
            .Size = New Size(430, 25),
            .ReadOnly = True
        }

        BtnBrowseProject = New Button With {
            .Text = "浏览...",
            .Location = New Point(580, yPos),
            .Size = New Size(80, 25)
        }
        yPos += rowGap

        ' 发布目录选择
        lblPublish = New Label With {
            .Text = "选择发布目录：",
            .Location = New Point(20, yPos),
            .Size = New Size(120, 25)
        }

        TxtPublishDir = New TextBox With {
            .Location = New Point(140, yPos),
            .Size = New Size(430, 25),
            .ReadOnly = True
        }

        BtnBrowsePublish = New Button With {
            .Text = "浏览...",
            .Location = New Point(580, yPos),
            .Size = New Size(80, 25)
        }

        yPos += rowGap

        ' MSBuild 目录
        LbAd = New Label With {
            .Text = "MSbuild目录：",
            .Location = New Point(20, yPos),
            .Size = New Size(120, 25)
        }
        TxtMsbuild = New TextBox With {
            .Location = New Point(140, yPos),
            .Size = New Size(430, 25),
            .ReadOnly = True
        }

        BtnBrowseMsbuild = New Button With {
            .Text = "浏览...",
            .Location = New Point(580, yPos),
            .Size = New Size(80, 25)
        }

        yPos += rowGap

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
            .Checked = False
        }

        yPos += rowGap

        ' --- 设置版本号输入框 ---
        lblCustomExeName = New Label With {
            .Text = "设置版本号：",
            .Location = New Point(20, yPos),
            .Size = New Size(120, 25),
            .Enabled = ChkSingleFile.Checked
        }

        TxtCustomExeName = New TextBox With {
            .Location = New Point(140, yPos),
            .Size = New Size(200, 25),
            .Enabled = ChkSingleFile.Checked
        }

        ' 提示文本
        Dim lblTip As New Label With {
            .Text = "(仅当勾选‘单文件’时生效，无需后缀,格式例如:1.0.2)",
            .Location = New Point(350, yPos),
            .AutoSize = True,
            .ForeColor = Color.Gray,
            .Font = New Font("微软雅黑", 8)
        }

        yPos += rowGap

        ' --- 依赖勾选列表框 ---
        lblDependencies = New Label With {
            .Text = "运行时依赖项：",
            .Location = New Point(20, yPos),
            .Size = New Size(120, 25)
        }

        ClbDependencies = New CheckedListBox With {
            .Location = New Point(140, yPos),
            .Size = New Size(430, 100),
            .CheckOnClick = True,
            .HorizontalScrollbar = True
        }

        yPos += 110

        ' 发布按钮
        BtnPublish = New Button With {
            .Text = "开始发布",
            .Location = New Point(140, yPos),
            .Size = New Size(120, 35),
            .BackColor = Color.FromArgb(0, 123, 255),
            .ForeColor = Color.White,
            .Font = New Font("微软雅黑", 10, FontStyle.Bold)
        }

        ' 取消按钮
        BtnCancel = New Button With {
            .Text = "取消",
            .Location = New Point(270, yPos),
            .Size = New Size(120, 35)
        }

        ' 制作压缩包按钮
        BtnCreateZip = New Button With {
            .Text = "制作压缩包",
            .Location = New Point(400, yPos),
            .Size = New Size(150, 35),
            .Enabled = False,
            .BackColor = Color.LightGreen
        }

        yPos += 45

        ' 日志文本框：Dock 到底部，避免被工具栏遮挡
        TxtLog = New TextBox With {
            .Multiline = True,
            .ScrollBars = ScrollBars.Vertical,
            .Dock = DockStyle.Bottom,
            .Height = 200,
            .ReadOnly = True,
            .Font = New Font("楷体", 10),
            .BackColor = Color.Black,
            .ForeColor = Color.White,
            .Visible = False
        }

        ' 添加到窗体（注意：TxtLog 最后添加，Dock 才能正确计算剩余空间）
        Me.Controls.AddRange({
            lblProject, TxtProjectFile, BtnBrowseProject,
            lblPublish, TxtPublishDir, BtnBrowsePublish,
            LbAd, TxtMsbuild, BtnBrowseMsbuild,
            lblRuntime, CboRuntime, ChkSelfContained, ChkSingleFile,
            BtnPublish, BtnCancel, ChkCreatExe, BtnCreateZip,
            lblCustomExeName, TxtCustomExeName, lblTip,
            lblDependencies, ClbDependencies,
            TxtLog
        })

        ' 压缩包复选框联动
        AddHandler ChkCreatExe.CheckedChanged, Sub()
                                                   BtnCreateZip.Enabled = ChkCreatExe.Checked
                                               End Sub

        ' 制作压缩包按钮点击
        AddHandler BtnCreateZip.Click, Sub()
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

        ' --- 绑定“单文件”复选框与版本号输入框的启用状态 ---
        AddHandler ChkSingleFile.CheckedChanged, Sub()
                                                     Dim isEnabled As Boolean = ChkSingleFile.Checked
                                                     lblCustomExeName.Enabled = isEnabled
                                                     TxtCustomExeName.Enabled = isEnabled
                                                     If Not isEnabled Then
                                                         TxtCustomExeName.Text = ""
                                                     Else
                                                         If Not String.IsNullOrEmpty(TxtProjectFile.Text) AndAlso System.IO.File.Exists(TxtProjectFile.Text) Then
                                                             LoadVersionFromProject(TxtProjectFile.Text)
                                                         End If
                                                     End If
                                                 End Sub
    End Sub

    ' ==================== 判断是否为忽略项 ====================
    Private Function IsIgnoredDependency(name As String) As Boolean
        If String.IsNullOrWhiteSpace(name) Then Return True

        For Each ignore In DependencyIgnoreNames
            If String.Equals(name, ignore, StringComparison.OrdinalIgnoreCase) Then
                Return True
            End If
        Next

        Dim ext As String = System.IO.Path.GetExtension(name)
        If Not String.IsNullOrEmpty(ext) Then
            For Each ignoreExt In DependencyIgnoreExtensions
                If String.Equals(ext, ignoreExt, StringComparison.OrdinalIgnoreCase) Then
                    Return True
                End If
            Next
        End If

        For Each ignoreExt In DependencyIgnoreExtensions
            If name.EndsWith(ignoreExt, StringComparison.OrdinalIgnoreCase) Then
                Return True
            End If
        Next

        Return False
    End Function

    ' ==================== 获取项目依赖（文件夹与文件） ====================
    Private Function GetProjectDependencies(projectFilePath As String) As List(Of String)
        Dim result As New List(Of String)
        Try
            If String.IsNullOrWhiteSpace(projectFilePath) OrElse Not System.IO.File.Exists(projectFilePath) Then
                Return result
            End If

            Dim projectDir As String = System.IO.Path.GetDirectoryName(projectFilePath)
            If String.IsNullOrWhiteSpace(projectDir) OrElse Not System.IO.Directory.Exists(projectDir) Then
                Return result
            End If

            Dim projectFileName As String = System.IO.Path.GetFileName(projectFilePath)

            For Each subFolderPath In System.IO.Directory.GetDirectories(projectDir)
                Dim name As String = System.IO.Path.GetFileName(subFolderPath)
                If IsIgnoredDependency(name) Then Continue For
                result.Add(subFolderPath)
            Next

            For Each file In System.IO.Directory.GetFiles(projectDir)
                Dim name As String = System.IO.Path.GetFileName(file)
                If IsIgnoredDependency(name) Then Continue For
                If String.Equals(name, projectFileName, StringComparison.OrdinalIgnoreCase) Then Continue For
                result.Add(file)
            Next
        Catch ex As Exception
            LogMessage($"[扫描项目依赖失败] {ex.Message}")
        End Try
        Return result
    End Function

    ' ==================== 刷新依赖勾选列表框 ====================
    Private Sub RefreshDependencyList(projectFilePath As String)
        Try
            If ClbDependencies Is Nothing Then Return
            ClbDependencies.Items.Clear()

            If String.IsNullOrWhiteSpace(projectFilePath) OrElse Not System.IO.File.Exists(projectFilePath) Then
                Return
            End If

            Dim deps As List(Of String) = GetProjectDependencies(projectFilePath)
            For Each item In deps
                Dim isDir As Boolean = System.IO.Directory.Exists(item)
                Dim name As String = System.IO.Path.GetFileName(item)
                Dim display As String = If(isDir, $"[文件夹] {name}", $"[文件]   {name}")
                ClbDependencies.Items.Add(New DependencyItem(item, display), True)
            Next
        Catch ex As Exception
            LogMessage($"[刷新依赖列表失败] {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' 依赖项包装类：保存完整路径与显示文本。
    ''' </summary>
    Private Class DependencyItem
        Public ReadOnly Property FullPath As String
        Public ReadOnly Property Display As String

        Public Sub New(fullPath As String, display As String)
            Me.FullPath = fullPath
            Me.Display = display
        End Sub

        Public Overrides Function ToString() As String
            Return Display
        End Function
    End Class

    ' ==================== 获取用户勾选的依赖 ====================
    Private Function GetCheckedDependencies() As List(Of String)
        Dim result As New List(Of String)
        Try
            If ClbDependencies Is Nothing Then Return result
            For i As Integer = 0 To ClbDependencies.Items.Count - 1
                If ClbDependencies.GetItemChecked(i) Then
                    Dim item As DependencyItem = TryCast(ClbDependencies.Items(i), DependencyItem)
                    If item IsNot Nothing Then
                        result.Add(item.FullPath)
                    End If
                End If
            Next
        Catch ex As Exception
            LogMessage($"[获取勾选依赖失败] {ex.Message}")
        End Try
        Return result
    End Function

    ' ==================== 日志输出勾选的依赖 ====================
    Private Sub LogCheckedDependencies()
        Try
            Dim deps As List(Of String) = GetCheckedDependencies()
            If deps.Count = 0 Then
                LogMessage("依赖检查: 本项目无依赖文件（或用户未勾选）")
                Return
            End If

            LogMessage("依赖检查: 将复制以下依赖项")
            For Each item In deps
                Dim isDir As Boolean = System.IO.Directory.Exists(item)
                Dim name As String = System.IO.Path.GetFileName(item)
                If isDir Then
                    LogMessage($"  [文件夹] {name}")
                Else
                    LogMessage($"  [文件]   {name}")
                End If
            Next
        Catch ex As Exception
            LogMessage($"[输出依赖信息失败] {ex.Message}")
        End Try
    End Sub

    ' ==================== 复制勾选的依赖到发布目录 ====================
    Private Sub CopyCheckedDependencies(publishDir As String)
        Try
            If String.IsNullOrWhiteSpace(publishDir) OrElse Not System.IO.Directory.Exists(publishDir) Then
                Return
            End If

            Dim deps As List(Of String) = GetCheckedDependencies()
            If deps.Count = 0 Then
                Return
            End If

            For Each item In deps
                Dim name As String = System.IO.Path.GetFileName(item)
                Dim targetPath As String = System.IO.Path.Combine(publishDir, name)

                If System.IO.Directory.Exists(item) Then
                    CopyDirectory(item, targetPath, True)
                    LogMessage($"已复制依赖文件夹: {name}")
                ElseIf System.IO.File.Exists(item) Then
                    System.IO.File.Copy(item, targetPath, True)
                    LogMessage($"已复制依赖文件: {name}")
                End If
            Next
        Catch ex As Exception
            LogMessage($"[复制依赖失败] {ex.Message}")
        End Try
    End Sub

    ' ==================== 递归复制目录 ====================
    Private Sub CopyDirectory(sourceDir As String, targetDir As String, overwrite As Boolean)
        Try
            If Not System.IO.Directory.Exists(targetDir) Then
                System.IO.Directory.CreateDirectory(targetDir)
            End If

            For Each file In System.IO.Directory.GetFiles(sourceDir)
                Dim fileName As String = System.IO.Path.GetFileName(file)
                Dim targetFile As String = System.IO.Path.Combine(targetDir, fileName)
                System.IO.File.Copy(file, targetFile, overwrite)
            Next

            For Each subFolderPath In System.IO.Directory.GetDirectories(sourceDir)
                Dim subFolderName As String = System.IO.Path.GetFileName(subFolderPath)
                Dim targetSubDir As String = System.IO.Path.Combine(targetDir, subFolderName)
                CopyDirectory(subFolderPath, targetSubDir, overwrite)
            Next
        Catch ex As Exception
            LogMessage($"[复制目录失败] {sourceDir} -> {targetDir}: {ex.Message}")
        End Try
    End Sub

    ' ==================== 自动查找 MSBuild 路径 ====================
    Private Function FindMsBuildPath() As String
        Try
            Dim vswhereCandidates As String() = {
                System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86), "Microsoft Visual Studio", "Installer", "vswhere.exe"),
                System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), "Microsoft Visual Studio", "Installer", "vswhere.exe")
            }

            For Each vswhere As String In vswhereCandidates
                If System.IO.File.Exists(vswhere) Then
                    Dim psi As New ProcessStartInfo With {
                        .FileName = vswhere,
                        .Arguments = "-latest -requires Microsoft.Component.MSBuild -find MSBuild\**\Bin\MSBuild.exe",
                        .UseShellExecute = False,
                        .RedirectStandardOutput = True,
                        .CreateNoWindow = True,
                        .StandardOutputEncoding = System.Text.Encoding.UTF8
                    }
                    Using p As Process = Process.Start(psi)
                        Dim output As String = p.StandardOutput.ReadToEnd()
                        p.WaitForExit()
                        If Not String.IsNullOrWhiteSpace(output) Then
                            Dim firstLine As String = output.Split({ControlChars.Cr, ControlChars.Lf}, StringSplitOptions.RemoveEmptyEntries).FirstOrDefault()
                            If Not String.IsNullOrWhiteSpace(firstLine) AndAlso System.IO.File.Exists(firstLine.Trim()) Then
                                LogMessage($"[自动获取MSBuild] {firstLine.Trim()}")
                                Return firstLine.Trim()
                            End If
                        End If
                    End Using
                End If
            Next

            Dim vsRoots As String() = {
                System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), "Microsoft Visual Studio"),
                System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86), "Microsoft Visual Studio")
            }
            For Each root In vsRoots
                If System.IO.Directory.Exists(root) Then
                    Dim found As String = FindFileInDirectory(root, "MSBuild.exe", 6)
                    If Not String.IsNullOrEmpty(found) Then
                        LogMessage($"[自动获取MSBuild] {found}")
                        Return found
                    End If
                End If
            Next

            Dim dotnetRoot As String = System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), "dotnet", "sdk")
            If System.IO.Directory.Exists(dotnetRoot) Then
                Dim sdkDirs As String() = System.IO.Directory.GetDirectories(dotnetRoot)
                Array.Sort(sdkDirs)
                Array.Reverse(sdkDirs)
                For Each sdkDir In sdkDirs
                    Dim candidate As String = System.IO.Path.Combine(sdkDir, "MSBuild.dll")
                    If System.IO.File.Exists(candidate) Then
                        Dim exeCandidate As String = System.IO.Path.Combine(sdkDir, "MSBuild.exe")
                        If System.IO.File.Exists(exeCandidate) Then
                            LogMessage($"[自动获取MSBuild] {exeCandidate}")
                            Return exeCandidate
                        End If
                    End If
                Next
            End If

            Dim envMsbuild As String = Environment.GetEnvironmentVariable("MSBuild")
            If Not String.IsNullOrWhiteSpace(envMsbuild) AndAlso System.IO.File.Exists(envMsbuild) Then
                LogMessage($"[自动获取MSBuild] {envMsbuild}")
                Return envMsbuild
            End If

            Dim pathVar As String = Environment.GetEnvironmentVariable("PATH")
            If Not String.IsNullOrWhiteSpace(pathVar) Then
                For Each pathItem In pathVar.Split(System.IO.Path.PathSeparator)
                    Try
                        Dim candidate As String = System.IO.Path.Combine(pathItem.Trim(), "MSBuild.exe")
                        If System.IO.File.Exists(candidate) Then
                            LogMessage($"[自动获取MSBuild] {candidate}")
                            Return candidate
                        End If
                    Catch
                    End Try
                Next
            End If
        Catch ex As Exception
            LogMessage($"[自动获取MSBuild失败] {ex.Message}")
        End Try

        Return ""
    End Function

    ' ==================== 在目录中递归查找文件 ====================
    Private Function FindFileInDirectory(rootDir As String, fileName As String, maxDepth As Integer) As String
        Try
            If maxDepth < 0 Then Return ""
            Dim files As String() = System.IO.Directory.GetFiles(rootDir, fileName)
            If files.Length > 0 Then
                Return files(0)
            End If
            If maxDepth = 0 Then Return ""
            For Each subFolderPath In System.IO.Directory.GetDirectories(rootDir)
                Dim found As String = FindFileInDirectory(subFolderPath, fileName, maxDepth - 1)
                If Not String.IsNullOrEmpty(found) Then Return found
            Next
        Catch
        End Try
        Return ""
    End Function

    ' ==================== 从项目文件读取版本号 ====================
    Private Sub LoadVersionFromProject(projectFilePath As String)
        Try
            If String.IsNullOrWhiteSpace(projectFilePath) OrElse Not System.IO.File.Exists(projectFilePath) Then
                Return
            End If

            Dim version As String = EnsureProjectVersion(projectFilePath)
            If Not String.IsNullOrWhiteSpace(version) Then
                If ChkSingleFile.Checked Then
                    TxtCustomExeName.Text = version
                End If
            End If
        Catch ex As Exception
            LogMessage($"[读取项目版本号失败] {ex.Message}")
        End Try
    End Sub

    ' ==================== 确保项目文件存在 <Version> 节点 ====================
    Private Function EnsureProjectVersion(projectFilePath As String) As String
        Const DefaultVersion As String = "1.0.0.0"

        Try
            Dim doc As New System.Xml.XmlDocument()
            doc.PreserveWhitespace = True
            doc.Load(projectFilePath)

            Dim versionNode As System.Xml.XmlNode = doc.SelectSingleNode("//PropertyGroup/Version")
            If versionNode IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(versionNode.InnerText) Then
                Return versionNode.InnerText.Trim()
            End If

            Dim targetGroup As System.Xml.XmlNode = Nothing
            Dim groups As System.Xml.XmlNodeList = doc.SelectNodes("//PropertyGroup")
            If groups IsNot Nothing Then
                For Each g As System.Xml.XmlNode In groups
                    Dim cond As System.Xml.XmlAttribute = g.Attributes("Condition")
                    If cond Is Nothing Then
                        targetGroup = g
                        Exit For
                    End If
                Next
                If targetGroup Is Nothing AndAlso groups.Count > 0 Then
                    targetGroup = groups(0)
                End If
            End If

            If targetGroup Is Nothing Then
                Dim projectRoot As System.Xml.XmlNode = doc.SelectSingleNode("//Project")
                If projectRoot Is Nothing Then
                    Return DefaultVersion
                End If
                targetGroup = doc.CreateElement("PropertyGroup")
                projectRoot.AppendChild(targetGroup)
            End If

            Dim newVersionNode As System.Xml.XmlElement = doc.CreateElement("Version")
            newVersionNode.InnerText = DefaultVersion
            targetGroup.AppendChild(newVersionNode)

            doc.Save(projectFilePath)

            LogMessage($"[已自动写入版本号] {DefaultVersion}")
            Return DefaultVersion
        Catch ex As Exception
            LogMessage($"[处理项目版本号失败] {ex.Message}")
            Return DefaultVersion
        End Try
    End Function

    ' ==================== 浏览项目文件 ====================
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
                LoadVersionFromProject(dialog.FileName)
                RefreshDependencyList(dialog.FileName)
                SaveLastParts()
            End If
        End Using
    End Sub

    ' ==================== 浏览 MSBuild ====================
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

    ' ==================== 浏览发布目录 ====================
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

    ' ==================== 开始发布按钮 ====================
    Private Sub BtnPublish_Click(sender As Object, e As EventArgs) Handles BtnPublish.Click

        SaveLastParts()

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

        If String.IsNullOrWhiteSpace(TxtMsbuild.Text) OrElse Not System.IO.File.Exists(TxtMsbuild.Text) Then
            Dim autoMsbuild As String = FindMsBuildPath()
            If Not String.IsNullOrEmpty(autoMsbuild) Then
                TxtMsbuild.Text = autoMsbuild
                SaveLastParts()
            End If
        End If

        If Not ConfirmDependencies() Then
            Return
        End If

NextStep:
        ProjectPulish()
    End Sub

    ' ==================== 发布前依赖确认 ====================
    Private Function ConfirmDependencies() As Boolean
        Try
            If ClbDependencies.Items.Count = 0 AndAlso Not String.IsNullOrEmpty(TxtProjectFile.Text) Then
                RefreshDependencyList(TxtProjectFile.Text)
            End If

            Dim msg As String
            If ClbDependencies.Items.Count = 0 Then
                msg = "未发现可复制的依赖项。" & vbCrLf & vbCrLf & "是否继续发布？"
            Else
                Dim checkedCount As Integer = GetCheckedDependencies().Count
                msg = $"已发现 {ClbDependencies.Items.Count} 个依赖项，其中 {checkedCount} 个已勾选。" & vbCrLf & vbCrLf &
                      "请在主界面的"“运行时依赖项”"列表中进行二次确认（默认已全部勾选）。" & vbCrLf & vbCrLf &
                      "确认后点击"“是”"继续发布，点击"“否”"返回调整。"
            End If

            Dim result As DialogResult = MessageBox.Show(msg, "依赖确认", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
            Return result = DialogResult.Yes
        Catch ex As Exception
            LogMessage($"[依赖确认失败] {ex.Message}")
            Return True
        End Try
    End Function

    ' ==================== 项目发布主流程 ====================
    Private Sub ProjectPulish()
        Dim projectFile As String = TxtProjectFile.Text
        Dim publishDir As String = TxtPublishDir.Text
        Dim runtimeId As String = If(CboRuntime.SelectedItem?.ToString(), "win-x86")
        Dim projectName As String = System.IO.Path.GetFileName(projectFile)

        Dim baseExeName As String = System.IO.Path.GetFileNameWithoutExtension(projectFile)

        Dim versionText As String = TxtCustomExeName.Text.Trim()
        If ChkSingleFile.Checked AndAlso String.IsNullOrWhiteSpace(versionText) Then
            versionText = EnsureProjectVersion(projectFile)
            TxtCustomExeName.Text = versionText
        End If

        ' 构建 MSBuild 参数
        ' 新增：/p:DebugType=None /p:DebugSymbols=false 让发布不产生 .pdb
        Dim arguments As String = $"""{projectFile}"" " &
                          "/t:Restore;Publish " &
                          "/p:Configuration=Release " &
                          "/p:DebugType=None " &
                          "/p:DebugSymbols=false " &
                          $"/p:SelfContained={If(ChkSelfContained.Checked, "true", "false")} " &
                          $"/p:PublishSingleFile={If(ChkSingleFile.Checked, "true", "false")} " &
                          $"/p:RuntimeIdentifier={runtimeId} " &
                          "/p:IncludeNativeLibrariesForSelfExtract=true " &
                          "/p:SignManifests=false " &
                          "/p:SignAssembly=false " &
                          "/p:GenerateClickOnceManifests=false " &
                          "/p:BootstrapperEnabled=false " &
                          $"/p:PublishDir=""{publishDir}"" " &
                          "/verbosity:minimal " &
                          "/nologo"

        ' 程序名称固定为项目基础名，不带版本号
        arguments = $"/p:AssemblyName={baseExeName} " & arguments

        Dim msbuildPath As String = TxtMsbuild.Text
        If String.IsNullOrEmpty(msbuildPath) Then
            MessageBox.Show("请重新选择Msbuild程序路径", "错误", MessageBoxButtons.OK)
            Return
        End If

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

        TxtLog.Visible = True
        TxtLog.Clear()

        ' ========== 发布前：列出用户勾选的依赖 ==========
        LogCheckedDependencies()
        LogMessage("")

        ' ========== 清理旧版本 ==========
        Try
            Dim newExeFullName As String = $"{baseExeName}.exe"

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

        Dim cur As Date = DateTime.Now
        BtnPublish.Enabled = False
        BtnCancel.Enabled = False

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

                While Not process.HasExited
                    Application.DoEvents()
                    Threading.Thread.Sleep(100)
                End While

                LogMessage("")
                LogMessage($"=== 发布完成 ===")
                LogMessage($"退出代码: {process.ExitCode}")

                Dim targetExePattern As String = $"{baseExeName}.exe"
                Dim generatedExe As String = System.IO.Directory.GetFiles(publishDir, targetExePattern).FirstOrDefault()

                If Not String.IsNullOrEmpty(generatedExe) Then
                    LogMessage($"状态: 成功 ✓ (主程序已生成: {System.IO.Path.GetFileName(generatedExe)})")

                    ' ========== 发布后：复制用户勾选的依赖到发布目录（直接覆盖） ==========
                    CopyCheckedDependencies(publishDir)

                    LogMessage($"总耗时:{DateDiff("s", cur, Now)}s")

                    Dim createShortcutResult As DialogResult = MessageBox.Show(
                        "发布成功！是否在桌面创建快捷方式？",
                        "创建快捷方式",
                        MessageBoxButtons.YesNo,
                        MessageBoxIcon.Question)

                    If createShortcutResult = DialogResult.Yes Then
                        Try
                            Dim desktopPath As String = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)

                            Dim targetPath As String = ""
                            Dim expectedExePath As String = System.IO.Path.Combine(publishDir, $"{baseExeName}.exe")

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

                            Dim defaultName As String = baseExeName
                            Dim shortcutName As String = InputBox("请输入桌面快捷方式的名称（包含.lnk后缀）：", "快捷方式命名", defaultName & ".lnk")

                            If String.IsNullOrWhiteSpace(shortcutName) Then
                                shortcutName = defaultName & ".lnk"
                            End If

                            If Not shortcutName.EndsWith(".lnk", StringComparison.OrdinalIgnoreCase) Then
                                shortcutName &= ".lnk"
                            End If

                            Dim shortcutPath As String = System.IO.Path.Combine(desktopPath, shortcutName)

                            Dim shellType As Type = Type.GetTypeFromProgID("WScript.Shell")
                            Dim shell As Object = Activator.CreateInstance(shellType)
                            Dim shortcut As Object = shellType.InvokeMember("CreateShortcut",
                                                                          System.Reflection.BindingFlags.InvokeMethod,
                                                                          Nothing,
                                                                          shell,
                                                                          New Object() {shortcutPath})

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
            Try
                ' 注意：已通过 /p:DebugType=None 禁止生成 .pdb，这里不再清理 *.pdb
                Dim extraFiles As String() = System.IO.Directory.GetFiles(publishDir, "setup.exe") _
                        .Concat(System.IO.Directory.GetFiles(publishDir, "*.application")) _
                        .Concat(System.IO.Directory.GetFiles(publishDir, "*.manifest")) _
                        .Concat(System.IO.Directory.GetFiles(publishDir, "*.tmp")) _
                        .ToArray()
                For Each f In extraFiles
                    If System.IO.File.Exists(f) Then
                        System.IO.File.Delete(f)
                        LogMessage($"已删除多余文件: {System.IO.Path.GetFileName(f)}")
                    End If
                Next

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

    ' ==================== 日志输出 ====================
    Private Sub LogMessage(message As String)
        TxtLog.AppendText($"[{DateTime.Now:HH:mm:ss}] {message}" & vbCrLf)
        TxtLog.ScrollToCaret()
    End Sub

    ' ==================== 取消按钮 ====================
    Private Sub BtnCancel_Click(sender As Object, e As EventArgs) Handles BtnCancel.Click
        Me.Close()
    End Sub

    ' ==================== 支持拖放项目文件到窗体 ====================
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
                LoadVersionFromProject(file)
                RefreshDependencyList(file)
            End If
        End If
    End Sub

    ' ==================== 支持拖放项目文件到文本框 ====================
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
                LoadVersionFromProject(file)
                RefreshDependencyList(file)
            End If
        End If
    End Sub

    ' ==================== 创建调试输出压缩包 ====================
    Private Sub CreateDebugOutputPackage(projectName As String, publishDir As String)
        If Not (System.IO.Directory.Exists(publishDir) OrElse System.IO.File.Exists(publishDir)) Then
            LogMessage($"[错误] 发布路径不存在: {publishDir}")
            Return
        End If

        Dim outputDir As String = System.IO.Path.GetDirectoryName(publishDir)
        Dim zipFile As String = System.IO.Path.Combine(outputDir, $"{projectName}.zip")

        If System.IO.File.Exists(zipFile) Then
            System.IO.File.Delete(zipFile)
            LogMessage($"删除旧文件: {zipFile}")
        End If

        If System.IO.Directory.Exists(publishDir) Then
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
            If System.IO.File.Exists(publishDir) Then
                LogMessage($"正在打包文件: {publishDir}")
                Try
                    Using zipStream As New System.IO.FileStream(zipFile, System.IO.FileMode.Create)
                        Using zipArchive As New System.IO.Compression.ZipArchive(zipStream, System.IO.Compression.ZipArchiveMode.Create)
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