Imports System.Drawing
Imports System.Drawing.Drawing2D
Imports System.IO
Imports System.Windows.Forms

Public Class Form1
    ' 控件声明
    Private WithEvents picPreview As PictureBox
    Private WithEvents txtFilePath As TextBox
    Private WithEvents btnGenerate As Button
    Private WithEvents btnExtract As Button
    Private WithEvents btnLoad As Button
    Private WithEvents btnBrowse As Button
    Private WithEvents lblDragDrop As Label

    ' 数据存储
    Private imageBytes As Byte()
    Private fileBytes As Byte()

    ' 文件对话框
    Private OpenFileDialog1 As New OpenFileDialog()
    Private SaveFileDialog1 As New SaveFileDialog()

    ' 状态栏
    Private statusBar As StatusStrip
    Private lblStatus As ToolStripStatusLabel
    Private progressBar As ToolStripProgressBar

    Public Sub New()
        ' 初始化窗体组件
        InitializeComponent()

        ' 自定义UI初始化
        InitializeUI()
    End Sub

    Private Sub InitializeUI()
        ' ========== 窗体基本设置 ==========
        Me.Text = "图种生成器 v114514"
        Me.Size = New Size(900, 650)
        Me.MinimumSize = New Size(700, 500)
        Me.StartPosition = FormStartPosition.CenterScreen
        Me.Font = New Font("微软雅黑", 9)
        Me.BackColor = Color.FromArgb(240, 240, 240)

        ' 如果没有图标资源，可以注释掉这行
        ' Me.Icon = My.Resources.AppIcon

        ' ========== 主容器布局 ==========
        Dim mainTable As New TableLayoutPanel With {
            .Dock = DockStyle.Fill,
            .ColumnCount = 2,
            .RowCount = 3,
            .CellBorderStyle = TableLayoutPanelCellBorderStyle.None,
            .Padding = New Padding(10),
            .BackColor = Color.Transparent
        }
        mainTable.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 60))
        mainTable.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 40))
        mainTable.RowStyles.Add(New RowStyle(SizeType.Absolute, 60))
        mainTable.RowStyles.Add(New RowStyle(SizeType.Percent, 100))
        mainTable.RowStyles.Add(New RowStyle(SizeType.Absolute, 40))

        Me.Controls.Add(mainTable)

        ' ========== 标题区域 ==========
        Dim pnlHeader As New Panel With {
            .Dock = DockStyle.Fill,
            .BackColor = Color.FromArgb(70, 130, 180)
        }

        Dim lblTitle As New Label With {
            .Text = "图种生成器",
            .Font = New Font("微软雅黑", 16, FontStyle.Bold),
            .ForeColor = Color.White,
            .Dock = DockStyle.Fill,
            .TextAlign = ContentAlignment.MiddleCenter
        }

        pnlHeader.Controls.Add(lblTitle)
        mainTable.Controls.Add(pnlHeader, 0, 0)
        mainTable.SetColumnSpan(pnlHeader, 2)

        ' ========== 左侧图片预览区 ==========
        Dim pnlPreview As New Panel With {
            .Dock = DockStyle.Fill,
            .BorderStyle = BorderStyle.FixedSingle,
            .BackColor = Color.White,
            .Padding = New Padding(5)
        }

        picPreview = New PictureBox With {
            .SizeMode = PictureBoxSizeMode.Zoom,
            .Dock = DockStyle.Fill,
            .BackColor = Color.White
        }

        lblDragDrop = New Label With {
            .Text = "拖放图片到此处或点击下方按钮选择",
            .Font = New Font("微软雅黑", 10, FontStyle.Italic),
            .ForeColor = Color.Gray,
            .TextAlign = ContentAlignment.MiddleCenter,
            .Dock = DockStyle.Fill,
            .AutoSize = False
        }

        pnlPreview.Controls.Add(picPreview)
        pnlPreview.Controls.Add(lblDragDrop)
        picPreview.BringToFront()

        ' 启用拖放功能
        pnlPreview.AllowDrop = True
        AddHandler pnlPreview.DragEnter, AddressOf Panel_DragEnter
        AddHandler pnlPreview.DragDrop, AddressOf Panel_DragDrop

        mainTable.Controls.Add(pnlPreview, 0, 1)

        ' ========== 右侧控制面板 ==========
        Dim pnlControl As New Panel With {
            .Dock = DockStyle.Fill,
            .Padding = New Padding(15, 10, 15, 10),
            .BackColor = Color.White
        }

        ' 使用垂直滚动面板
        Dim scrollPanel As New ScrollableControl With {
            .Dock = DockStyle.Fill,
            .AutoScroll = True
        }

        ' 控制面板内容容器
        Dim contentPanel As New TableLayoutPanel With {
            .Dock = DockStyle.Top,
            .AutoSize = True,
            .AutoSizeMode = AutoSizeMode.GrowAndShrink,
            .ColumnCount = 1,
            .RowCount = 4,
            .Margin = New Padding(0)
        }
        ' 1. 图片选择部分
        Dim grpImage As New GroupBox With {
    .Text = " 1. 选择图片 ",
    .Dock = DockStyle.Top,
    .Font = New Font("微软雅黑", 10, FontStyle.Bold),
    .ForeColor = Color.FromArgb(70, 130, 180),
    .Height = 100,
    .Margin = New Padding(0, 0, 0, 15),
    .Padding = New Padding(10)
}

        btnLoad = CreateStyledButton("选择图片...", Color.FromArgb(100, 149, 237))
        With btnLoad
            .Height = 40
            .Width = 250  ' 设置固定宽度，使按钮更长
            .Margin = New Padding(0, 15, 0, 0)
            .Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right  ' 添加锚点使按钮能随窗体调整
        End With
        AddHandler btnLoad.Click, AddressOf btnLoadImage_Click

        Dim lblImageInfo As New Label With {
    .Text = "支持格式: JPG, PNG, BMP",
    .Font = New Font("微软雅黑", 8),
    .ForeColor = Color.Gray,
    .Dock = DockStyle.Bottom,
    .TextAlign = ContentAlignment.MiddleLeft,
    .Height = 20,
    .Margin = New Padding(0, 5, 0, 0)
}

        ' 使用面板控制内部布局
        Dim pnlImageContent As New Panel With {
    .Dock = DockStyle.Fill,
    .Padding = New Padding(5)
}

        ' 使按钮在面板中水平居中
        btnLoad.Left = (pnlImageContent.Width - btnLoad.Width) \ 2

        pnlImageContent.Controls.Add(btnLoad)
        pnlImageContent.Controls.Add(lblImageInfo)

        grpImage.Controls.Add(pnlImageContent)
        contentPanel.Controls.Add(grpImage)
        ' 2. 文件选择部分
        Dim grpFile As New GroupBox With {
            .Text = " 2. 选择要隐藏的文件 ",
            .Dock = DockStyle.Top,
            .Font = New Font("微软雅黑", 10, FontStyle.Bold),
            .ForeColor = Color.FromArgb(70, 130, 180),
            .Height = 140,
            .Margin = New Padding(0, 0, 0, 15)
        }

        txtFilePath = New TextBox With {
            .Dock = DockStyle.Top,
            .Height = 30,
            .Margin = New Padding(0, 0, 0, 10),
            .ReadOnly = True,
            .Font = New Font("微软雅黑", 9),
            .BackColor = Color.WhiteSmoke
        }

        btnBrowse = CreateStyledButton("浏览文件...", Color.FromArgb(147, 112, 219))
        btnBrowse.Height = 40
        btnBrowse.Dock = DockStyle.Top
        AddHandler btnBrowse.Click, AddressOf btnBrowseFile_Click

        Dim lblFileInfo As New Label With {
            .Text = "支持任何文件类型",
            .Font = New Font("微软雅黑", 8),
            .ForeColor = Color.Gray,
            .Dock = DockStyle.Bottom,
            .TextAlign = ContentAlignment.MiddleLeft,
            .Height = 20
        }

        grpFile.Controls.Add(btnBrowse)
        grpFile.Controls.Add(txtFilePath)
        grpFile.Controls.Add(lblFileInfo)
        contentPanel.Controls.Add(grpFile)

        ' 3. 操作按钮部分
        Dim grpActions As New GroupBox With {
            .Text = " 3. 执行操作 ",
            .Dock = DockStyle.Top,
            .Font = New Font("微软雅黑", 10, FontStyle.Bold),
            .ForeColor = Color.FromArgb(70, 130, 180),
            .Height = 180,
            .Margin = New Padding(0, 0, 0, 15)
        }

        btnGenerate = CreateStyledButton("生成图种", Color.FromArgb(50, 205, 50))
        btnGenerate.Height = 45
        btnGenerate.Dock = DockStyle.Top
        btnGenerate.Margin = New Padding(0, 0, 0, 10)
        AddHandler btnGenerate.Click, AddressOf btnGenerate_Click

        btnExtract = CreateStyledButton("从图种提取文件", Color.FromArgb(30, 144, 255))
        btnExtract.Height = 45
        btnExtract.Dock = DockStyle.Top
        AddHandler btnExtract.Click, AddressOf btnExtract_Click

        grpActions.Controls.Add(btnGenerate)
        grpActions.Controls.Add(btnExtract)
        contentPanel.Controls.Add(grpActions)

        ' 4. 关于信息
        Dim grpAbout As New GroupBox With {
            .Text = " 关于 ",
            .Dock = DockStyle.Top,
            .Font = New Font("微软雅黑", 10, FontStyle.Bold),
            .ForeColor = Color.FromArgb(70, 130, 180),
            .Height = 80
        }

        Dim lblAbout As New Label With {
            .Text = "图种生成器 v114514" & vbCrLf & "将文件隐藏到图片中" & vbCrLf & "by xrlzu",
            .Font = New Font("微软雅黑", 8),
            .ForeColor = Color.DimGray,
            .Dock = DockStyle.Fill,
            .TextAlign = ContentAlignment.MiddleCenter
        }

        grpAbout.Controls.Add(lblAbout)
        contentPanel.Controls.Add(grpAbout)

        scrollPanel.Controls.Add(contentPanel)
        pnlControl.Controls.Add(scrollPanel)
        mainTable.Controls.Add(pnlControl, 1, 1)

        ' ========== 底部状态栏 ==========
        statusBar = New StatusStrip With {
            .Dock = DockStyle.Fill,
            .RenderMode = ToolStripRenderMode.ManagerRenderMode,
            .BackColor = Color.FromArgb(240, 240, 240)
        }

        lblStatus = New ToolStripStatusLabel With {
            .Text = "就绪",
            .Spring = True,
            .TextAlign = ContentAlignment.MiddleLeft
        }

        progressBar = New ToolStripProgressBar With {
            .Visible = False,
            .Style = ProgressBarStyle.Continuous
        }

        statusBar.Items.Add(lblStatus)
        statusBar.Items.Add(progressBar)
        mainTable.Controls.Add(statusBar, 0, 2)
        mainTable.SetColumnSpan(statusBar, 2)

        ' ========== 配置文件对话框 ==========
        OpenFileDialog1.Filter = "图片文件|*.jpg;*.jpeg;*.png;*.bmp|所有文件|*.*"
        SaveFileDialog1.Filter = "JPEG 图片|*.jpg|PNG 图片|*.png|BMP 图片|*.bmp"
    End Sub

    ' ========== 自定义美化方法 ==========

    ' 创建带样式的按钮
    Private Function CreateStyledButton(text As String, backColor As Color) As Button
        Dim btn As New Button With {
            .Text = text,
            .Font = New Font("微软雅黑", 10, FontStyle.Bold),
            .ForeColor = Color.White,
            .BackColor = backColor,
            .FlatStyle = FlatStyle.Flat,
            .Cursor = Cursors.Hand
        }

        btn.FlatAppearance.BorderSize = 0
        btn.FlatAppearance.MouseOverBackColor = ControlPaint.Light(backColor, 0.2)
        btn.FlatAppearance.MouseDownBackColor = ControlPaint.Dark(backColor, 0.2)

        ' 添加悬停动画效果
        AddHandler btn.MouseEnter, Sub(sender, e)
                                       Dim b = CType(sender, Button)
                                       b.Font = New Font(b.Font, FontStyle.Underline Or FontStyle.Bold)
                                   End Sub

        AddHandler btn.MouseLeave, Sub(sender, e)
                                       Dim b = CType(sender, Button)
                                       b.Font = New Font(b.Font, FontStyle.Bold)
                                   End Sub

        Return btn
    End Function

    ' ========== 拖放功能实现 ==========

    Private Sub Panel_DragEnter(sender As Object, e As DragEventArgs)
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            e.Effect = DragDropEffects.Copy
        Else
            e.Effect = DragDropEffects.None
        End If
    End Sub

    Private Sub Panel_DragDrop(sender As Object, e As DragEventArgs)
        Dim files() As String = CType(e.Data.GetData(DataFormats.FileDrop), String())
        If files.Length > 0 Then
            Dim ext As String = Path.GetExtension(files(0)).ToLower()
            If {".jpg", ".jpeg", ".png", ".bmp"}.Contains(ext) Then
                Try
                    picPreview.Image = Image.FromFile(files(0))
                    imageBytes = File.ReadAllBytes(files(0))
                    lblDragDrop.Visible = False
                    lblStatus.Text = $"已加载图片: {Path.GetFileName(files(0))}"
                Catch ex As Exception
                    MessageBox.Show($"加载图片失败: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    lblStatus.Text = "图片加载失败"
                End Try
            Else
                MessageBox.Show("请选择有效的图片文件 (JPG, PNG, BMP)", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
        End If
    End Sub

    ' ========== 按钮点击事件 ==========

    Private Sub btnLoadImage_Click(sender As Object, e As EventArgs)
        OpenFileDialog1.Filter = "图片文件|*.jpg;*.jpeg;*.png;*.bmp"
        If OpenFileDialog1.ShowDialog() = DialogResult.OK Then
            Try
                picPreview.Image = Image.FromFile(OpenFileDialog1.FileName)
                imageBytes = File.ReadAllBytes(OpenFileDialog1.FileName)
                lblDragDrop.Visible = False
                lblStatus.Text = $"已加载图片: {Path.GetFileName(OpenFileDialog1.FileName)}"
            Catch ex As Exception
                MessageBox.Show($"加载图片失败: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
                lblStatus.Text = "图片加载失败"
            End Try
        End If
    End Sub

    Private Sub btnBrowseFile_Click(sender As Object, e As EventArgs)
        OpenFileDialog1.Filter = "所有文件|*.*"
        If OpenFileDialog1.ShowDialog() = DialogResult.OK Then
            txtFilePath.Text = OpenFileDialog1.FileName
            fileBytes = File.ReadAllBytes(OpenFileDialog1.FileName)
            lblStatus.Text = $"已选择文件: {Path.GetFileName(OpenFileDialog1.FileName)} (大小: {FormatFileSize(fileBytes.Length)})"
        End If
    End Sub

    Private Sub btnGenerate_Click(sender As Object, e As EventArgs)
        If imageBytes Is Nothing OrElse imageBytes.Length = 0 Then
            MessageBox.Show("请先选择一张图片!", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            lblStatus.Text = "请先选择图片"
            Return
        End If

        If fileBytes Is Nothing OrElse fileBytes.Length = 0 Then
            MessageBox.Show("请先选择要隐藏的文件!", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            lblStatus.Text = "请先选择要隐藏的文件"
            Return
        End If

        ' 设置保存文件对话框
        SaveFileDialog1.FileName = Path.GetFileNameWithoutExtension(OpenFileDialog1.FileName) & "_hidden" & Path.GetExtension(OpenFileDialog1.FileName)
        If SaveFileDialog1.ShowDialog() = DialogResult.OK Then
            Try
                ' 合并文件（传入原始文件名）
                Dim mergedBytes As Byte() = MergeFiles(imageBytes, fileBytes, Path.GetFileName(OpenFileDialog1.FileName))

                ' 保存文件
                File.WriteAllBytes(SaveFileDialog1.FileName, mergedBytes)

                ' 显示成功信息
                Dim originalSize = imageBytes.Length
                Dim newSize = mergedBytes.Length
                Dim addedSize = newSize - originalSize

                MessageBox.Show($"图种生成成功!{Environment.NewLine}" &
                                $"原始图片大小: {FormatFileSize(originalSize)}{Environment.NewLine}" &
                                $"隐藏文件大小: {FormatFileSize(fileBytes.Length)}{Environment.NewLine}" &
                                $"生成图种大小: {FormatFileSize(newSize)}{Environment.NewLine}" &
                                $"增加大小: {FormatFileSize(addedSize)}",
                                "成功", MessageBoxButtons.OK, MessageBoxIcon.Information)

                lblStatus.Text = $"图种生成成功: {Path.GetFileName(SaveFileDialog1.FileName)}"
            Catch ex As Exception
                MessageBox.Show($"生成图种时出错: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
                lblStatus.Text = "图种生成失败"
            End Try
        End If
    End Sub

    Private Sub btnExtract_Click(sender As Object, e As EventArgs) Handles btnExtract.Click
        ' 1. 首先选择包含隐藏文件的图片
        Dim imageDialog As New OpenFileDialog With {
        .Filter = "图片文件|*.jpg;*.jpeg;*.png;*.bmp",
        .Title = "选择包含隐藏文件的图片"
    }

        If imageDialog.ShowDialog() <> DialogResult.OK Then
            Return ' 用户取消了图片选择
        End If

        Try
            ' 2. 读取图片数据并提取隐藏文件
            Dim imageData As Byte() = File.ReadAllBytes(imageDialog.FileName)
            Dim originalFileName As String = ""
            Dim fileData As Byte() = ExtractFile(imageData, originalFileName)

            If fileData Is Nothing Then
                MessageBox.Show("未找到隐藏文件!", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information)
                lblStatus.Text = "未找到隐藏文件"
                Return
            End If

            ' 3. 设置保存文件对话框
            Dim saveDialog As New SaveFileDialog With {
            .Filter = "所有文件|*.*",
            .Title = "保存提取的文件"
        }

            ' 使用原始文件名作为默认文件名
            If Not String.IsNullOrEmpty(originalFileName) Then
                saveDialog.FileName = originalFileName
            Else
                saveDialog.FileName = "extracted_file" & GetDefaultExtension(fileData)
            End If

            ' 4. 让用户选择保存位置
            If saveDialog.ShowDialog() = DialogResult.OK Then
                File.WriteAllBytes(saveDialog.FileName, fileData)
                MessageBox.Show($"文件提取成功!{Environment.NewLine}文件大小: {FormatFileSize(fileData.Length)}",
                          "成功", MessageBoxButtons.OK, MessageBoxIcon.Information)
                lblStatus.Text = $"文件提取成功: {Path.GetFileName(saveDialog.FileName)}"
            End If

        Catch ex As Exception
            MessageBox.Show($"提取文件时出错: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
            lblStatus.Text = "文件提取失败"
        Finally
            ' 5. 确保对话框资源被释放
            imageDialog.Dispose()
        End Try
    End Sub

    ' 辅助方法：根据文件内容猜测扩展名
    Private Function GetDefaultExtension(fileData As Byte()) As String
        ' 简单的文件类型检测
        If fileData.Length >= 4 Then
            ' ZIP文件qq
            If fileData(0) = &H50 AndAlso fileData(1) = &H4B AndAlso fileData(2) = &H3 AndAlso fileData(3) = &H4 Then
                Return ".zip"
            End If
            ' PDF文件
            If fileData(0) = &H25 AndAlso fileData(1) = &H50 AndAlso fileData(2) = &H44 AndAlso fileData(3) = &H46 Then
                Return ".pdf"
            End If
        End If
        Return ".dat" ' 默认扩展名
    End Function

    ' ========== 核心功能方法 ==========

    ' 合并图片和文件的函数
    Private Function MergeFiles(imageData As Byte(), fileData As Byte(), originalFileName As String) As Byte()
        ' 创建一个标记来标识文件数据的开始
        Dim marker As Byte() = System.Text.Encoding.ASCII.GetBytes("FILE_START_HIDDEN_DATA")

        ' 获取文件名和扩展名
        Dim fileNameBytes As Byte() = System.Text.Encoding.UTF8.GetBytes(originalFileName)
        Dim fileNameLengthBytes As Byte() = BitConverter.GetBytes(fileNameBytes.Length)

        ' 创建一个内存流来存储合并后的数据
        Using ms As New MemoryStream()
            ' 写入图片数据
            ms.Write(imageData, 0, imageData.Length)

            ' 写入标记
            ms.Write(marker, 0, marker.Length)

            ' 写入文件名长度(4字节)
            ms.Write(fileNameLengthBytes, 0, fileNameLengthBytes.Length)

            ' 写入文件名
            ms.Write(fileNameBytes, 0, fileNameBytes.Length)

            ' 写入文件长度(8字节)
            Dim lengthBytes As Byte() = BitConverter.GetBytes(fileData.LongLength)
            ms.Write(lengthBytes, 0, lengthBytes.Length)

            ' 写入文件数据
            ms.Write(fileData, 0, fileData.Length)

            Return ms.ToArray()
        End Using
    End Function

    ' 从图种中提取文件的函数
    Private Function ExtractFile(imageData As Byte(), ByRef originalFileName As String) As Byte()
        Dim marker As Byte() = System.Text.Encoding.ASCII.GetBytes("FILE_START_HIDDEN_DATA")
        Dim markerIndex As Integer = IndexOfSequence(imageData, marker)

        If markerIndex = -1 Then Return Nothing

        ' 读取文件名长度(4字节)
        Dim fileNameLengthOffset As Integer = markerIndex + marker.Length
        Dim fileNameLength As Integer = BitConverter.ToInt32(imageData, fileNameLengthOffset)

        ' 读取文件名
        Dim fileNameOffset As Integer = fileNameLengthOffset + 4
        originalFileName = System.Text.Encoding.UTF8.GetString(imageData, fileNameOffset, fileNameLength)

        ' 读取文件长度(8字节)
        Dim lengthOffset As Integer = fileNameOffset + fileNameLength
        Dim fileLength As Long = BitConverter.ToInt64(imageData, lengthOffset)

        ' 检查文件长度是否合理
        If fileLength <= 0 OrElse lengthOffset + 8 + fileLength > imageData.Length Then
            Return Nothing
        End If

        ' 读取文件数据
        Dim fileData(fileLength - 1) As Byte
        Array.Copy(imageData, lengthOffset + 8, fileData, 0, fileLength)

        Return fileData
    End Function

    ' 在字节数组中查找子序列的辅助函数
    Private Function IndexOfSequence(source As Byte(), sequence As Byte()) As Integer
        For i As Integer = 0 To source.Length - sequence.Length
            Dim match As Boolean = True
            For j As Integer = 0 To sequence.Length - 1
                If source(i + j) <> sequence(j) Then
                    match = False
                    Exit For
                End If
            Next
            If match Then Return i
        Next
        Return -1
    End Function

    ' 格式化文件大小显示
    Private Function FormatFileSize(bytes As Long) As String
        Dim sizes() As String = {"B", "KB", "MB", "GB"}
        Dim order As Integer = 0
        Dim size As Double = bytes

        While size >= 1024 AndAlso order < sizes.Length - 1
            order += 1
            size = size / 1024
        End While

        Return $"{size:0.##} {sizes(order)}"
    End Function
End Class