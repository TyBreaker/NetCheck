Imports System.Windows.Forms.DataVisualization.Charting

Public Class Form1

    Const MAXGRAPHS As Integer = 5 ' live, day, week, month and year graphs
    Private rand As New Random ' cause a random delay when spawning multiple threads
    Private dataGatherers As New List(Of PingBGW) ' list of ping threads for sites
    Public pingInterval As Long = 5 ' ping frequency in seconds
    Public pingDatabase As String = Application.UserAppDataPath ' location of ping data files
    Private pingTables(MAXGRAPHS) As List(Of DataTable) ' all the data
    Private graphTimer As New Timer ' auto-update graphs
    Private scanning As Boolean = False ' track if scanning mode is on
    Private Enum Extent
        LIVE = 1 ' last hour graphed (in hours)
        DAY = 24 ' last day graphed (in hours)
        WEEK = 168 ' last week graphed (in hours)
        MONTH = 720 ' last month graphed (in hours)
        YEAR = 8760 ' last year graphed (in hours)
    End Enum
    Private Enum Graphs
        LIVE = 0 ' ID for last hour graph
        DAY = 1 ' ID for last day graph
        WEEK = 2 ' ID for last week graph
        MONTH = 3 ' ID for last month graph
        YEAR = 4 ' ID for last year graph
    End Enum
    Private locker As New Object ' for locking critical sections of code
    Private told As Boolean = False ' Ensure balloon tip only appears once

    Private Function getAvg(ByVal chartView As DataView) As Double
        ' work out the average value for the average series so as to suitably scale the Y axis
        Dim seriesTotal As Double = 0
        Dim numPoints As Double = 0
        For Each dataPoint As DataRowView In chartView
            seriesTotal += dataPoint.Row.Item("pingResponse")
            numPoints += 1
        Next
        If numPoints = 0 Then
            Return 0
        Else
            Return seriesTotal / numPoints
        End If
    End Function

    Private Sub DeactivateGraph(ByVal sender As Object, ByVal e As EventArgs)
        ' MouseLeave event for each graph
        Dim whichChart As Chart = CType(sender, Chart)
        whichChart.BorderlineWidth = 0
        whichChart.ChartAreas("Default").AxisX.TitleForeColor = Color.Black
        whichChart.ChartAreas("Default").AxisY.TitleForeColor = Color.Black
        whichChart.ChartAreas("Default").CursorX.Position = Double.NaN
        whichChart.ChartAreas("Default").CursorY.Position = Double.NaN
        For Each legend As Legend In whichChart.Legends
            legend.ForeColor = Color.Black
        Next
        Me.TextBox1.Select()
    End Sub

    Private Sub ActivateGraph(ByVal sender As Object, ByVal e As EventArgs)
        ' MouseEnter event for each graph
        Dim whichChart As Chart = CType(sender, Chart)
        whichChart.Select()
        whichChart.BorderlineColor = Color.Red
        whichChart.BorderlineWidth = 1
        whichChart.BorderlineDashStyle = ChartDashStyle.Solid
        whichChart.ChartAreas("Default").AxisX.TitleForeColor = Color.Red
        whichChart.ChartAreas("Default").AxisY.TitleForeColor = Color.Red
        For Each legend As Legend In whichChart.Legends
            legend.ForeColor = Color.Red
        Next
    End Sub

    Private Sub highlightGraph(ByVal sender As Object, ByVal e As MouseEventArgs)
        ' MouseMove event for each graph
        Dim whichChart As Chart = CType(sender, Chart)
        SyncLock locker
            whichChart.DataBind()
            If whichChart.DataSource.Count > 0 Then
                Dim mousePoint As Point = New Point(e.X, e.Y)
                whichChart.ChartAreas("Default").CursorX.SetCursorPixelPosition(mousePoint, False)
                whichChart.ChartAreas("Default").CursorY.SetCursorPixelPosition(mousePoint, False)
            End If
        End SyncLock
    End Sub

    Private Sub ZoomGraph(ByVal sender As Object, ByVal e As KeyEventArgs)
        ' KeyUp event for each graph
        Dim whichChart As Chart = CType(sender, Chart)
        Select Case e.KeyCode
            Case Keys.Escape ' return to previous level of zoom
                whichChart.ChartAreas("Default").AxisX.ScaleView.ZoomReset()
                whichChart.ChartAreas("Default").CursorX.SelectionStart = Double.NaN
                whichChart.ChartAreas("Default").CursorX.SelectionEnd = Double.NaN
                whichChart.ChartAreas("Default").AxisY.ScaleView.ZoomReset()
                whichChart.ChartAreas("Default").CursorY.SelectionStart = Double.NaN
                whichChart.ChartAreas("Default").CursorY.SelectionEnd = Double.NaN
            Case Keys.Oemplus, Keys.Add, Keys.PageUp ' zoom in
                whichChart.ChartAreas("Default").AxisY.ScaleView.Zoom(0, whichChart.ChartAreas("Default").AxisY.ScaleView.ViewMaximum / 2)
            Case Keys.OemMinus, Keys.Subtract, Keys.PageDown ' zoom out
                whichChart.ChartAreas("Default").AxisY.ScaleView.Zoom(0, whichChart.ChartAreas("Default").AxisY.ScaleView.ViewMaximum * 2)
        End Select
    End Sub

    Private Sub allowZoom(ByVal graphArea As ChartArea)
        ' allow interactive zoom
        graphArea.AxisX.ScaleView.Zoomable = True
        graphArea.AxisY.ScaleView.Zoomable = True
        graphArea.CursorX.IsUserEnabled = True
        graphArea.CursorX.IsUserSelectionEnabled = True
        graphArea.CursorX.Interval = 0
        graphArea.CursorX.LineColor = Color.Red
        graphArea.CursorX.LineWidth = 1
        graphArea.CursorX.LineDashStyle = ChartDashStyle.Dot
        graphArea.CursorX.Interval = 0
        graphArea.CursorX.SelectionColor = Color.Yellow
        graphArea.CursorY.IsUserEnabled = True
        graphArea.CursorY.IsUserSelectionEnabled = True
        graphArea.CursorY.Interval = 0
        graphArea.CursorY.LineColor = Color.Red
        graphArea.CursorY.LineWidth = 1
        graphArea.CursorY.LineDashStyle = ChartDashStyle.Dot
        graphArea.CursorY.Interval = 0
        graphArea.CursorY.SelectionColor = Color.Yellow
    End Sub

    Private Sub ShowGraph(ByVal siteName As String, ByVal graphName As String, ByVal graphData As List(Of DataTable), ByVal graphExtent As Double, ByVal graphType As Graphs, ByVal avgSeriesTitle As String, ByVal avgSeriesTooltip As String, ByVal rngSeriesTooltip As String, ByVal rngSeriesTitle As String, ByVal xAxisLabel As String, ByVal xAxisIntervalType As DateTimeIntervalType, ByVal xAxisInterval As Double, ByVal timeLimit As Long, ByVal xAxisTitlePrefix As String, ByVal xAxisTitleFormat As String)
        ' render the five graphs
        Dim graph As New Chart
        graph.BackColor = Color.White

        Dim graphLegend As Legend = New Legend()
        graphLegend.BackColor = Color.White

        ' Setup the average series:
        Dim avgSeries As Series = New Series()
        avgSeries.ChartType = SeriesChartType.Line
        avgSeries.Name = avgSeriesTitle
        avgSeries.XValueMember = “pingTime”
        avgSeries.XValueType = ChartValueType.DateTime
        avgSeries.YValueMembers = “pingResponse”
        avgSeries.ToolTip = avgSeriesTooltip
        avgSeries.IsVisibleInLegend = True
        If graphType = Graphs.LIVE Then
            avgSeries.Color = Color.Blue
        Else
            avgSeries.Color = Color.Magenta
        End If

        ' Setup the range series:
        Dim rngSeries As Series = New Series()
        If graphType <> Graphs.LIVE Then
            rngSeries.ChartType = SeriesChartType.Range
            rngSeries.Name = rngSeriesTitle
            rngSeries.XValueMember = “pingTime”
            rngSeries.XValueType = ChartValueType.DateTime
            rngSeries.YValueMembers = “minLatency, maxLatency”
            rngSeries.ToolTip = rngSeriesTooltip
            rngSeries.BackGradientStyle = GradientStyle.TopBottom
            rngSeries.Color = Color.FromArgb(128, Color.DarkBlue)
            rngSeries.BackSecondaryColor = Color.FromArgb(128, Color.LightBlue)
            rngSeries.IsVisibleInLegend = True
        End If

        ' Format the graph:
        Dim graphArea As New ChartArea
        Dim backImage As New NamedImage("Background", My.Resources.waiting)
        graph.Images.Add(backImage)
        graphArea.Name = "Default"
        graphArea.BackImageWrapMode = ChartImageWrapMode.Unscaled
        graphArea.BackImageAlignment = ChartImageAlignmentStyle.Left
        graphArea.AxisX.LabelStyle.Format = xAxisLabel
        graphArea.AxisX.IntervalType = xAxisIntervalType
        graphArea.AxisX.Interval = xAxisInterval
        graphArea.AxisX.MajorGrid.LineColor = Color.LightGray
        graphArea.AxisY.LabelStyle.Format = "D"
        graphArea.AxisY.Title = "Reply" & vbCrLf & "(ms)"
        graphArea.AxisY.IntervalAutoMode = IntervalAutoMode.VariableCount
        graphArea.AxisY.ScaleBreakStyle.Enabled = True
        graphArea.AxisY.ScaleBreakStyle.Spacing = 4
        graphArea.AxisY.ScaleBreakStyle.BreakLineStyle = BreakLineStyle.Straight
        graphArea.AxisY.ScaleBreakStyle.StartFromZero = StartFromZero.Auto
        graphArea.AxisY.MajorGrid.LineColor = Color.LightGray
        graphArea.AxisY.ScrollBar.BackColor = Color.White
        graphArea.AxisY.ScrollBar.ButtonColor = Color.LightGray

        ' select data to graph:
        Dim now As Date = Date.Now
        Dim chartView As DataView = New DataView
        SyncLock locker
            chartView.Table = graphData.Find(Function(x) x.TableName.Equals(siteName))
            chartView.Sort = "pingTime ASC"
            chartView.RowFilter = String.Format(Globalization.CultureInfo.InvariantCulture.DateTimeFormat, "pingTime > #{0}#", now.AddHours(-graphExtent))
            If chartView.Count > 0 Then
                If CDate(chartView(chartView.Count - 1)("pingTime")).Ticks - CDate(chartView(0)("pingTime")).Ticks < timeLimit Then
                    graphArea.AxisX.Minimum = CDate(chartView(0)("pingTime")).ToOADate
                    graphArea.AxisX.Maximum = CDate(chartView(0)("pingTime")).AddHours(graphExtent).ToOADate
                    graphArea.AxisX.Title = xAxisTitlePrefix & CDate(chartView(0)("pingTime")).ToString(xAxisTitleFormat) & " - " & CDate(chartView(chartView.Count - 1)("pingTime")).ToString(xAxisTitleFormat) & ")"
                Else
                    graphArea.AxisX.Minimum = CDate(chartView(chartView.Count - 1)("pingTime")).AddHours(-graphExtent).ToOADate
                    graphArea.AxisX.Maximum = CDate(chartView(chartView.Count - 1)("pingTime")).ToOADate
                    graphArea.AxisX.Title = xAxisTitlePrefix & CDate(chartView(0)("pingTime")).ToString(xAxisTitleFormat) & " - " & CDate(chartView(chartView.Count - 1)("pingTime")).ToString(xAxisTitleFormat) & ")"
                End If
                graphArea.BackColor = Color.Beige
                Dim avgY As Double = getAvg(chartView)
                If avgY > 0 Then
                    graphArea.AxisY.ScaleView.Zoom(0, getAvg(chartView) * 3) ' the starting zoom level
                End If
                allowZoom(graphArea)
            Else
                graphArea.BackColor = Color.White
                graphArea.AxisX.Minimum = now.ToOADate
                graphArea.AxisX.Maximum = now.AddHours(graphExtent).ToOADate
                graphArea.AxisX.Title = xAxisTitlePrefix & "none"
                graphArea.BackImage = "Background"
                graphArea.CursorX.IsUserSelectionEnabled = False
                graphArea.CursorY.IsUserSelectionEnabled = False
            End If
            graph.DataSource = chartView
        End SyncLock
        ' Bring it all together:
        graph.ChartAreas.Add(graphArea)
        graph.Legends.Add(graphLegend)
        graph.Name = graphName
        If graphType <> Graphs.LIVE Then
            graph.Series.Add(rngSeries)
        End If
        graph.Series.Add(avgSeries)
        graph.Size = New Drawing.Size(Me.FlowLayoutPanel2.ClientSize.Width - Me.FlowLayoutPanel2.Margin.Left - Me.FlowLayoutPanel2.Margin.Right,
                                      Me.FlowLayoutPanel2.ClientSize.Height / MAXGRAPHS - Me.FlowLayoutPanel2.Margin.Top - Me.FlowLayoutPanel2.Margin.Bottom - 1)

        ' add the control to the panel:
        AddHandler graph.KeyUp, AddressOf ZoomGraph ' so we can zoom the graph
        AddHandler graph.MouseLeave, AddressOf DeactivateGraph ' so we can remove the highlight from mouse over
        AddHandler graph.MouseEnter, AddressOf ActivateGraph ' so we can highlight the graph on mouse over 
        AddHandler graph.MouseMove, AddressOf highlightGraph ' allow selection of chart
        Me.FlowLayoutPanel2.Controls.Add(graph)
    End Sub

    Private Sub ShowGraphs()
        ' display the ping graphs for the site
        If Me.ListBox1.SelectedIndex >= 0 Then
            Me.FlowLayoutPanel2.Controls.Clear()
            ShowGraph(Me.ListBox1.SelectedItem.ToString, "Live", pingTables(Graphs.LIVE), Extent.LIVE, Graphs.LIVE, "each ping", "#VAL ms\n#VALX{h:mm:ss}", "", "", "h:mm", DateTimeIntervalType.Auto, 0, TimeSpan.TicksPerHour, "last 60 mins (available data: ", "h:mm")
            ShowGraph(Me.ListBox1.SelectedItem.ToString, "Day", pingTables(Graphs.DAY), Extent.DAY, Graphs.DAY, "average over 5 mins", "#VAL ms\n#VALX{ddd h:mmtt}", "Min: #VALY1 ms\nMax: #VALY2 ms\n#VALX{ddd h:mmtt}", "range over 5 mins", "H ", DateTimeIntervalType.Hours, 2, TimeSpan.TicksPerDay, "last 24 hrs (available data: ", "H:mm")
            ShowGraph(Me.ListBox1.SelectedItem.ToString, "Week", pingTables(Graphs.WEEK), Extent.WEEK, Graphs.WEEK, "average over 30 mins", "#VAL ms\n#VALX{ddd h:mmtt}", "Min: #VALY1 ms\nMax: #VALY2 ms\n#VALX{ddd h:mmtt}", "range over 30 mins", "ddd", DateTimeIntervalType.Days, 1, TimeSpan.TicksPerDay * 7, "last 7 days (available data: ", "H:mm ddd")
            ShowGraph(Me.ListBox1.SelectedItem.ToString, "Month", pingTables(Graphs.MONTH), Extent.MONTH, Graphs.MONTH, "average over 2 hrs", "#VAL ms\n#VALX{ddd h:mmtt}", "Min: #VALY1 ms\nMax: #VALY2 ms\n#VALX{ddd h:mmtt}", "range over 2 hrs", "d ", DateTimeIntervalType.Days, 2, TimeSpan.TicksPerDay * 30, "last 30 days (available data: ", "htt d MMM")
            ShowGraph(Me.ListBox1.SelectedItem.ToString, "Year", pingTables(Graphs.YEAR), Extent.YEAR, Graphs.YEAR, "average over 1 day", "#VAL ms\n#VALX{ddd d MMM}", "Min: #VALY1 ms\nMax: #VALY2 ms\n#VALX{ddd d MMM}", "range over 1 day", "MMM", DateTimeIntervalType.Months, 1, TimeSpan.TicksPerDay * 365, "last 12 months (available data: ", "d MMM")
        End If
    End Sub

    Private Sub loadData(ByVal siteName As String, ByRef pingDataTable As DataTable, ByVal suffix As String)
        ' load data from database folder:
        Dim pingDataFileRow As String() ' to capture each row of the CSV
        Dim pingDataFile As FileIO.TextFieldParser
        If IO.File.Exists(pingDatabase & "\" & siteName & suffix) Then
            pingDataFile = New FileIO.TextFieldParser(pingDatabase & "\" & siteName & suffix)
            pingDataFile.TextFieldType = FileIO.FieldType.Delimited
            pingDataFile.SetDelimiters(",")
            While Not pingDataFile.EndOfData
                Try
                    pingDataFileRow = pingDataFile.ReadFields()
                    SyncLock locker
                        pingDataTable.Rows.Add(CDate(pingDataFileRow(0)), CLng(pingDataFileRow(1)), CLng(pingDataFileRow(2)), CLng(pingDataFileRow(3)))
                    End SyncLock
                Catch
                End Try
            End While
            pingDataFile.Close()
        End If
    End Sub

    Private Sub trimData(ByVal siteName As String, ByVal suffix As String, ByVal timeLimit As Long)
        ' purge data older than that required for this graph type:
        Dim now As Date = Date.Now
        'Dim keepData As Boolean
        Dim pingDataFileRow As String() ' to capture each row of the CSV
        Dim pingInputDataFile As FileIO.TextFieldParser
        If IO.File.Exists(pingDatabase & "\" & siteName & suffix) Then
            pingInputDataFile = New FileIO.TextFieldParser(pingDatabase & "\" & siteName & suffix)
            pingInputDataFile.TextFieldType = FileIO.FieldType.Delimited
            pingInputDataFile.SetDelimiters(",")
            While Not pingInputDataFile.EndOfData
                Try
                    pingDataFileRow = pingInputDataFile.ReadFields()
                    If (now - CDate(pingDataFileRow(0))).Ticks < timeLimit Then
                        My.Computer.FileSystem.WriteAllText(pingDatabase & "\" & siteName & suffix & "_tmp", pingDataFileRow(0) & "," & pingDataFileRow(1) & "," & pingDataFileRow(2) & "," & pingDataFileRow(3) & vbCrLf, True)
                    End If
                Catch
                End Try
            End While
            pingInputDataFile.Close()
            ' make temp file final:
            If IO.File.Exists(pingDatabase & "\" & siteName & suffix) Then
                IO.File.Delete(pingDatabase & "\" & siteName & suffix)
            End If
            If IO.File.Exists(pingDatabase & "\" & siteName & suffix & "_tmp") Then
                IO.File.Move(pingDatabase & "\" & siteName & suffix & "_tmp", pingDatabase & "\" & siteName & suffix)
            End If
        End If
    End Sub

    Private Function LoadGraphData(ByVal siteName As String) As String
        ' Read in all the data stored on the file system.
        Dim hostname As String
        Try
            hostname = Net.Dns.GetHostEntry(siteName).HostName.Replace(".local", "") ' try to resolve the IP address to a name
        Catch ex As Exception
            hostname = siteName ' fine, stick with the IP address then
        End Try
        If Not IO.File.Exists(pingDatabase & "\" & siteName & ".lock") Then ' check if a thread is already monitoring this site
            ' Purge older data from the database files
            trimData(hostname, "_live.csv", TimeSpan.TicksPerHour)
            trimData(hostname, "_daily.csv", TimeSpan.TicksPerDay)
            trimData(hostname, "_weekly.csv", TimeSpan.TicksPerDay * 7)
            trimData(hostname, "_monthly.csv", TimeSpan.TicksPerDay * 30)
            trimData(hostname, "_yearly.csv", TimeSpan.TicksPerDay * 365)

            ' Read the ping data for the site into memory
            loadData(hostname, pingTables(Graphs.LIVE).Find(Function(x) x.TableName.Equals(hostname)), "_live.csv")
            loadData(hostname, pingTables(Graphs.DAY).Find(Function(x) x.TableName.Equals(hostname)), "_daily.csv")
            loadData(hostname, pingTables(Graphs.WEEK).Find(Function(x) x.TableName.Equals(hostname)), "_weekly.csv")
            loadData(hostname, pingTables(Graphs.MONTH).Find(Function(x) x.TableName.Equals(hostname)), "_monthly.csv")
            loadData(hostname, pingTables(Graphs.YEAR).Find(Function(x) x.TableName.Equals(hostname)), "_yearly.csv")
        End If
        Return hostname
    End Function

    Private Sub gatherData(ByVal siteName As String, ByVal pingDataTables As List(Of DataTable), ByVal pingTime As Date, ByVal pingResult As Long, ByRef lastInterval As Integer, ByVal thisInterval As Integer, ByRef pingIntervalTotal As Long, ByRef pingIntervalCount As Long, ByRef pingIntervalMin As Long, ByRef pingIntervalMax As Long, ByVal intervalDuration As Long, ByVal dataFileSuffix As String)
        ' record the new ping data
        Dim dt As DataTable = pingDataTables.Find(Function(x) x.TableName.Equals(siteName)) ' which graph data of which site are we recording?
        If dataFileSuffix = "_live.csv" Then
            ' for the live graph, record every ping
            dt.Rows.Add(pingTime, pingResult, 0, 0)
            If pingTime.Ticks - CDate(dt.Rows(0)("pingTime")).Ticks > intervalDuration Then ' only save last hour's data
                dt.Rows(0).Delete() ' purge the oldest row in the table
            End If
            My.Computer.FileSystem.WriteAllText(pingDatabase & "\" & siteName & dataFileSuffix, pingTime.ToString & "," & pingResult.ToString & ",0,0" & vbCrLf, True)
        Else
            ' for the day/week/month/year graphs, only record the average for the applicable interval
            If lastInterval = -1 Or thisInterval = lastInterval Then
                ' within the same interval so accumulate to calculate average
                pingIntervalTotal += pingResult
                pingIntervalCount += 1
                pingIntervalMin = IIf(pingResult < pingIntervalMin Or pingIntervalMin = 0, pingResult, pingIntervalMin)
                pingIntervalMax = IIf(pingResult > pingIntervalMax Or pingIntervalMax = 0, pingResult, pingIntervalMax)
            Else
                ' entered a new interval so save the previous interval's ping average:
                dt.Rows.Add(pingTime, pingIntervalTotal \ pingIntervalCount, pingIntervalMin, pingIntervalMax) ' add ping data to in-memory database
                If pingTime.Ticks - CDate(dt.Rows(0)("pingTime")).Ticks > intervalDuration Then
                    dt.Rows(0).Delete()
                End If
                My.Computer.FileSystem.WriteAllText(pingDatabase & "\" & siteName & dataFileSuffix, pingTime.ToString & "," & (pingIntervalTotal \ pingIntervalCount).ToString & "," & pingIntervalMin.ToString & "," & pingIntervalMax.ToString & vbCrLf, True)
                pingIntervalTotal = pingResult
                pingIntervalCount = 1
                pingIntervalMin = pingResult
                pingIntervalMax = pingResult
            End If
            lastInterval = thisInterval
        End If
    End Sub

    Private Sub dataLoaded(sender As Object, e As ComponentModel.ProgressChangedEventArgs)
        ' Redisplay the graphs once all data has been loaded from files
        Dim names() As String = CStr(e.UserState).Split(":") ' IP address:hostname pair
        If names(0) = "remove" Then
            RemoveSite(Me.ListBox1.FindStringExact(names(1)))
        ElseIf Not IO.File.Exists(pingDatabase & "\" & names(1) & ".lock") Then ' check if a thread is already monitoring this site
            Me.ListBox1.Items.Item(Me.ListBox1.FindStringExact(names(0))) = names(1) ' rename the listbox entry
            pingTables(Graphs.LIVE).Find(Function(x) x.TableName.Equals(names(0))).TableName = names(1) ' rename the data table
            pingTables(Graphs.DAY).Find(Function(x) x.TableName.Equals(names(0))).TableName = names(1) ' rename the data table
            pingTables(Graphs.WEEK).Find(Function(x) x.TableName.Equals(names(0))).TableName = names(1) ' rename the data table
            pingTables(Graphs.MONTH).Find(Function(x) x.TableName.Equals(names(0))).TableName = names(1) ' rename the data table
            pingTables(Graphs.YEAR).Find(Function(x) x.TableName.Equals(names(0))).TableName = names(1) ' rename the data table
            dataGatherers.Find(Function(x) x.pingID.Equals(names(0))).pingID = names(1) 'rename the thread ID
            graphUpdate() ' redraw the graph
        End If
    End Sub

    Private Sub PingSite(ByVal sender As Object, ByVal e As ComponentModel.DoWorkEventArgs)
        ' worker task for threads: pings the sites, records the data
        Dim dataGatherer As PingBGW = CType(sender, PingBGW) ' this thread
        Dim doPing As New Net.NetworkInformation.Ping
        Dim pingResult As Net.NetworkInformation.PingReply ' how long did the ping take?
        Dim pingTime As Date ' when was the ping sent?
        Dim pingIntervalTotal() As Long = {0, 0, 0, 0, 0} ' accumulate the ping response times for the interval
        Dim pingIntervalCount() As Long = {0, 0, 0, 0, 0} ' note how pany pings we've sent in the interval
        Dim lastInterval() As Integer = {-1, -1, -1, -1, -1} ' initial iteration of the interval
        Dim pingIntervalMin() As Long = {0, 0, 0, 0, 0} ' remember the quickest response in the interval
        Dim pingIntervalMax() As Long = {0, 0, 0, 0, 0} ' remember the slowest response in the interval
        Dim multipleSites As Boolean ' used to stagger the start of each thread to avoid flooding

        Dim siteName As String = LoadGraphData(e.Argument) ' first, read in site data from database files
        dataGatherer.ReportProgress(0, e.Argument & ":" & siteName) ' redraw the graphs

        If IO.File.Exists(pingDatabase & "\" & siteName & ".lock") Then ' check if a thread is already monitoring this site
            dataGatherer.ReportProgress(0, "remove:" & e.Argument) ' tell the parent process to remove the duplicate site from the list
        Else
            IO.File.Create(pingDatabase & "\" & siteName & ".lock") ' this site is now being monitored
            SyncLock locker
                multipleSites = pingTables(Graphs.LIVE).Count > 1
            End SyncLock
            If multipleSites Then ' if there are multiple sites being monitored
                Threading.Thread.Sleep(rand.Next(0, 59) * 1000) ' randomise the starting time for the thread to avoid flooding
            End If
            Do Until dataGatherer.CancellationPending
                pingTime = Date.Now ' time of the ping
                Try
                    pingResult = doPing.Send(siteName) ' the ping
                    SyncLock locker
                        If pingResult.Status = Net.NetworkInformation.IPStatus.Success Then
                            gatherData(siteName, pingTables(Graphs.LIVE), pingTime, pingResult.RoundtripTime, lastInterval(Graphs.LIVE), 0, pingIntervalTotal(Graphs.LIVE), pingIntervalCount(Graphs.LIVE), pingIntervalMin(Graphs.LIVE), pingIntervalMax(Graphs.LIVE), TimeSpan.TicksPerHour, "_live.csv")
                            gatherData(siteName, pingTables(Graphs.DAY), pingTime, pingResult.RoundtripTime, lastInterval(Graphs.DAY), pingTime.Minute \ 5, pingIntervalTotal(Graphs.DAY), pingIntervalCount(Graphs.DAY), pingIntervalMin(Graphs.DAY), pingIntervalMax(Graphs.DAY), TimeSpan.TicksPerDay, "_daily.csv")
                            gatherData(siteName, pingTables(Graphs.WEEK), pingTime, pingResult.RoundtripTime, lastInterval(Graphs.WEEK), pingTime.Minute \ 30, pingIntervalTotal(Graphs.WEEK), pingIntervalCount(Graphs.WEEK), pingIntervalMin(Graphs.WEEK), pingIntervalMax(Graphs.WEEK), TimeSpan.TicksPerDay * 7, "_weekly.csv")
                            gatherData(siteName, pingTables(Graphs.MONTH), pingTime, pingResult.RoundtripTime, lastInterval(Graphs.MONTH), pingTime.Hour \ 2, pingIntervalTotal(Graphs.MONTH), pingIntervalCount(Graphs.MONTH), pingIntervalMin(Graphs.MONTH), pingIntervalMax(Graphs.MONTH), TimeSpan.TicksPerDay * 30, "_monthly.csv")
                            gatherData(siteName, pingTables(Graphs.YEAR), pingTime, pingResult.RoundtripTime, lastInterval(Graphs.YEAR), pingTime.Day, pingIntervalTotal(Graphs.YEAR), pingIntervalCount(Graphs.YEAR), pingIntervalMin(Graphs.YEAR), pingIntervalMax(Graphs.YEAR), TimeSpan.TicksPerDay * 365, "_yearly.csv")
                        Else
                            gatherData(siteName, pingTables(Graphs.LIVE), pingTime, 0, lastInterval(Graphs.LIVE), 0, pingIntervalTotal(Graphs.LIVE), pingIntervalCount(Graphs.LIVE), pingIntervalMin(Graphs.LIVE), pingIntervalMax(Graphs.LIVE), TimeSpan.TicksPerHour, "_live.csv")
                            gatherData(siteName, pingTables(Graphs.DAY), pingTime, 0, lastInterval(Graphs.DAY), pingTime.Minute \ 5, pingIntervalTotal(Graphs.DAY), pingIntervalCount(Graphs.DAY), pingIntervalMin(Graphs.DAY), pingIntervalMax(Graphs.DAY), TimeSpan.TicksPerDay, "_daily.csv")
                            gatherData(siteName, pingTables(Graphs.WEEK), pingTime, 0, lastInterval(Graphs.WEEK), pingTime.Minute \ 30, pingIntervalTotal(Graphs.WEEK), pingIntervalCount(Graphs.WEEK), pingIntervalMin(Graphs.WEEK), pingIntervalMax(Graphs.WEEK), TimeSpan.TicksPerDay * 7, "_weekly.csv")
                            gatherData(siteName, pingTables(Graphs.MONTH), pingTime, 0, lastInterval(Graphs.MONTH), pingTime.Hour \ 2, pingIntervalTotal(Graphs.MONTH), pingIntervalCount(Graphs.MONTH), pingIntervalMin(Graphs.MONTH), pingIntervalMax(Graphs.MONTH), TimeSpan.TicksPerDay * 30, "_monthly.csv")
                            gatherData(siteName, pingTables(Graphs.YEAR), pingTime, 0, lastInterval(Graphs.YEAR), pingTime.Day, pingIntervalTotal(Graphs.YEAR), pingIntervalCount(Graphs.YEAR), pingIntervalMin(Graphs.YEAR), pingIntervalMax(Graphs.YEAR), TimeSpan.TicksPerDay * 365, "_yearly.csv")
                        End If
                    End SyncLock
                Catch ex As Exception
                End Try
                Threading.Thread.Sleep(pingInterval * 1000)
            Loop
        End If
    End Sub

    Private Sub updateStatus(ByVal msg As String, ByVal colour As Color)
        ' for placing messages in the status bar at the bottom of the window
        Me.ToolStripStatusLabel2.ForeColor = colour
        Me.ToolStripStatusLabel2.Text = msg
        If Not scanning Then
            Me.ToolStripStatusLabel1.Text = ""
        End If
    End Sub

    Private Function createDataTable(ByVal siteName As String) As DataTable
        Dim newDataTable As New DataTable ' establish a new in-memory database for the ping data
        newDataTable.TableName = siteName
        newDataTable.Columns.Add("pingTime", GetType(Date))
        newDataTable.Columns.Add("pingResponse", GetType(Long))
        newDataTable.Columns.Add("minLatency", GetType(Long))
        newDataTable.Columns.Add("maxLatency", GetType(Long))
        Return newDataTable
    End Function

    Private Sub AddSite(ByVal userSite As String, ByVal silence As Boolean)
        ' create a thread to monitor the new site
        Dim siteName As New String(userSite.Where(Function(x) Not Char.IsWhiteSpace(x)).ToArray()) ' remove any white space
        If siteName <> "" Then
            If Me.ListBox1.FindStringExact(siteName) = -1 Then ' only add the site if it's not already in the list
                Me.ListBox1.Items.Add(siteName)
                If Not silence Then
                    updateStatus("Added " & siteName, Color.Black)
                End If

                ' create the database for the site
                SyncLock locker
                    pingTables(Graphs.LIVE).Add(createDataTable(siteName))
                    pingTables(Graphs.DAY).Add(createDataTable(siteName))
                    pingTables(Graphs.WEEK).Add(createDataTable(siteName))
                    pingTables(Graphs.MONTH).Add(createDataTable(siteName))
                    pingTables(Graphs.YEAR).Add(createDataTable(siteName))
                End SyncLock

                ' create the data-gathering thread
                Dim dataGatherer As New PingBGW(siteName) ' Backgroundworker with name!
                AddHandler dataGatherer.DoWork, AddressOf PingSite
                AddHandler dataGatherer.ProgressChanged, AddressOf dataLoaded
                dataGatherer.WorkerReportsProgress = True
                dataGatherer.WorkerSupportsCancellation = True
                dataGatherer.RunWorkerAsync(siteName) ' starts the thread
                dataGatherers.Add(dataGatherer) ' add the new thread to our list
            ElseIf Not silence Then ' don't report warnings when silenced
                ' site already listed
                Media.SystemSounds.Hand.Play()
                updateStatus(siteName & " is already listed!", Color.Red)
            End If
        End If
    End Sub

    Private Sub RemoveSite(ByVal target As Integer)
        ' remove a site from being monitored
        If target = -1 Then ' remove the user-selected site
            Dim siteName As String = Me.ListBox1.SelectedItem
            updateStatus("Removed " & Me.ListBox1.SelectedItem, Color.Black)
            Me.ListBox1.Items.RemoveAt(ListBox1.SelectedIndex)
            If Me.ListBox1.Items.Count = 0 Then
                Me.RemoveToolStripMenuItem.Enabled = False
            End If
            graphTimer.Stop() ' stop updating the graphs
            Me.FlowLayoutPanel2.Controls.Clear() ' remove the graphs for the site
            dataGatherers.Find(Function(x) x.pingID.Equals(siteName)).CancelAsync() ' stop the thread for this site
            dataGatherers.Remove(dataGatherers.Find(Function(x) x.pingID.Equals(siteName))) ' remove the thread for this site
            ' Remove the lock file
            If IO.File.Exists(pingDatabase & "\" & siteName & ".lock") Then
                IO.File.Delete(pingDatabase & "\" & siteName & ".lock")
            End If
            ' Remove database files:
            If IO.File.Exists(pingDatabase & "\" & siteName & "_live.csv") Then
                IO.File.Delete(pingDatabase & "\" & siteName & "_live.csv")
            End If
            If IO.File.Exists(pingDatabase & "\" & siteName & "_daily.csv") Then
                IO.File.Delete(pingDatabase & "\" & siteName & "_daily.csv")
            End If
            If IO.File.Exists(pingDatabase & "\" & siteName & "_weekly.csv") Then
                IO.File.Delete(pingDatabase & "\" & siteName & "_weekly.csv")
            End If
            If IO.File.Exists(pingDatabase & "\" & siteName & "_monthly.csv") Then
                IO.File.Delete(pingDatabase & "\" & siteName & "_monthly.csv")
            End If
            If IO.File.Exists(pingDatabase & "\" & siteName & "_yearly.csv") Then
                IO.File.Delete(pingDatabase & "\" & siteName & "_yearly.csv")
            End If
        Else ' remote the duplicate site
            Me.ListBox1.Items.RemoveAt(target)
        End If
    End Sub

    Private Sub CloseToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CloseToolStripMenuItem.Click
        ' user selects File/Close
        Application.Exit()
    End Sub

    Private Sub SplitContainer1_SplitterMoved(sender As Object, e As SplitterEventArgs) Handles SplitContainer1.SplitterMoved
        ' reposition controls when splitter is moved
        AdjustControlsAfterResize()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        ' The Add button next to the Textbox is clicked: add the site to the list
        AddSite(Me.TextBox1.Text, False)
    End Sub

    Private Sub TextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox1.KeyDown
        ' allow user to press ENTER in text box to add site:
        If e.KeyCode = Keys.Enter Then
            AddSite(Me.TextBox1.Text, False)
        End If
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        ' user changed contents of text box
        If Me.TextBox1.Text = "" Then
            ' text box is empty:
            Me.Button1.Enabled = False
            Me.AddToolStripMenuItem.Enabled = False
        Else
            ' site has been entered:
            Me.Button1.Enabled = True
            Me.AddToolStripMenuItem.Enabled = True
        End If
    End Sub

    Private Sub AddToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AddToolStripMenuItem.Click
        ' The Add button in the toolbar is clicked: add the site to the list
        AddSite(Me.TextBox1.Text, False)
    End Sub

    Private Sub RemoveToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles RemoveToolStripMenuItem.Click
        ' The Remove button in the toolbar is clicked: remove the site from the list
        RemoveSite(-1)
    End Sub

    Private Sub updateGraph(ByVal graphName As String, ByVal timeLimit As Long, ByVal graphType As Graphs, ByVal graphExtent As Double, ByVal graphTitleStart As String, ByVal xAxisFormat As String)
        ' redraw each graph with the new data
        Dim whichChart As Chart = CType(Me.FlowLayoutPanel2.Controls.Item(graphName), Chart)
        If whichChart IsNot Nothing Then
            SyncLock locker
                If whichChart.DataSource.Count > 0 Then ' there is data to graph
                    If CDate(whichChart.DataSource(whichChart.DataSource.Count - 1)("pingTime")).Ticks - CDate(whichChart.DataSource(0)("pingTime")).Ticks < timeLimit Then
                        whichChart.ChartAreas("Default").AxisX.Minimum = CDate(whichChart.DataSource(0)("pingTime")).ToOADate
                        whichChart.ChartAreas("Default").AxisX.Maximum = CDate(whichChart.DataSource(0)("pingTime")).AddHours(graphExtent).ToOADate
                        whichChart.ChartAreas("Default").AxisX.Title = graphTitleStart & CDate(whichChart.DataSource(0)("pingTime")).ToString(xAxisFormat) & " - " & CDate(whichChart.DataSource(whichChart.DataSource.Count - 1)("pingTime")).ToString(xAxisFormat) & ")"
                    Else
                        whichChart.ChartAreas("Default").AxisX.Minimum = whichChart.DataSource(whichChart.DataSource.Count - 1)("pingTime").AddHours(-graphExtent).ToOADate
                        whichChart.ChartAreas("Default").AxisX.Maximum = CDate(whichChart.DataSource(whichChart.DataSource.Count - 1)("pingTime")).ToOADate
                        whichChart.ChartAreas("Default").AxisX.Title = graphTitleStart & CDate(whichChart.DataSource(0)("pingTime")).ToString(xAxisFormat) & " - " & CDate(whichChart.DataSource(whichChart.DataSource.Count - 1)("pingTime")).ToString(xAxisFormat) & ")"
                    End If
                    whichChart.ChartAreas("Default").BackColor = Color.Beige
                    whichChart.ChartAreas("Default").BackImage = "" ' remove the "waiting for data..." notice
                    allowZoom(whichChart.ChartAreas("Default"))
                Else
                    whichChart.ChartAreas("Default").BackColor = Color.White
                    whichChart.ChartAreas("Default").AxisX.Minimum = Now.ToOADate
                    whichChart.ChartAreas("Default").AxisX.Maximum = Now.AddHours(graphExtent).ToOADate
                    whichChart.ChartAreas("Default").AxisX.Title = graphTitleStart & "none"
                    whichChart.ChartAreas("Default").CursorX.IsUserSelectionEnabled = False
                    whichChart.ChartAreas("Default").CursorY.IsUserSelectionEnabled = False
                End If
                whichChart.DataBind() ' graph the new data
            End SyncLock
        End If
    End Sub

    Private Sub graphUpdate()
        ' display the latest data in the graphs
        If Me.ListBox1.SelectedIndex >= 0 Then
            updateGraph("Live", TimeSpan.TicksPerHour, Graphs.LIVE, Extent.LIVE, "last 60 mins (available data: ", "h:mm")
            updateGraph("Day", TimeSpan.TicksPerDay, Graphs.DAY, Extent.DAY, "last 24 hrs (available data: ", "H:mm")
            updateGraph("Week", TimeSpan.TicksPerDay * 7, Graphs.WEEK, Extent.WEEK, "last 7 days (available data: ", "H:mm ddd")
            updateGraph("Month", TimeSpan.TicksPerDay * 30, Graphs.MONTH, Extent.MONTH, "last 30 days (available data: ", "htt d MMM")
            updateGraph("Year", TimeSpan.TicksPerDay * 365, Graphs.YEAR, Extent.YEAR, "last 12 months (available data: ", "d MMM")
        End If
    End Sub

    Private Sub graphUpdateTimer(ByVal sender As Object, ByVal e As EventArgs)
        ' This is called from the timer to auto update the graphs:
        graphUpdate()
    End Sub

    Private Sub ListBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox1.SelectedIndexChanged
        ' user clicks on a site in the list: display the graphs
        If Me.ListBox1.SelectedIndex >= 0 Then
            graphTimer.Stop()
            updateStatus("", Color.Black)
            Me.RemoveToolStripMenuItem.Enabled = True
            If Not scanning Then
                Me.ToolStripStatusLabel1.Text = ""
            End If
            ShowGraphs()
            graphTimer.Start() ' graphs will auto update
        Else
            Me.RemoveToolStripMenuItem.Enabled = False
            graphTimer.Stop()
        End If
        Me.TextBox1.Select()
    End Sub

    Private Sub SettingsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SettingsToolStripMenuItem.Click
        ' Show settings pop-up:
        My.Forms.Form2.ShowDialog(Me)
    End Sub

    Private Sub AdjustControlsAfterResize()
        ' Cope with window being resized:
        Me.ListBox1.Width = Me.FlowLayoutPanel1.Width - Me.FlowLayoutPanel1.Margin.Left - Me.FlowLayoutPanel1.Margin.Right
        Me.ListBox1.Height = Me.FlowLayoutPanel1.Height - Me.TextBox1.Height - Me.FlowLayoutPanel1.Margin.Top - Me.FlowLayoutPanel1.Margin.Bottom - Me.TextBox1.Margin.Top
        Me.TextBox1.Width = Me.ListBox1.Width - Me.Button1.Width - Me.Label1.Width - Me.Button1.Margin.Left - Me.Button1.Margin.Right - Me.Label1.Margin.Left - Me.Label1.Margin.Right
    End Sub

    Private Sub PreloadSites()
        ' on startup, identify what sites need monitoring and create their threads
        pingTables(Graphs.LIVE) = New List(Of DataTable)
        pingTables(Graphs.DAY) = New List(Of DataTable)
        pingTables(Graphs.WEEK) = New List(Of DataTable)
        pingTables(Graphs.MONTH) = New List(Of DataTable)
        pingTables(Graphs.YEAR) = New List(Of DataTable)
        ' load all sites contained in database folder
        For Each siteDataFile As String In IO.Directory.GetFiles(pingDatabase, "*.csv")
            If IO.File.Exists(pingDatabase & "\" & IO.Path.GetFileNameWithoutExtension(siteDataFile).Split("_")(0) & ".lock") Then
                IO.File.Delete(pingDatabase & "\" & IO.Path.GetFileNameWithoutExtension(siteDataFile).Split("_")(0) & ".lock")
            End If
            AddSite(IO.Path.GetFileNameWithoutExtension(siteDataFile).Split("_")(0), True) ' create a thread for the site
        Next
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' here is where it all begins!
        Me.Icon = My.Resources.globe
        Me.NotifyIcon1.Icon = My.Resources.globe
        NotifyIcon1.Visible = True ' system tray icon
        Me.TextBox1.Select() ' sets default focus on textbox
        PreloadSites() ' list all sites in database folder
        If dataGatherers.Count = 0 Then
            Me.WindowState = FormWindowState.Normal ' present the window only if no sites have yet been configured
        End If
        graphTimer.Interval = pingInterval * 1000 ' redraw graphs after every ping
        AddHandler graphTimer.Tick, AddressOf graphUpdateTimer
    End Sub

    Private Sub FlowLayoutPanel1_Resize(sender As Object, e As EventArgs) Handles FlowLayoutPanel1.Resize
        ' reposition controls after window is resized
        AdjustControlsAfterResize()
    End Sub

    Private Sub FlowLayoutPanel2_Resize(sender As Object, e As EventArgs) Handles FlowLayoutPanel2.Resize
        ' redraw graphs after window is resized
        ShowGraphs()
    End Sub

    Private Sub NotifyIcon1_MouseDoubleClick(sender As Object, e As MouseEventArgs) Handles NotifyIcon1.MouseDoubleClick
        ' user double-clicked on system tray icon so open the window
        Me.Show()
        Me.WindowState = FormWindowState.Normal
        Me.Activate() ' bring window to top of desktop
    End Sub

    Private Sub AboutToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AboutToolStripMenuItem.Click
        ' show the About window
        My.Forms.AboutBox1.ShowDialog(Me)
    End Sub

    Private Sub UsersGuideToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles UsersGuideToolStripMenuItem.Click
        ' show the User Guide
        My.Forms.Form3.ShowDialog(Me)
    End Sub

    Private Sub bgw_ProgressChanged(ByVal sender As Object, ByVal e As ComponentModel.ProgressChangedEventArgs)
        ' create monitoring threads for the successful IP addresses
        Me.ToolStripProgressBar1.Value = e.ProgressPercentage
        Dim result() As String = CStr(e.UserState).Split(":")
        updateStatus(result(1), Color.Black)
        If result(0) = "yes" Then
            AddSite(result(1), False) ' add site silently
        End If
    End Sub

    Private Sub scanLAN(ByVal sender As Object, ByVal e As ComponentModel.DoWorkEventArgs)
        ' Discover LAN: automatically add all devices on the LAN
        Dim bgw As ComponentModel.BackgroundWorker = CType(sender, ComponentModel.BackgroundWorker)
        Dim doPing As New Net.NetworkInformation.Ping
        Dim pingResult As Net.NetworkInformation.PingReply
        Dim myIP As Net.IPAddress ' the first IPv4 address found for the localhost
        Dim ADDip As Net.IPAddress
        Dim hostname As Net.IPHostEntry
        Dim siteName As String

        ' get this computer's IP address:
        For Each ip As Net.IPAddress In Net.Dns.GetHostAddresses(Net.Dns.GetHostName())
            If ip.AddressFamily = Net.Sockets.AddressFamily.InterNetwork Then
                myIP = ip ' return the IPv4 address
                Exit For
            End If
        Next

        ' Now that we have the IP address, we're going to assume a subnet mask of 255.255.255.0:
        Dim addr_bytes() As Byte = myIP.GetAddressBytes()
        For i As Byte = 1 To 254 ' check all IP addresses in the subnet
            ADDip = New Net.IPAddress({addr_bytes(0), addr_bytes(1), addr_bytes(2), i})
            Try
                hostname = Net.Dns.GetHostEntry(ADDip)
                siteName = hostname.HostName ' we'll use the hostname instead
            Catch ex As Exception
                siteName = ADDip.ToString ' fine, we'll stick with the IP address
            End Try
            pingResult = doPing.Send(siteName) ' see if this address exists
            If pingResult.Status = Net.NetworkInformation.IPStatus.Success Then
                bgw.ReportProgress(i, "yes:" & siteName)
            Else
                bgw.ReportProgress(i, "no:" & siteName)
            End If
        Next
        e.Result = 0
    End Sub

    Private Sub bgw_RunWorkerCompleted(ByVal sender As Object, ByVal e As ComponentModel.RunWorkerCompletedEventArgs)
        ' Discover LAN: When the scan concludes, clean up:
        Me.ToolStripProgressBar1.Visible = False ' hide the progress bar
        Me.AddAllToolStripMenuItem.Enabled = True ' allow user to initiate another scan if they wish
        Me.ToolStripStatusLabel1.ForeColor = Color.Black
        Me.ToolStripStatusLabel1.Text = "Discover LAN completed!"
        scanning = False ' turn off the scanning notification
    End Sub

    Private Sub AddAllToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AddAllToolStripMenuItem.Click
        ' Discover LAN button - adds all available clients on this LAN
        Me.AddAllToolStripMenuItem.Enabled = False ' ensure user can't invoke it while it's already running

        ' inform user that a scan is in progress:
        Me.ToolStripStatusLabel1.ForeColor = Color.Black
        Me.ToolStripStatusLabel1.Text = "Discover LAN progress: "

        ' display the scan progress bar:
        Me.ToolStripProgressBar1.Maximum = 254
        Me.ToolStripProgressBar1.Minimum = 0
        Me.ToolStripProgressBar1.Value = 0
        Me.ToolStripProgressBar1.Visible = True

        scanning = True ' stop warnings appearing in the status bar as a result of the scan

        ' make the lengthy scan run in a background thread so the application can continue to be used:
        Dim bgw As New ComponentModel.BackgroundWorker
        AddHandler bgw.DoWork, AddressOf scanLAN
        AddHandler bgw.ProgressChanged, AddressOf bgw_ProgressChanged
        AddHandler bgw.RunWorkerCompleted, AddressOf bgw_RunWorkerCompleted
        bgw.WorkerReportsProgress = True
        bgw.RunWorkerAsync()
    End Sub

    Private Sub Form1_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        ' override closinng the window annd make it minimise instead.
        If e.CloseReason = CloseReason.UserClosing Then
            Me.WindowState = FormWindowState.Minimized
            If Not told Then ' only want the balloon tip to appear once
                Me.NotifyIcon1.BalloonTipIcon = ToolTipIcon.Info
                Me.NotifyIcon1.BalloonTipTitle = "I'm still running..."
                Me.NotifyIcon1.BalloonTipText = "NetCheck will continue to monitor your sites.  Choose File/Close to stop NetCheck."
                Me.NotifyIcon1.ShowBalloonTip(5000)
                told = True
            End If
            Me.Hide()
            e.Cancel = True
        End If
    End Sub

    Private Sub Form1_Resize(sender As Object, e As EventArgs) Handles Me.Resize
        ' when the window is resized
        If Me.WindowState = FormWindowState.Minimized Then
            Me.Hide()
        Else
            AdjustControlsAfterResize()
        End If
    End Sub

    Private Sub ToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem1.Click
        Application.Exit()
    End Sub

    Private Sub OpenToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles OpenToolStripMenuItem.Click
        Me.Show()
        Me.WindowState = FormWindowState.Normal
        Me.Activate() ' bring window to top of desktop
    End Sub

End Class
