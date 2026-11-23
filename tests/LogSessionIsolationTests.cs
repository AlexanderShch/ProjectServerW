// Runs the real C++/CLI session methods with the Excel worker paused before dequeue.
// No controller, Excel application, visible window, or two-hour wait is required.
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Net;
using System.Net.Sockets;
using System.Reflection;
using System.Threading;
using System.Windows.Forms;

static class LogSessionIsolationTests
{
    const BindingFlags Instance = BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic;
    const BindingFlags Static = BindingFlags.Static | BindingFlags.Public | BindingFlags.NonPublic;
    static Type formType, excelType;
    static DateTime epoch = new DateTime(2026, 11, 12, 12, 0, 0);
    static int passed;

    static object Get(object target, string name) { return target.GetType().GetField(name, Instance).GetValue(target); }
    static void Set(object target, string name, object value) { target.GetType().GetField(name, Instance).SetValue(target, value); }
    static object Call(object target, string name, params object[] args)
    {
        foreach (var method in target.GetType().GetMethods(Instance))
            if (method.Name == name && method.GetParameters().Length == args.Length)
                return method.Invoke(target, args);
        throw new MissingMethodException(name);
    }
    static object Queue { get { return excelType.GetField("excelExportQueue", Static).GetValue(null); } }
    static List<object> Jobs()
    {
        var result = new List<object>();
        foreach (var job in (IEnumerable)Queue) result.Add(job);
        return result;
    }
    static void ClearQueue()
    {
        object[] args = { null };
        while ((bool)Queue.GetType().GetMethod("TryDequeue").Invoke(Queue, args)) { }
    }
    static Form NewForm()
    {
        ClearQueue();
        var form = (Form)Activator.CreateInstance(formType);
        // Create a handle without showing a window, so real BeginInvoke callbacks work.
        var handle = form.Handle;
        Set(form, "settingsLoading", true);
        ((RadioButton)Get(form, "radioButtonSiemens")).Checked = false;
        Set(form, "settingsLoading", false);
        Set(form, "dataCollectionStartTime", epoch);
        Set(form, "commandConnection", new object());
        return form;
    }
    static DataTable Table(object form) { return (DataTable)Get(form, "dataTable"); }
    static void Row(object form, string label, bool pending)
    {
        var row = Table(form).NewRow();
        row["RealTime"] = label;
        if (pending) { Set(form, "pendingRow", row); Set(form, "pendingRowWrkBit", true); }
        else Table(form).Rows.Add(row);
    }
    static string Labels(DataTable table)
    {
        var labels = new List<string>();
        foreach (DataRow row in table.Rows) labels.Add((string)row["RealTime"]);
        return String.Join(",", labels.ToArray());
    }
    static void Equal(object expected, object actual)
    {
        if (!Object.Equals(expected, actual))
            throw new Exception("Expected " + expected + "; got " + actual);
    }
    static void Job(object job, string prefix, string labels)
    {
        Equal(prefix, Get(job, "filePrefix"));
        Equal(labels, Labels((DataTable)Get(job, "tableSnapshot")));
    }
    static object EndWork(object form)
    {
        Equal(true, Call(form, "QueueCompletedCycle", epoch.AddMinutes(10), -1, false));
        return Jobs()[Jobs().Count - 1];
    }
    static void Complete(object form, object job, bool saved)
    {
        if ((bool)Get(job, "completesCycle"))
            Call(form, "OnCompletedCycleSaved", Get(job, "sessionId"), saved);
        Call(form, "OnExcelExportCompleted", Get(job, "sessionId"), false, saved, false);
    }
    static void Test(string name, Action body)
    {
        body(); ++passed; Console.WriteLine("PASS " + name);
    }
    static ulong SocketHandle(TcpClient client)
    {
        return unchecked((ulong)client.Client.Handle.ToInt64());
    }

    static void EnqueueResponse(object form, byte[] response)
    {
        var socket = (ulong)Get(form, "clientSocket");
        Call(form, "EnqueueResponse", (response[0] & 0x80) != 0 ? response : Tagged(form, response), socket, Get(form, "commandConnection"));
    }

    static void DeferredResponse(object form, byte[] response) {
        var copy = (byte[])response.Clone();
        var peer = new Thread(delegate() {
            var deadline = DateTime.Now.AddSeconds(3);
            while (DateTime.Now < deadline) {
                if ((int)Get(form, "exchangeCode") == copy[1] && (int)Get(form, "exchangeType") == copy[0]) {
                    EnqueueResponse(form, copy); return;
                }
                Thread.Sleep(1);
            }
        }); peer.IsBackground = true; peer.Start();
    }
    static void Telemetry(object form, int time, bool working, bool shutdown, bool butStart, ushort limits = 0, bool ioActive = true)
    {
        var packet = new byte[66]; packet[3] = 60;
        Action<int, ushort> word = delegate(int offset, ushort value) {
            packet[4 + offset] = (byte)value; packet[5 + offset] = (byte)(value >> 8);
        };
        word(0, (ushort)time); packet[6] = 7;
        packet[4 + 9] = 4; packet[4 + 16] = (byte)(ioActive ? 1 : 0);
        word(29, (ushort)(limits | (butStart ? 0x4000 : 0)));
        word(43, (ushort)(working ? 0x4000 : 0)); packet[4 + 45] = (byte)(shutdown ? 1 : 0);
        Set(form, "firstDataReceived", true);
        Call(form, "AddDataToTableThreadSafe", packet, packet.Length, 8866, Get(form, "clientSocket"), Get(form, "commandConnection"));
    }
    static void ControlLog(object form, float marker)
    {
        var packet = new byte[95]; packet[0] = 1; packet[3] = 89;
        Array.Copy(BitConverter.GetBytes(marker), 0, packet, 4, 4);
        Call(form, "AppendControlLogToDataRow", packet, packet.Length);
    }
    static string Times(DataTable table)
    {
        var times = new List<string>();
        foreach (DataRow row in table.Rows) times.Add(row["Time"].ToString());
        return String.Join(",", times.ToArray());
    }
    static byte[] Response(byte type, byte code, byte[] payload)
    {
        var packet = new byte[payload.Length + 6];
        packet[0] = type; packet[1] = code; packet[3] = (byte)payload.Length;
        Array.Copy(payload, 0, packet, 4, payload.Length);
        ushort crc = 0xffff;
        for (int i = 0; i < packet.Length - 2; ++i) {
            crc ^= packet[i];
            for (int bit = 0; bit < 8; ++bit) crc = (ushort)((crc & 1) != 0 ? (crc >> 1) ^ 0xa001 : crc >> 1);
        }
        packet[packet.Length - 2] = (byte)crc; packet[packet.Length - 1] = (byte)(crc >> 8);
        return packet;
    }
    static byte[] Tagged(object form, byte[] plain, ulong? id = null) {
        var payload = new byte[plain[3] + 8];
        Array.Copy(BitConverter.GetBytes(id ?? (ulong)Get(form, "exchangeRequestId")), payload, 8);
        Array.Copy(plain, 4, payload, 8, plain[3]);
        return Response((byte)(plain[0] | 0x80), plain[1], payload);
    }
    static byte[] Group7(float[] values)
    {
        var payload = new byte[82]; payload[0] = 7;
        for (int i = 0; i < values.Length; ++i) Array.Copy(BitConverter.GetBytes(values[i]), 0, payload, 2 + 4 * i, 4);
        return Response(3, 7, payload);
    }
    static byte[] ReadCommand(NetworkStream stream)
    {
        var header = new byte[3];
        for (int i = 0; i < 3; ++i) { int b = stream.ReadByte(); if (b < 0) throw new Exception("Truncated command"); header[i] = (byte)b; }
        var packet = new byte[header[2] + 5]; Array.Copy(header, packet, 3);
        for (int i = 3; i < packet.Length; ++i) { int b = stream.ReadByte(); if (b < 0) throw new Exception("Truncated payload"); packet[i] = (byte)b; }
        if ((packet[0] & 0x80) == 0 || packet[2] < 8) throw new Exception("Request ID missing");
        var plain = new byte[packet.Length - 8];
        plain[0] = (byte)(packet[0] & 0x7f); plain[1] = packet[1]; plain[2] = (byte)(packet[2] - 8);
        Array.Copy(packet, 11, plain, 3, packet.Length - 11);
        return plain;
    }

    [STAThread]
    static int Main(string[] args)
    {
        try
        {
            var assembly = Assembly.LoadFrom(Path.GetFullPath(args[0]));
            formType = assembly.GetType("ProjectServerW.DataForm", true);
            excelType = assembly.GetType("ProjectServerW.FormExcel", true);
            // Pause the real queue consumer, not session code. All exports use real snapshots.
            System.Runtime.CompilerServices.RuntimeHelpers.RunClassConstructor(excelType.TypeHandle);
            excelType.GetField("excelExportWorkerStarted", Static).SetValue(null, true);
            excelType.GetField("excelExportWorkerThread", Static).SetValue(null, Thread.CurrentThread);

            Test("START while WorkData is queued: exact Work/End/new row partition", delegate {
                var f = NewForm(); Row(f, "work1", false); Row(f, "work2", true);
                var work = EndWork(f); Row(f, "end1", false); Row(f, "end2", false);
                Call(f, "PrepareNewLogSession", epoch.AddMinutes(11)); Row(f, "new1", true);
                var jobs = Jobs(); Equal(2, jobs.Count);
                Job(work, "WorkData", "work1,work2"); Job(jobs[1], "EndData", "end1,end2");
                Equal("", Labels(Table(f))); Equal("new1", ((DataRow)Get(f, "pendingRow"))["RealTime"]);
                Complete(f, work, true);
                Equal("new1", ((DataRow)Get(f, "pendingRow"))["RealTime"]);
                Equal(1, Get(f, "excelExportInProgress"));
            });
            Test("START while EndData is queued: old completion never clears new work", delegate {
                var f = NewForm(); Row(f, "work", false); var work = EndWork(f);
                Complete(f, work, true); Row(f, "end", false);
                Call(f, "PrepareNewLogSession", epoch.AddMinutes(11)); Row(f, "new", false);
                var tail = Jobs()[1]; Complete(f, tail, true);
                Job(tail, "EndData", "end"); Equal("new", Labels(Table(f)));
                Equal(0, Get(f, "excelExportInProgress"));
            });
            Test("same timestamps: stale callback cannot finish a different session", delegate {
                var f = NewForm(); Row(f, "old", false); var old = EndWork(f);
                Call(f, "PrepareNewLogSession", epoch); Row(f, "new", false); var current = EndWork(f);
                Equal(Get(old, "sessionStart"), Get(current, "sessionStart"));
                Complete(f, old, true); Equal(true, Get(f, "endDataCapturePending"));
                Equal("new", Labels(Table(f))); Complete(f, current, true);
                Equal("", Labels(Table(f))); Equal(true, Get(f, "endDataCaptureActive"));
            });
            Test("new START before shutdown export includes the pending last WorkData row", delegate {
                var f = NewForm(); Row(f, "old1", false); Row(f, "old2", true);
                Set(f, "stopExportPending", true);
                Call(f, "PrepareNewLogSession", epoch.AddMinutes(10)); Row(f, "new", false);
                Equal(1, Jobs().Count); Job(Jobs()[0], "WorkData", "old1,old2");
                Equal("new", Labels(Table(f))); Equal(false, Get(f, "stopExportPending"));
            });
            Test("tail received during SaveAs survives successful WorkData completion", delegate {
                var f = NewForm(); Row(f, "work", false); var work = EndWork(f);
                Row(f, "tail1", false); Row(f, "tail2", false);
                var table = Table(f);
                var grid = (DataGridView)Get(f, "dataGridView");
                int individualDeletions = 0;
                table.DefaultView.ListChanged += delegate(object sender, System.ComponentModel.ListChangedEventArgs e) {
                    if (e.ListChangedType == System.ComponentModel.ListChangedType.ItemDeleted) ++individualDeletions;
                    if (e.ListChangedType == System.ComponentModel.ListChangedType.ItemAdded) Equal(null, grid.DataSource);
                };
                Complete(f, work, true);
                Equal(table, Table(f));
                Equal(table, grid.DataSource);
                Equal(0, individualDeletions);
                Equal("tail1,tail2", Labels(Table(f))); Job(work, "WorkData", "work");
                Equal(epoch.AddMinutes(10).AddHours(2), Get(f, "endDataDeadline"));
            });
            Test("two-hour expiry while WorkData is still queued excludes next cycle", delegate {
                var f = NewForm(); Row(f, "work", false); var work = EndWork(f); Row(f, "tail", false);
                Equal(true, Call(f, "FinishEndDataCapture", epoch.AddHours(3)));
                var tail = Jobs()[1]; Job(tail, "EndData", "tail");
                Equal(epoch.AddMinutes(10).AddHours(2), Get(tail, "sessionEnd"));
                Call(f, "PrepareNewLogSession", epoch.AddHours(3)); Row(f, "new", false);
                Complete(f, work, true); Complete(f, tail, true); Equal("new", Labels(Table(f)));
            });
            Test("export guard counts all queued snapshots and does not fake enqueue success", delegate {
                var f = NewForm(); Row(f, "work", false); var work = EndWork(f); Row(f, "tail", false);
                Call(f, "PrepareNewLogSession", epoch.AddMinutes(11)); Row(f, "new", false); var next = EndWork(f);
                Equal(3, Get(f, "excelExportInProgress"));
                Complete(f, work, true); Equal(2, Get(f, "excelExportInProgress"));
                Equal(false, Call(f, "StartExcelExportThread", true)); Equal(3, Jobs().Count);
                Complete(f, Jobs()[1], true); Equal(1, Get(f, "excelExportInProgress"));
                Complete(f, next, true); Equal(0, Get(f, "excelExportInProgress"));
            });
            Test("confirmed START and repeat ACK do not split an already-running cycle", delegate {
                var f = NewForm(); Row(f, "old", false); var old = EndWork(f); Row(f, "tail", false);
                Call(f, "ApplyConfirmedStart"); var id = Get(f, "logSessionId"); Row(f, "new", false);
                Call(f, "ApplyConfirmedStart"); Equal(id, Get(f, "logSessionId")); Equal(2, Jobs().Count);
                Complete(f, old, true); Equal("new", Labels(Table(f)));
            });
            Test("failed WorkData callback does not clear the table", delegate {
                var f = NewForm(); Row(f, "work", false); var work = EndWork(f); Row(f, "tail", false);
                Complete(f, work, false); Equal("work,tail", Labels(Table(f)));
                Equal(true, Get(f, "endDataCapturePending"));
            });
            Test("empty tail creates no duplicate WorkData and no spurious EndData", delegate {
                var f = NewForm(); Row(f, "work", false); var work = EndWork(f);
                Call(f, "PrepareNewLogSession", epoch.AddMinutes(10));
                Equal(1, Jobs().Count); Complete(f, work, true); Equal("", Labels(Table(f)));
            });
            Test("manual/emergency export during pending SaveAs contains only EndData", delegate {
                var f = NewForm(); Row(f, "work", false); EndWork(f); Row(f, "tail1", false); Row(f, "tail2", false);
                Equal(true, Call(f, "StartExcelExportThread", true, -1, true, true));
                Job(Jobs()[1], "EndData", "tail1,tail2");
                Equal(epoch.AddMinutes(10), Get(Jobs()[1], "sessionStart"));
                Equal("work,tail1,tail2", Labels(Table(f)));
            });
            Test("STM32 real telemetry/control-log transition with delayed WorkData callback", delegate {
                var f = NewForm(); Telemetry(f, 100, true, false, false); ControlLog(f, 1.0f);
                for (int i = 101; i <= 105; ++i) Telemetry(f, i, false, true, false);
                for (int i = 106; i <= 115; ++i) Telemetry(f, i, false, false, false, 0x2000);
                Equal(1, Jobs().Count); var work = Jobs()[0];
                Equal("100,101,102,103,104,105,106,107,108,109,110,111,112,113,114,115", Times((DataTable)Get(work, "tableSnapshot")));
                Telemetry(f, 116, false, false, false); Telemetry(f, 117, false, false, false);
                Telemetry(f, 118, true, false, false); Complete(f, work, true); ControlLog(f, 2.0f);
                Equal(2, Jobs().Count); Equal("116,117", Times((DataTable)Get(Jobs()[1], "tableSnapshot")));
                Equal("EndData", Get(Jobs()[1], "filePrefix")); Equal("118", Times(Table(f)));
                Equal(1.0f, ((DataTable)Get(work, "tableSnapshot")).Rows[0]["T_filt_0"]);
                Equal(2.0f, Table(f).Rows[0]["T_filt_0"]);
            });
            Test("Queued packets from a replaced connection cannot modify current rows", delegate {
                var f = NewForm();
                Telemetry(f, 100, true, false, false);
                var pending = Get(f, "pendingRow");
                var oldConnection = Get(f, "commandConnection");
                Set(f, "commandConnection", new object()); // Same socket value, new connection.
                Call(f, "AddDataToTableThreadSafe", new byte[0], 0, 8866, Get(f, "clientSocket"), oldConnection);
                var log = new byte[95]; log[0] = 1; log[3] = 89;
                Array.Copy(BitConverter.GetBytes(99.0f), 0, log, 4, 4);
                Call(f, "AppendControlLogFromConnection", log, log.Length, Get(f, "clientSocket"), oldConnection);
                Equal(pending, Get(f, "pendingRow"));
                Equal((ushort)100, Get(f, "lastTelemetryDeviceSeconds"));
                Equal(0, Table(f).Rows.Count);
            });
            Test("STM32 reconnect without observed shutdown waits for valid gate confirmation", delegate {
                var f = NewForm();
                Telemetry(f, 100, true, false, false);
                var session = Get(f, "logSessionId");
                Call(f, "ResetProgramStateAfterLinkLoss");
                Telemetry(f, 150, false, false, false, 0x2000, false);
                Telemetry(f, 151, false, false, false, 0x3000);
                Telemetry(f, 152, false, false, false, 0x1000);
                Equal(true, Get(f, "controllerAutoModeActive"));
                Equal(false, Get(f, "stopExportPending"));
                Telemetry(f, 153, false, false, false, 0x2000);
                Equal(false, Get(f, "controllerAutoModeActive"));
                Equal(true, Get(f, "stopExportPending"));
                Equal(session, Get(f, "logSessionId"));
                Equal(true, ((Label)Get(f, "labelDefrosterState")).Text.Contains("Цикл дефростации завершён"));
            });
            Test("STM32 gate timeout keeps cycle active even after reconnect", delegate {
                var f = NewForm();
                Telemetry(f, 1, true, false, false);
                Call(f, "ResetProgramStateAfterLinkLoss");
                for (int i = 2; i < 25; ++i) Telemetry(f, i, false, true, false);
                Telemetry(f, 25, false, true, false, 0x2000);
                Equal(true, Get(f, "controllerAutoModeActive"));
                Equal(false, Get(f, "stopExportPending"));
                Telemetry(f, 26, false, false, false, 0x2000);
                Equal(true, Get(f, "stopExportPending"));
            });
            Test("Siemens But_Start transition keeps first new sample out of EndData", delegate {
                var f = NewForm(); Set(f, "settingsLoading", true);
                ((RadioButton)Get(f, "radioButtonSiemens")).Checked = true; Set(f, "settingsLoading", false);
                Telemetry(f, 100, false, false, true);
                for (int i = 101; i <= 110; ++i) Telemetry(f, i, false, false, false);
                Equal(1, Jobs().Count); var work = Jobs()[0];
                Telemetry(f, 111, false, false, false); Telemetry(f, 112, false, false, false);
                Telemetry(f, 113, false, false, true); Complete(f, work, true);
                Equal("100,101,102,103,104,105,106,107,108,109,110", Times((DataTable)Get(work, "tableSnapshot")));
                Equal("111,112", Times((DataTable)Get(Jobs()[1], "tableSnapshot")));
                Equal("113", Times(Table(f))); Equal(true, Get(f, "siemensRecordingActive"));
            });
            Test("20 rapid restarts and reverse completions conserve every row exactly once", delegate {
                var f = NewForm(); var ids = new HashSet<Guid>();
                for (int i = 0; i < 20; ++i) {
                    Row(f, "w" + i + "a", false); Row(f, "w" + i + "b", true);
                    var work = EndWork(f); ids.Add((Guid)Get(work, "sessionId")); Row(f, "e" + i, false);
                    Call(f, "PrepareNewLogSession", epoch);
                }
                var jobs = Jobs(); Equal(40, jobs.Count); Equal(20, ids.Count);
                Row(f, "next", true);
                for (int i = 0; i < 20; ++i) {
                    Job(jobs[2 * i], "WorkData", "w" + i + "a,w" + i + "b");
                    Job(jobs[2 * i + 1], "EndData", "e" + i);
                }
                for (int i = jobs.Count - 1; i >= 0; --i) Complete(f, jobs[i], true);
                Equal(0, Get(f, "excelExportInProgress")); Equal("", Labels(Table(f)));
                Equal("next", ((DataRow)Get(f, "pendingRow"))["RealTime"]);
            });
            Test("mandatory PID: 19 editable values round-trip through unchanged group-7 slots", delegate {
                var f = NewForm();
                var listener = new TcpListener(IPAddress.Loopback, 0); listener.Start();
                using (var client = new TcpClient()) {
                    client.Connect((IPEndPoint)listener.LocalEndpoint);
                    using (var controller = listener.AcceptTcpClient()) {
                        listener.Stop(); controller.ReceiveTimeout = 2000;
                        Call(f, "AttachClientSocket", SocketHandle(client), new object());
                        float[] values = { 0, .18f, .003f, 60, 30, 10, .6f, .3f, 15, .2f, .02f, .0002f, 0, 10, 15, 90, 1, .14285714f, 1, 1 };
                        DeferredResponse(f, Group7(values));
                        Call(f, "ReadPlateauSettings", null, EventArgs.Empty);
                        var request = ReadCommand(controller.GetStream()); Equal((byte)3, request[0]); Equal((byte)7, request[1]);
                        var grid = (DataGridView)Get(f, "plateauSettingsGrid");
                        Equal(19, grid.Rows.Count); Equal("tempKp", grid.Rows[0].Cells[0].Value);
                        Equal("humidityOverRH", grid.Rows[18].Cells[0].Value);
                        foreach (DataGridViewRow row in grid.Rows) {
                            if (row.Cells[0].Value.ToString() == "enabled" || row.Cells[0].Value.ToString() == "reservedAlgorithm")
                                throw new Exception("Algorithm selector is visible");
                        }
                        grid.Rows[0].Cells[2].Value = "0.25"; values[0] = 1; values[1] = .25f;
                        DeferredResponse(f, Response(2, 5, new byte[0]));
                        DeferredResponse(f, Group7(values));
                        Call(f, "ApplyPlateauSettings", null, EventArgs.Empty);
                        var write = ReadCommand(controller.GetStream());
                        Equal((byte)2, write[0]); Equal((byte)5, write[1]); Equal((byte)81, write[2]); Equal((byte)7, write[3]);
                        for (int i = 0; i < values.Length; ++i) Equal(values[i], BitConverter.ToSingle(write, 4 + i * 4));
                        var verify = ReadCommand(controller.GetStream()); Equal((byte)3, verify[0]); Equal((byte)7, verify[1]);
                        // Экспорт сохраняет значения контроллера, даже если в таблице есть черновик.
                        grid.Rows[0].Cells[2].Value = "1.5";
                        DeferredResponse(f, Group7(values));
                        var pidSnapshot = (DataTable)Call(f, "BuildPlateauParamsSnapshot");
                        var snapshotRequest = ReadCommand(controller.GetStream());
                        Equal((byte)3, snapshotRequest[0]); Equal((byte)7, snapshotRequest[1]);
                        Equal(19, pidSnapshot.Rows.Count);
                        Equal("tempKp", pidSnapshot.Rows[0][0]);
                        Equal(.25, pidSnapshot.Rows[0][2]);
                        Equal("1.5", grid.Rows[0].Cells[2].Value);
                        if (!((Label)Get(f, "plateauSettingsStatus")).Text.StartsWith("Записано в EEPROM")) throw new Exception("Verification failed");
                        Set(f, "clientSocket", UInt64.MaxValue);
                    }
                }
            });
            Test("Request IDs reject same-command stale replies; process log is a separate correlated request", delegate {
                using (var f = NewForm()) {
                    var listener = new TcpListener(IPAddress.Loopback, 0); listener.Start();
                    using (var client = new TcpClient()) {
                        client.Connect((IPEndPoint)listener.LocalEndpoint);
                        using (var controller = listener.AcceptTcpClient()) {
                            listener.Stop(); controller.ReceiveTimeout = 2000;
                            var connection = new object();
                            Call(f, "AttachClientSocket", SocketHandle(client), connection);
                            Set(f, "startupSequenceCompleted", true); Set(f, "firstDataReceived", true);
                            Call(f, "OnSendStateTimerTick", null, EventArgs.Empty);
                            Equal((byte)8, ReadCommand(controller.GetStream())[1]);
                            var firstId = (ulong)Get(f, "exchangeRequestId");
                            var payload = new byte[61]; payload[0] = 100; payload[2] = 7;
                            payload[9] = 4; payload[16] = 1; payload[44] = 0x40; payload[60] = 2;
                            var plain = Response(3, 8, payload); var wire = Tagged(f, plain);
                            Equal(false, Call(f, "AcceptControllerPacket", plain, connection)); // No legacy downgrade.
                            Equal(true, Call(f, "AcceptControllerPacket", wire, connection));
                            Equal(true, Call(f, "PrepareTelemetryAck", connection, true));
                            Equal(true, Call(f, "WaitControllerExchange", 0)); // No unsolicited second response.
                            Call(f, "AddDataToTableThreadSafe", plain, plain.Length, 8866, Get(f,"clientSocket"), connection);
                            var request = ReadCommand(controller.GetStream()); Equal((byte)15, request[1]);
                            Equal(firstId, BitConverter.ToUInt64(request, 3));
                            var processId = (ulong)Get(f, "exchangeRequestId");
                            if (firstId == processId) throw new Exception("Reused request ID");
                            Equal(false, Call(f, "AcceptControllerPacket", wire, connection));
                            Equal(false, Call(f, "AcceptControllerPacket", Tagged(f, Response(3, 15, new byte[112])), connection));
                            var logPath = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "log.txt");
                            using (var logStream = new System.IO.FileStream(logPath, System.IO.FileMode.Open,
                                System.IO.FileAccess.Read, System.IO.FileShare.ReadWrite))
                            using (var logReader = new System.IO.StreamReader(logStream)) {
                                var log = logReader.ReadToEnd();
                                if (!log.Contains("unexpected GET_PROCESS_LOG payload length (actual=112, expected=113)") ||
                                    !log.Contains("Type=0x03, Code=0x0F, Status=0x00, DataLen=120, requestId=" + processId.ToString("X16")))
                                    throw new Exception("Rejected process response is missing identifying fields in log.txt");
                            }
                            Equal(false, Call(f, "WaitControllerExchange", 0));
                            var values = new byte[113]; Array.Copy(BitConverter.GetBytes(42.0f), values, 4);
                            EnqueueResponse(f, Response(3, 15, values));
                            Call(f, "DrainProcessLogResponse");
                            Equal(1, Table(f).Rows.Count); Equal(42.0f, Table(f).Rows[0]["T_filt_0"]);
                            Equal(false, Call(f,"AcceptControllerPacket",Tagged(f,Response(0,1,values)),connection));
                            Call(f, "OnSendStateTimerTick", null, EventArgs.Empty); ReadCommand(controller.GetStream());
                            Equal(false, Call(f, "AcceptControllerPacket", wire, connection)); // Same code, previous ID.
                            Equal(false, Call(f, "WaitControllerExchange", 0));
                            Equal(true, Call(f, "WaitControllerExchange", 1200));
                            Equal(false, Call(f, "AcceptControllerPacket", Tagged(f,plain), connection)); // Expired current ID.
                            Call(f, "OnSendStateTimerTick", null, EventArgs.Empty); ReadCommand(controller.GetStream());
                            payload[0] = 101; plain = Response(3,8,payload);
                            Equal(true, Call(f,"AcceptControllerPacket",Tagged(f,plain),connection));
                            Equal(true, Call(f,"PrepareTelemetryAck",connection,true));
                            Call(f,"AddDataToTableThreadSafe",plain,plain.Length,8866,Get(f,"clientSocket"),connection);
                            Equal((byte)15,ReadCommand(controller.GetStream())[1]);
                            Equal(false,Call(f,"AcceptControllerPacket",Tagged(f,Response(3,15,values),processId),connection));
                            Equal(1,Table(f).Rows.Count);
                            Array.Copy(BitConverter.GetBytes(43.0f),values,4);
                            EnqueueResponse(f,Response(3,15,values)); Call(f,"DrainProcessLogResponse");
                            Equal(2,Table(f).Rows.Count); Equal(43.0f,Table(f).Rows[1]["T_filt_0"]);
                            Call(f,"OnSendStateTimerTick",null,EventArgs.Empty); ReadCommand(controller.GetStream());
                            payload[0]=102;payload[44]=0;payload[60]=0;plain=Response(3,8,payload);
                            Equal(true,Call(f,"AcceptControllerPacket",Tagged(f,plain),connection));
                            Equal(true,Call(f,"PrepareTelemetryAck",connection,true));
                            Call(f,"AddDataToTableThreadSafe",plain,plain.Length,8866,Get(f,"clientSocket"),connection);
                            Equal(false,controller.GetStream().DataAvailable); // No process request in idle mode.

                        }
                    }
                }
            });
            Test("Event reply releases parameter request without UI pumping or mixing replies", delegate {
                using (var f = NewForm()) {
                    var listener = new TcpListener(IPAddress.Loopback, 0); listener.Start();
                    using (var client = new TcpClient()) {
                        client.Connect((IPEndPoint)listener.LocalEndpoint);
                        using (var controller = listener.AcceptTcpClient()) {
                            listener.Stop(); controller.ReceiveTimeout = 3000;
                            Call(f, "AttachClientSocket", SocketHandle(client), new object());
                            Set(f, "startupSequenceCompleted", true);
                            Set(f, "hasTelemetry", true); Set(f, "unifiedEvents", true); Set(f, "eventsAvailable", true);
                            Call(f, "PollButtonEvents", null, EventArgs.Empty);
                            var stream = controller.GetStream(); var request = ReadCommand(stream);
                            Equal((byte)13, request[1]);
                            Exception peerError = null;
                            var peer = new Thread(delegate() {
                                try {
                                    Thread.Sleep(120);
                                    Equal(false, stream.DataAvailable); // No parameter request before event reply.
                                    EnqueueResponse(f, Response(3, 13, new byte[25]));
                                    request = ReadCommand(stream); Equal((byte)7, request[1]); Equal((byte)7, request[3]);
                                    float[] values = { 1, .18f, .003f, 60, 30, 10, .6f, .3f, 15, .2f, .02f, .0002f, 0, 10, 15, 90, 1, .14285714f, 1, 1 };
                                    DeferredResponse(f, Group7(values));
                                } catch (Exception ex) { peerError = ex; }
                            }); peer.Start();
                            Call(f, "ReadPlateauSettings", null, EventArgs.Empty);
                            if (!peer.Join(4000)) throw new Exception("Event/parameter deadlock");
                            if (peerError != null) throw peerError;
                            Equal(19, ((DataGridView)Get(f, "plateauSettingsGrid")).Rows.Count);
                            Application.DoEvents();
                            Equal(false, Get(f, "eventsAvailable")); Equal(0, Table(f).Rows.Count);
                            Set(f, "eventsAvailable", true);
                            Call(f, "PollButtonEvents", null, EventArgs.Empty);
                            request = ReadCommand(stream); Equal((byte)13, request[1]);
                            var corrupt = Tagged(f, Response(3, 13, new byte[25])); corrupt[5] ^= 1;
                            EnqueueResponse(f, corrupt);
                            Equal(false, Call(f, "WaitControllerExchange", 0));
                            var timeout = System.Diagnostics.Stopwatch.StartNew();
                            Equal(true, Call(f, "WaitControllerExchange", 1100));
                            if (timeout.ElapsedMilliseconds > 1300) throw new Exception("Unbounded diagnostic wait");

                        }
                    }
                }
            });
            Test("Event ACK yields to telemetry and stale connection callbacks are ignored", delegate {
                var f = NewForm(); Set(f, "unifiedEvents", true); Set(f, "eventsAvailable", true);
                Set(f, "moreEventsAfterAck", true); Set(f, "buttonEventsPending", true);
                var ack = Response(3, 14, new byte[0]);
                Call(f, "HandleBackgroundResponse", ack, Get(f, "clientSocket"), new object());
                Equal(true, Get(f, "buttonEventsPending"));
                Call(f, "HandleBackgroundResponse", ack, Get(f, "clientSocket"), Get(f, "commandConnection"));
                Equal(false, Get(f, "buttonEventsPending")); Equal(false, Get(f, "eventsAvailable"));
                Set(f, "startupSequenceCompleted", true); Set(f, "lastSensorDiagTime", DateTime.Now.AddMinutes(-5));
                Call(f, "ScheduleSensorDiagPoll"); Equal(false, Get(f, "sensorDiagInProgress"));
            });
            Test("Startup tolerates a delayed version and a stray group 5 response, even when group 7 times out", delegate {
                using (var form = NewForm()) {
                    var listener = new TcpListener(IPAddress.Loopback, 0);
                    listener.Start();
                    using (var client = new TcpClient()) {
                        client.Connect((IPEndPoint)listener.LocalEndpoint);
                        using (var controller = listener.AcceptTcpClient()) {
                            listener.Stop();
                            controller.ReceiveTimeout = 2500;
                            Call(form, "AttachClientSocket", SocketHandle(client), new object());
                            Set(form, "modeSwitchHasSavedSetting", true);
                            Exception peerError = null;
                            var peer = new Thread(delegate() {
                                try {
                                    var stream = controller.GetStream();
                                    var command = ReadCommand(stream);
                                    Equal((byte)3, command[0]); Equal((byte)2, command[1]);
                                    Thread.Sleep(150); // Longer than the old GET_VERSION timeout.
                                    EnqueueResponse(form, Response(3, 2, System.Text.Encoding.ASCII.GetBytes("3.3.1")));
                                    foreach (byte group in new byte[] { 5, 6 }) {
                                        command = ReadCommand(stream);
                                        Equal((byte)3, command[0]); Equal((byte)7, command[1]); Equal(group, command[3]);
                                        var payload = new byte[group == 5 ? 74 : 62]; payload[0] = group;
                                        if (group == 6) {
                                            var lateGroup5 = new byte[74]; lateGroup5[0] = 5;
                                            EnqueueResponse(form, Response(3, 7, lateGroup5));
                                            payload[48] = 1;
                                        }
                                        EnqueueResponse(form, Response(3, 7, payload));
                                    }
                                    command = ReadCommand(stream);
                                    Equal((byte)3, command[0]); Equal((byte)7, command[1]); Equal((byte)7, command[3]);
                                    var wait = System.Diagnostics.Stopwatch.StartNew();
                                    command = ReadCommand(stream);
                                    if (wait.ElapsedMilliseconds > 1000)
                                        throw new Exception("Group 7 delayed SET_INTERVAL by " + wait.ElapsedMilliseconds + " ms");
                                    Equal((byte)2, command[0]); Equal((byte)2, command[1]);
                                    EnqueueResponse(form, Response(2, 2, new byte[0]));
                                } catch (Exception ex) { peerError = ex; }
                            });
                            peer.IsBackground = true;
                            peer.Start();
                            try {
                                var started = (bool)Call(form, "ExecuteStartupCommandSequence");
                                if (!peer.Join(3000)) throw new Exception("Test controller did not finish");
                                if (peerError != null) throw peerError;
                                Equal(true, started);
                                Equal(true, Get(form, "startupSequenceCompleted"));
                                Equal(0, ((DataGridView)Get(form, "plateauSettingsGrid")).Rows.Count);
                            } finally { Set(form, "clientSocket", UInt64.MaxValue); }
                        }
                    }
                }
            });
            Test("PID graph shares a panel and shifts only displayed values", delegate {
                var source = new DataTable(); source.Columns.Add("RealTime", typeof(string));
                var columns = new Dictionary<string, string> {
                    { "uCommon", "uCommon_TEN" }, { "P", "PID_P" },
                    { "I", "PID_I" }, { "D", "PID_D" }, { "FF", "PID_FF" }
                };
                foreach (var name in columns.Values) source.Columns.Add(name, typeof(double));
                var first = source.Rows.Add("12:00:00");
                var second = source.Rows.Add("12:00:01");
                double value = -4.0;
                foreach (var column in columns.Values) {
                    first[column] = value++;
                    second[column] = value + .25;
                }
                using (var graph = (Form)Activator.CreateInstance(assembly.GetType("ProjectServerW.PidGraphForm"), new object[] { source, epoch })) {
                    int commonPanel = -1;
                    foreach (var curve in (IEnumerable)Get(graph, "series")) {
                        var name = (string)Get(curve, "name");
                        if (name == "uCommon") commonPanel = (int)Get(curve, "panel");
                    }
                    if (commonPanel < 0) throw new Exception("Missing uCommon");
                    double spacing = (double)Get(graph, "pidSpacing");
                    Equal(7.0, spacing);
                    var bands = new List<string> { "uCommon", "P", "I", "D", "FF" };
                    foreach (var curve in (IEnumerable)Get(graph, "series")) {
                        var name = (string)Get(curve, "name");
                        if (!columns.ContainsKey(name)) {
                            if ((name == "Слева" || name == "Справа") && commonPanel == (int)Get(curve, "panel"))
                                throw new Exception("Side request shares the PID panel");
                            continue;
                        }
                        Equal(commonPanel, Get(curve, "panel"));
                        Equal(bands.IndexOf(name) * spacing, Get(curve, "plotOffset"));
                        var values = (IList)Get(curve, "y");
                        Equal(first[columns[name]], values[0]);
                        Equal(second[columns[name]], values[1]);
                        columns.Remove(name);
                    }
                    Equal(0, columns.Count);
                }
            });
            Test("Air graph uses the regulator input and shows all four heater bands", delegate {
                var f = NewForm();
                Telemetry(f, 100, true, false, false);
                var packet = new byte[119]; packet[0] = 1; packet[3] = 113;
                Array.Copy(BitConverter.GetBytes(18.5f), 0, packet, 113, 4);
                Call(f, "AppendControlLogToDataRow", packet, packet.Length);
                Equal(18.5f, Table(f).Rows[0]["T_sup_avg_C"]);
                Telemetry(f, 101, true, false, false);
                ControlLog(f, 0);
                Equal(DBNull.Value, Table(f).Rows[1]["T_sup_avg_C"]);

                var source = new DataTable(); source.Columns.Add("RealTime", typeof(string));
                foreach (var column in new[] { "T0", "T1", "T_filt_0", "T_filt_1", "eT_common",
                    "T_sup_avg_C", "Ten1_Left", "Ten2_Left", "Ten1_Right", "Ten2_Right" })
                    source.Columns.Add(column, typeof(double));
                source.Rows.Add("12:00:00", 21, 22, 23, 24, 5.5, 18.5, 0, 0, 0, 0);
                source.Rows.Add("12:00:01", 21, 22, 23, 24, 5.5, 18.5, 1, 1, 1, 1);
                using (var graph = (Form)Activator.CreateInstance(assembly.GetType("ProjectServerW.PidGraphForm"), new object[] { source, epoch })) {
                    var air = new Dictionary<string, double> {
                        { "Сырая T слева", 21 }, { "Сырая T справа", 22 },
                        { "Фильтр 32 слева", 23 }, { "Фильтр 32 справа", 24 },
                        { "Средняя T регулятора", 18.5 }, { "Уставка", 24 }
                    };
                    int heater = 0;
                    foreach (var curve in (IEnumerable)Get(graph, "series")) {
                        var panel = (int)Get(curve, "panel");
                        var values = (IList)Get(curve, "y");
                        if (panel == 0) {
                            var name = (string)Get(curve, "name");
                            Equal(air[name], values[0]);
                            air.Remove(name);
                        } else if (panel == 2) {
                            Equal((double)heater, values[0]);
                            Equal(heater + 1.0, values[1]);
                            ++heater;
                        }
                    }
                    Equal(0, air.Count);
                    Equal(4, heater);
                }
            });
            Test("PID state shows four heater confirmations and updates without a control log", delegate {
                var f = NewForm(); var table = Table(f); var row = table.NewRow();
                row["RealTime"] = "12:00:00"; row["Typ6"] = (byte)4; row["Act6"] = (byte)1;
                row["Ten1_Left"] = true; row["Ten2_Left"] = false;
                row["Ten1_Right"] = false; row["Ten2_Right"] = true;
                row["_H0"] = false; row["_H1"] = true; row["_H2"] = true; row["_H3"] = false;
                row["PID_P"] = .25f; row["PID_I"] = .75f; row["_Inj"] = true; table.Rows.Add(row);
                Call(f, "ReadPlateauDiagnostics", null, EventArgs.Empty);
                var state = (DataGridView)Get(f, "plateauStateGrid"); Equal(12, state.Rows.Count);
                var reminder = (TextBox)Get(f, "plateauPidReminder");
                Equal(true, reminder.ReadOnly); Equal(true, reminder.Multiline);
                if (!reminder.Text.Contains("uCommon_TEN = heatScale") || !reminder.Text.Contains("FF = ffMaxEq")) throw new Exception("PID reminder resource missing");
                var stateSplit = (SplitContainer)Get(f, "plateauStateSplit");
                Equal(Orientation.Horizontal, stateSplit.Orientation);
                Equal(stateSplit.Panel1, state.Parent); Equal(stateSplit.Panel2, reminder.Parent);
                Equal("1", state.Rows[8].Cells[1].Value); Equal("0", state.Rows[9].Cells[1].Value);
                Equal("0", state.Rows[10].Cells[1].Value); Equal("1", state.Rows[11].Cells[1].Value);
                Equal("0.250", state.Rows[2].Cells[1].Value);
                Equal("0.750", state.Rows[3].Cells[1].Value);
                Equal("1", state.Rows[7].Cells[1].Value);
                row["_Inj"] = false; Call(f, "UpdatePlateauStateFromLog", row);
                Equal("0", state.Rows[7].Cells[1].Value);
                row["_Inj"] = DBNull.Value; Call(f, "UpdatePlateauStateFromLog", row);
                Equal("нет данных", state.Rows[7].Cells[1].Value);
                // A stopped controller still sends telemetry, with all four feedback bits OFF.
                Telemetry(f, 200, false, false, false);
                for (int i = 8; i < 12; ++i) Equal("0", state.Rows[i].Cells[1].Value);
                Equal("0.250", state.Rows[2].Cells[1].Value);
                Equal("0.750", state.Rows[3].Cells[1].Value);
                row["Act6"] = (byte)0; Call(f, "UpdatePlateauHeaterStates", row);
                for (int i = 8; i < 12; ++i) Equal("нет данных", state.Rows[i].Cells[1].Value);
                row["Act6"] = (byte)1; row["Ten2_Left"] = DBNull.Value; Call(f, "UpdatePlateauHeaterStates", row);
                Equal("1", state.Rows[8].Cells[1].Value); Equal("нет данных", state.Rows[9].Cells[1].Value);
            });
            Console.WriteLine("Passed " + passed + " regression tests.");
            return 0;
        }
        catch (Exception ex) { Console.Error.WriteLine(ex); return 1; }
    }
}
