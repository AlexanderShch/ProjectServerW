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
    static void Telemetry(object form, int time, bool working, bool shutdown, bool butStart)
    {
        var packet = new byte[66]; packet[3] = 60;
        Action<int, ushort> word = delegate(int offset, ushort value) {
            packet[4 + offset] = (byte)value; packet[5 + offset] = (byte)(value >> 8);
        };
        word(0, (ushort)time); packet[6] = 7;
        packet[4 + 9] = 4; packet[4 + 16] = 1;
        word(29, (ushort)(butStart ? 0x4000 : 0));
        word(43, (ushort)(working ? 0x4000 : 0)); packet[4 + 45] = (byte)(shutdown ? 1 : 0);
        Set(form, "firstDataReceived", true);
        Call(form, "AddDataToTableThreadSafe", packet, packet.Length, 8866);
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
        return packet;
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
                Row(f, "tail1", false); Row(f, "tail2", false); Complete(f, work, true);
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
                for (int i = 106; i <= 115; ++i) Telemetry(f, i, false, false, false);
                Equal(1, Jobs().Count); var work = Jobs()[0];
                Equal("100,101,102,103,104,105,106,107,108,109,110,111,112,113,114,115", Times((DataTable)Get(work, "tableSnapshot")));
                Telemetry(f, 116, false, false, false); Telemetry(f, 117, false, false, false);
                Telemetry(f, 118, true, false, false); Complete(f, work, true); ControlLog(f, 2.0f);
                Equal(2, Jobs().Count); Equal("116,117", Times((DataTable)Get(Jobs()[1], "tableSnapshot")));
                Equal("EndData", Get(Jobs()[1], "filePrefix")); Equal("118", Times(Table(f)));
                Equal(1.0f, ((DataTable)Get(work, "tableSnapshot")).Rows[0]["T_filt_0"]);
                Equal(2.0f, Table(f).Rows[0]["T_filt_0"]);
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
                        Set(f, "clientSocket", unchecked((ulong)client.Client.Handle.ToInt64()));
                        float[] values = { 0, .18f, .003f, 60, 30, 10, .6f, .3f, 15, .2f, .02f, .0002f, 0, 10, 15, 90, 1, .14285714f, 1, 1 };
                        Call(f, "EnqueueResponse", Group7(values));
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
                        Call(f, "EnqueueResponse", Response(2, 5, new byte[0]));
                        Call(f, "EnqueueResponse", Group7(values));
                        Call(f, "ApplyPlateauSettings", null, EventArgs.Empty);
                        var write = ReadCommand(controller.GetStream());
                        Equal((byte)2, write[0]); Equal((byte)5, write[1]); Equal((byte)81, write[2]); Equal((byte)7, write[3]);
                        for (int i = 0; i < values.Length; ++i) Equal(values[i], BitConverter.ToSingle(write, 4 + i * 4));
                        var verify = ReadCommand(controller.GetStream()); Equal((byte)3, verify[0]); Equal((byte)7, verify[1]);
                        if (!((Label)Get(f, "plateauSettingsStatus")).Text.StartsWith("Записано в EEPROM")) throw new Exception("Verification failed");
                        Set(f, "clientSocket", UInt64.MaxValue);
                    }
                }
            });
            Test("PID graph uses fixed 0/3/6/9/12 offsets and preserves negative/raw values", delegate {
                var source = new DataTable(); source.Columns.Add("RealTime", typeof(string));
                foreach (var name in new[] { "uCommon_TEN", "uRight_TEN", "PID_I", "PID_D", "PID_FF" }) source.Columns.Add(name, typeof(double));
                source.Rows.Add("12:00:00", 1.75, 1.75, -4.0, 0.0, 0.0);
                source.Rows.Add("12:00:01", 1.75, 1.75, 1.75, 0.0, 0.0);
                var graph = Activator.CreateInstance(assembly.GetType("ProjectServerW.PidGraphForm"), new object[] { source, epoch });
                int curves = 0;
                foreach (var curve in (IEnumerable)Get(graph, "series")) {
                    if ((int)Get(curve, "panel") != 4) continue;
                    double minimum = (double)Get(curve, "plotMin"), maximum = (double)Get(curve, "plotMax");
                    var values = (IList)Get(curve, "y");
                    Equal((double)source.Rows[0][curves + 1], values[0]);
                    Equal((double)source.Rows[1][curves + 1], values[1]);
                    double shift = (double)Get(curve, "offset");
                    Equal(3.0 * curves, shift);
                    foreach (double value in values) if (value + shift < minimum || value + shift > maximum) throw new Exception("Clipped PID value");
                    if (curves == 0) { Equal(0.0, minimum); Equal(2.0, maximum); Equal(0.0, shift); }
                    ++curves;
                }
                Equal(5, curves); Equal(-4.0, source.Rows[0]["PID_I"]);
                ((Form)graph).Dispose();
            });
            Console.WriteLine("Passed " + passed + " regression tests.");
            return 0;
        }
        catch (Exception ex) { Console.Error.WriteLine(ex); return 1; }
    }
}
