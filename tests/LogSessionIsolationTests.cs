// Runs the real C++/CLI session methods with the Excel worker paused before dequeue.
// No controller, Excel application, visible window, or two-hour wait is required.
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.IO;
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
            Console.WriteLine("Passed " + passed + " session-isolation tests.");
            return 0;
        }
        catch (Exception ex) { Console.Error.WriteLine(ex); return 1; }
    }
}
