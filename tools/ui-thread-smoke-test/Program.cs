using System;
using System.Collections.Generic;
using System.Threading;
using System.Windows.Forms;
using WordOllama.lib;

internal static class Program
{
    private static int Main()
    {
        var ready = new ManualResetEvent(false);
        var done = new ManualResetEvent(false);
        var errors = new List<string>();
        int officeThreadId = 0;

        var officeThread = new Thread(() =>
        {
            try
            {
                UiThread.InitializeOfficeThread();
                officeThreadId = Thread.CurrentThread.ManagedThreadId;
                ready.Set();
                Application.Run();
            }
            catch (Exception ex)
            {
                lock (errors) errors.Add(ex.ToString());
                ready.Set();
                done.Set();
            }
        });
        officeThread.SetApartmentState(ApartmentState.STA);
        officeThread.Start();

        if (!ready.WaitOne(TimeSpan.FromSeconds(5)))
            return Fail("STA initialization timed out.");

        ThreadPool.QueueUserWorkItem(_ =>
        {
            try
            {
                int actualThreadId = UiThread.RunOfficeSync(() => Thread.CurrentThread.ManagedThreadId);
                if (actualThreadId != officeThreadId)
                {
                    lock (errors)
                        errors.Add("Expected STA " + officeThreadId + " but ran on " + actualThreadId + ".");
                }
                UiThread.RunOfficeSync(Application.ExitThread);
            }
            catch (Exception ex)
            {
                lock (errors) errors.Add(ex.ToString());
            }
            finally { done.Set(); }
        });

        if (!done.WaitOne(TimeSpan.FromSeconds(10)))
            return Fail("STA dispatch timed out.");
        officeThread.Join(TimeSpan.FromSeconds(5));

        lock (errors)
        {
            if (errors.Count > 0) return Fail(string.Join(Environment.NewLine, errors));
        }

        Console.WriteLine("PASS: background work was marshaled to the captured Office STA thread.");
        return 0;
    }

    private static int Fail(string message)
    {
        Console.Error.WriteLine("UiThread smoke test failed: " + message);
        return 1;
    }
}
