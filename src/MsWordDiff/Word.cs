public static partial class Word
{
    public static async Task Launch(string path1, string path2, bool quiet = false)
    {
        var wordType = Type.GetTypeFromProgID("Word.Application");
        if (wordType == null)
        {
            throw new("Microsoft Word is not installed");
        }

        var job = JobObject.Create();

        // Snapshot existing Word PIDs before creating the COM instance, then poll
        // for the newly-spawned WINWORD.EXE so we can put it in the Job Object before
        // any operation that can throw. This is critical for the DiffEngineTray
        // "accept" flow: tray hard-kills diffword.exe (Process.Kill), which only
        // reaps WINWORD via KILL_ON_JOB_CLOSE if WINWORD was actually assigned to
        // our job. A short retry loop is needed because Process.GetProcessesByName
        // can lag the Activator return — particularly under Click-to-Run / AppV
        // where Word spawns through a launcher chain.
        var existingPids = GetWordProcessIds();
        dynamic word = Activator.CreateInstance(wordType)!;
        var process = WaitForNewWordProcess(existingPids, TimeSpan.FromSeconds(5))
            ?? throw new("Failed to locate the WINWORD.EXE process spawned by COM activation");
        JobObject.AssignProcess(job, process.Handle);

        try
        {
            // WdAlertLevel.wdAlertsNone = 0
            word.DisplayAlerts = 0;

            // Disable AutoRecover to prevent "serious error" recovery dialogs
            word.Options.SaveInterval = 0;

            var doc1 = Open(word, path1);

            var doc2 = Open(word, path2);

            var compare = LaunchCompare(word, doc1, doc2);

            word.Visible = true;

            ApplyQuiet(quiet, word);

            HideNavigationPane(word);

            MinimizeRibbon(word);

            // Bring Word to the foreground
            SetForegroundWindow((IntPtr)word.ActiveWindow.Hwnd);

            await process.WaitForExitAsync();

            Marshal.ReleaseComObject(compare);
        }
        catch
        {
            // If setup fails (e.g. invalid file path), gracefully quit Word
            // then force-kill as a fallback to prevent zombie processes.
            QuitAndKill(word, process);
            throw;
        }
        finally
        {
            Marshal.ReleaseComObject(word);
            process.Dispose();
            JobObject.Close(job);
        }

        RestoreRibbon(wordType);
    }

    internal static dynamic LaunchCompare(dynamic word, dynamic doc1, dynamic doc2)
    {
        // WdCompareDestination.wdCompareDestinationNew = 2
        // WdGranularity.wdGranularityWordLevel = 1
        var compare = word.CompareDocuments(
            doc1,
            doc2,
            Destination: 2,
            Granularity: 1,
            CompareFormatting: true,
            CompareCaseChanges: true,
            CompareWhitespace: true,
            CompareTables: true,
            CompareHeaders: true,
            CompareFootnotes: true,
            CompareTextboxes: true,
            CompareFields: true,
            CompareComments: true,
            CompareMoves: true,
            RevisedAuthor: "",
            IgnoreAllComparisonWarnings: true);

        doc1.Close(SaveChanges: false);
        doc2.Close(SaveChanges: false);

        // Mark as saved so Word won't prompt to save on close
        compare.Saved = true;

        compare.AutoSaveOn = false;
        compare.ShowSpellingErrors = false;
        compare.ShowGrammaticalErrors = false;
        return compare;
    }


    internal static void ApplyQuiet(bool quiet, dynamic word)
    {
        if (quiet)
        {
            // WdShowSourceDocuments.wdShowSourceDocumentsNone = 0
            // Hides the source documents, showing only the comparison
            word.ActiveWindow.ShowSourceDocuments = 0;
        }
        else
        {
            // WdShowSourceDocuments.wdShowSourceDocumentsBoth = 3
            // Shows the original and revised documents alongside the comparison
            word.ActiveWindow.ShowSourceDocuments = 3;
        }
    }

    internal static dynamic Open(dynamic word, string path)
    {
        var doc = word.Documents.Open(
            path,
            ConfirmConversions: false,
            ReadOnly: true,
            AddToRecentFiles: false,
            OpenAndRepair: false,
            NoEncodingDialog: true);
        // Hide document window to prevent flickering while preparing comparison
        doc.ActiveWindow.Visible = false;
        Unprotect(doc);
        return doc;
    }

    // CompareDocuments blocks forever on a document with enforced editing protection: Word waits
    // on a "remove protection to compare the documents" prompt that DisplayAlerts cannot suppress,
    // leaving WINWORD.EXE alive with no visible window (Launch only sets word.Visible AFTER the
    // compare). Remove protection from the in-memory copy so the compare proceeds. The document
    // was opened ReadOnly and is never saved, so the file on disk is untouched. The empty-password
    // overload throws — rather than showing its own blocking password prompt — when the document
    // is password protected, so a stray protected input surfaces as a logged "Failed to launch"
    // instead of another silent hang.
    internal static void Unprotect(dynamic doc)
    {
        // WdProtectionType.wdNoProtection = -1
        if (doc.ProtectionType == -1)
        {
            return;
        }

        doc.Unprotect(Password: "");
    }

    static void HideNavigationPane(dynamic word) =>
        word.ActiveWindow.DocumentMap = false;

    static void MinimizeRibbon(dynamic word)
    {
        if (!word.CommandBars.GetPressedMso("MinimizeRibbon"))
        {
            word.CommandBars.ExecuteMso("MinimizeRibbon");
        }
    }

    // RestoreRibbon creates a temporary Word instance solely to un-minimize the
    // ribbon so the user's next normal Word session isn't affected. This instance
    // is assigned to its own Job Object and has a kill fallback to prevent zombies
    // (previously it had neither, making it the primary source of leaked processes).
    static void RestoreRibbon(Type wordType)
    {
        var job = JobObject.Create();
        // Same poll-and-assign-before-throwables pattern as Launch. Previously this
        // path had no fallback when FindNewWordProcess returned null, so a failed
        // Quit() left the process orphaned (process was null → QuitAndKill couldn't
        // force-kill, and the job had nothing assigned to reap).
        var existingPids = GetWordProcessIds();
        dynamic word = Activator.CreateInstance(wordType)!;
        var process = WaitForNewWordProcess(existingPids, TimeSpan.FromSeconds(5))
            ?? throw new("Failed to locate the WINWORD.EXE process spawned by COM activation");
        JobObject.AssignProcess(job, process.Handle);

        try
        {
            word.DisplayAlerts = 0;

            // Must be visible for settings to persist, but minimize to reduce flash
            // WdWindowState.wdWindowStateMinimize = 2
            word.WindowState = 2;
            word.Visible = true;

            if (word.CommandBars.GetPressedMso("MinimizeRibbon"))
            {
                word.CommandBars.ExecuteMso("MinimizeRibbon");
            }

            word.Quit();
        }
        catch
        {
            QuitAndKill(word, process);
        }
        finally
        {
            Marshal.ReleaseComObject(word);
            process.Dispose();
            JobObject.Close(job);
        }
    }

    // Attempts a graceful COM Quit, then force-kills the process as a fallback.
    // All exceptions are swallowed because this runs in error/cleanup paths where
    // COM may already be disconnected or the process may have exited.
    internal static void QuitAndKill(dynamic word, Process? process)
    {
        try { word.Quit(SaveChanges: false); }
        catch { /* COM may already be disconnected */ }

        if (process is { HasExited: false })
        {
            try { process.Kill(); }
            catch { /* Process may have exited between check and kill */ }
        }
    }

    // Snapshots current WINWORD PIDs. Used with FindNewWordProcess to identify
    // the process created by Activator.CreateInstance without needing a window handle.
    internal static HashSet<int> GetWordProcessIds()
    {
        var pids = new HashSet<int>();
        foreach (var p in Process.GetProcessesByName("WINWORD"))
        {
            pids.Add(p.Id);
            p.Dispose();
        }
        return pids;
    }

    // Polls FindNewWordProcess until a new WINWORD appears or the timeout elapses.
    // Process.GetProcessesByName can briefly lag the Activator.CreateInstance return
    // (especially under Click-to-Run / AppV), so a single snapshot at t=0 can miss
    // the spawn. Without this, the new WINWORD is never assigned to the Job Object
    // and survives diffword.exe being killed (e.g. by DiffEngineTray on "accept").
    internal static Process? WaitForNewWordProcess(HashSet<int> existingPids, TimeSpan timeout)
    {
        var watch = Stopwatch.StartNew();
        while (true)
        {
            var found = FindNewWordProcess(existingPids);
            if (found != null)
            {
                return found;
            }

            if (watch.Elapsed > timeout)
            {
                return null;
            }

            Thread.Sleep(50);
        }
    }

    // Finds the WINWORD process that appeared after the snapshot was taken.
    // If multiple new processes appear (rare race condition), keeps the last one found.
    internal static Process? FindNewWordProcess(HashSet<int> existingPids)
    {
        Process? found = null;
        foreach (var p in Process.GetProcessesByName("WINWORD"))
        {
            if (!existingPids.Contains(p.Id))
            {
                found?.Dispose();
                found = p;
            }
            else
            {
                p.Dispose();
            }
        }
        return found;
    }

    [LibraryImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static partial bool SetForegroundWindow(IntPtr hWnd);
}
