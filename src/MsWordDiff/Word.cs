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

        var (word, process) = await CreateWordInJob(wordType, job);
        Log.Information("WINWORD {WordPid} assigned to job. Comparing {Path1} and {Path2}", process.Id, path1, path2);

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
            Log.Information("WINWORD {WordPid} exited", process.Id);

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

        await RestoreRibbon(wordType);
    }

    // Serialize the snapshot-launch-identify sequence across concurrent diffword
    // instances (same pattern as diffexcel's SpreadsheetCompare). Without this,
    // instances launched near-simultaneously (eg Verify reporting several failed
    // snapshots from one test run) snapshot overlapping PID sets and can claim the
    // same, or each other's, WINWORD. A misidentified WINWORD is assigned to the
    // wrong Job Object (or none), so killing its diffword (DiffEngineTray "accept")
    // leaves it running as an invisible zombie that still holds locks on the
    // compared files. A file lock is used instead of a Mutex because file locks
    // are not thread-affine, allowing async code within the critical section.
    static readonly string identifyLockPath = Path.Combine(Path.GetTempPath(), "MsWordDiff.identify.lock");

    internal static async Task<FileStream> AcquireIdentifyLock()
    {
        for (var i = 0; i < 300; i++)
        {
            try
            {
                return new(identifyLockPath, FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.None);
            }
            catch (IOException)
            {
                await Task.Delay(100);
            }
        }

        throw new IOException($"Failed to acquire lock file: {identifyLockPath}");
    }

    // Creates the Word COM instance and assigns its WINWORD to the job before any
    // other operation. Snapshot existing Word PIDs before creating the COM instance,
    // then poll for the newly-spawned WINWORD.EXE. This is critical for the
    // DiffEngineTray "accept" flow: tray hard-kills diffword.exe (Process.Kill),
    // which only reaps WINWORD via KILL_ON_JOB_CLOSE if WINWORD was actually
    // assigned to our job. The poll retry loop is needed because
    // Process.GetProcessesByName can lag the Activator return — particularly under
    // Click-to-Run / AppV where Word spawns through a launcher chain.
    // If identification or assignment fails, the just-created instance is quit or
    // killed rather than left orphaned.
    static async Task<(dynamic word, Process process)> CreateWordInJob(Type wordType, IntPtr job)
    {
        using (await AcquireIdentifyLock())
        {
            var existingPids = GetWordProcessIds();
            dynamic word = Activator.CreateInstance(wordType)!;
            Process? process = null;
            try
            {
                process = WaitForNewWordProcess(existingPids, TimeSpan.FromSeconds(5))
                    ?? throw new("Failed to locate the WINWORD.EXE process spawned by COM activation");
                JobObject.AssignProcess(job, process.Handle);
                return (word, process);
            }
            catch
            {
                QuitAndKill(word, process);
                Marshal.ReleaseComObject(word);
                throw;
            }
        }
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
    static async Task RestoreRibbon(Type wordType)
    {
        var job = JobObject.Create();
        var (word, process) = await CreateWordInJob(wordType, job);

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
