public static class SpreadsheetCompare
{
    static readonly string[] programFolders =
    [
        Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles),
        Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86)
    ];

    static readonly string[] searchRelativePaths =
    [
        // Office 16 (Microsoft 365 / Office 2016+) - most common
        @"Microsoft Office\root\Office16\DCF\SPREADSHEETCOMPARE.EXE",
        // Click-to-Run installs place the exe inside a virtual filesystem (vfs) directory
        @"Microsoft Office\root\vfs\ProgramFilesX86\Microsoft Office\Office16\DCF\SPREADSHEETCOMPARE.EXE",
        @"Microsoft Office\root\vfs\ProgramFilesX64\Microsoft Office\Office16\DCF\SPREADSHEETCOMPARE.EXE",
        // Office 15 (Office 2013)
        @"Microsoft Office\root\Office15\DCF\SPREADSHEETCOMPARE.EXE",
        @"Microsoft Office\root\vfs\ProgramFilesX86\Microsoft Office\Office15\DCF\SPREADSHEETCOMPARE.EXE",
        @"Microsoft Office\root\vfs\ProgramFilesX64\Microsoft Office\Office15\DCF\SPREADSHEETCOMPARE.EXE"
    ];

    public static string? FindExecutable(string? settingsPath = null)
    {
        if (settingsPath != null &&
            File.Exists(settingsPath))
        {
            return settingsPath;
        }

        foreach (var folder in programFolders)
        {
            foreach (var relative in searchRelativePaths)
            {
                var path = Path.Combine(folder, relative);
                if (File.Exists(path))
                {
                    return path;
                }
            }
        }

        return null;
    }

    static string? FindAppVlp()
    {
        foreach (var folder in programFolders)
        {
            var path = Path.Combine(folder, @"Microsoft Office\root\Client\AppVLP.exe");
            if (File.Exists(path))
            {
                return path;
            }
        }

        return null;
    }

    public static async Task Launch(string path1, string path2, string? exePath = null)
    {
        var exe = FindExecutable(exePath);
        if (exe == null)
        {
            throw new(
                """
                Spreadsheet Compare (SPREADSHEETCOMPARE.EXE) was not found.
                It is included with Office Professional Plus / Microsoft 365 Apps for Enterprise.
                If installed in a custom location, use the 'set-path' command to configure the path.
                """);
        }

        // SPREADSHEETCOMPARE.EXE takes a single argument: a path to a file
        // containing the two workbook paths (one per line)
        var tempFile = TempFiles.Create($"{path1}{Environment.NewLine}{path2}");

        var job = JobObject.Create();

        try
        {
            using var process = await LaunchProcess(exe, tempFile);

            JobObject.AssignProcess(job, process.Handle);
            await WindowLayout.MaximizeAndCenterSplits(process);
            await process.WaitForExitAsync();
        }
        catch when (TempFiles.TryDelete(tempFile))
        {
            // unreachable: TryDeleteTempFile always returns false
            throw;
        }
        finally
        {
            JobObject.Close(job);
        }
    }

    static async Task<Process> LaunchProcess(string exe, string tempFile)
    {
        // Click-to-Run Office installs require launching via AppVLP.exe (the App-V
        // virtualization layer). SPREADSHEETCOMPARE.EXE crashes if launched directly.
        var appVlp = FindAppVlp();

        if (appVlp == null)
        {
            return LaunchDirect(exe, tempFile);
        }

        return await LaunchViaAppVlp(appVlp, exe, tempFile);
    }

    static Process LaunchDirect(string exe, string tempFile) =>
        Process.Start(
            new ProcessStartInfo(exe, tempFile)
            {
                UseShellExecute = true
            })
        ?? throw new("Failed to start Spreadsheet Compare process");

    static readonly string lockFilePath = Path.Combine(TempFiles.TempDirectory, ".lock");

    static async Task<Process> LaunchViaAppVlp(string appVlp, string exe, string tempFile)
    {
        // Serialize the snapshot-launch-identify sequence across concurrent
        // diffexcel instances. Without this, concurrent instances snapshot the
        // same PID set, race to claim the same SPREADSHEETCOMPARE process, and
        // leave others orphaned (not in any job object, so they survive when
        // diffexcel is killed).
        // Uses a file lock instead of a Mutex because file locks are not
        // thread-affine, allowing async code within the critical section.
        using (await AcquireFileLock())
        {
            var existingPids = GetSpreadsheetComparePids();

            using var launcher = Process.Start(
                                     new ProcessStartInfo(appVlp, $"\"{exe}\" {tempFile}")
                                     {
                                         UseShellExecute = false
                                     })
                                 ?? throw new("Failed to start AppVLP process");

            // AppVLP.exe is a launcher that exits after starting the real process.
            // Find the actual SPREADSHEETCOMPARE process and wait on it.
            await launcher.WaitForExitAsync();

            return await WaitForProcess(existingPids)
                   ?? throw new("Spreadsheet Compare did not start. Ensure the application is installed correctly.");
        }
    }

    static async Task<FileStream> AcquireFileLock()
    {
        for (var i = 0; i < 300; i++)
        {
            try
            {
                return new(lockFilePath, FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.None);
            }
            catch (IOException)
            {
                await Task.Delay(100);
            }
        }

        throw new IOException($"Failed to acquire lock file: {lockFilePath}");
    }

    static HashSet<int> GetSpreadsheetComparePids() =>
        GetProcessPids("SPREADSHEETCOMPARE");

    internal static HashSet<int> GetProcessPids(string processName)
    {
        var processes = Process.GetProcessesByName(processName);
        var pids = processes.Select(_ => _.Id).ToHashSet();
        foreach (var process in processes)
        {
            process.Dispose();
        }

        return pids;
    }

    static Task<Process?> WaitForProcess(HashSet<int> existingPids) =>
        WaitForProcess("SPREADSHEETCOMPARE", existingPids);

    internal static async Task<Process?> WaitForProcess(string processName, HashSet<int> existingPids, int maxAttempts = 100)
    {
        for (var i = 0; i < maxAttempts; i++)
        {
            var processes = Process.GetProcessesByName(processName);
            Process? result = null;
            foreach (var process in processes)
            {
                if (result == null && !existingPids.Contains(process.Id))
                {
                    result = process;
                }
                else
                {
                    process.Dispose();
                }
            }

            if (result != null)
            {
                return result;
            }

            await Task.Delay(100);
        }

        return null;
    }
}