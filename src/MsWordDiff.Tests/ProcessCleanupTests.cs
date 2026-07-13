[NotInParallel("MicrosoftWord")]
public class ProcessCleanupTests
{
    [Test]
    public async Task GetWordProcessIds_ReturnsHashSet()
    {
        var pids = Word.GetWordProcessIds();
        await Assert.That(pids).IsNotNull();
    }

    [Test]
    public async Task FindNewWordProcess_WhenNoNewProcesses_ReturnsNull()
    {
        var existingPids = Word.GetWordProcessIds();
        var result = Word.FindNewWordProcess(existingPids);
        await Assert.That(result).IsNull();
    }

    [Test]
    public void QuitAndKill_WithNullProcess_DoesNotThrow() =>
        Word.QuitAndKill((dynamic)new object(), null);

    [Test]
    public void QuitAndKill_WithExitedProcess_DoesNotThrow()
    {
        var process = Process.Start(
            new ProcessStartInfo
            {
                FileName = "cmd.exe",
                Arguments = "/c exit 0",
                CreateNoWindow = true,
                UseShellExecute = false
            })!;
        process.WaitForExit();

        Word.QuitAndKill((dynamic)new object(), process);
        process.Dispose();
    }

    [Test]
    public async Task QuitAndKill_WithRunningProcess_KillsProcess()
    {
        var process = Process.Start(
            new ProcessStartInfo
            {
                FileName = "ping",
                Arguments = "-n 60 127.0.0.1",
                CreateNoWindow = true,
                UseShellExecute = false
            })!;

        await Assert.That(process.HasExited).IsFalse();

        Word.QuitAndKill((dynamic)new object(), process);

        process.WaitForExit(5000);
        await Assert.That(process.HasExited).IsTrue()
            .Because("QuitAndKill should kill running processes");
        process.Dispose();
    }

    [Test]
    public async Task WaitForNewWordProcess_WhenNoNewProcesses_ReturnsNullAfterTimeout()
    {
        var existingPids = Word.GetWordProcessIds();
        var watch = Stopwatch.StartNew();
        var result = Word.WaitForNewWordProcess(existingPids, TimeSpan.FromMilliseconds(200));
        watch.Stop();

        await Assert.That(result).IsNull();
        await Assert.That(watch.Elapsed).IsGreaterThanOrEqualTo(TimeSpan.FromMilliseconds(200))
            .Because("Should poll for at least the timeout before giving up");
        await Assert.That(watch.Elapsed).IsLessThan(TimeSpan.FromSeconds(2))
            .Because("Should not poll significantly past the timeout");
    }

    [Test]
    public async Task AssignProcess_WithInvalidHandles_ThrowsWin32Exception() =>
        await Assert.That(() => JobObject.AssignProcess(IntPtr.Zero, IntPtr.Zero))
            .Throws<Win32Exception>();

    [Test]
    [Explicit]
    public async Task Launch_WithInvalidPath_DoesNotLeaveZombieProcess()
    {
        var wordType = Type.GetTypeFromProgID("Word.Application");
        if (wordType == null)
        {
            Skip.Test("Microsoft Word is not installed");
        }

        var beforePids = Word.GetWordProcessIds();

        try
        {
            await Word.Launch(
                @"C:\nonexistent\file1.docx",
                @"C:\nonexistent\file2.docx");
        }
        catch
        {
            // Expected - invalid file paths
        }

        // Give Word a moment to fully shut down
        await Task.Delay(3000);

        var afterPids = Word.GetWordProcessIds();
        afterPids.ExceptWith(beforePids);

        // Clean up any zombie processes (safety net)
        foreach (var pid in afterPids)
        {
            try
            {
                using var p = Process.GetProcessById(pid);
                p.Kill();
            }
            catch
            {
                // Process may have already exited
            }
        }

        await Assert.That(afterPids.Count).IsEqualTo(0)
            .Because("No zombie Word processes should remain after a failed Launch");
    }
}
