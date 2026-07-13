[NotInParallel("MicrosoftWord")]
public class SourceDocumentsVisibleTests
{
    [Test]
    public async Task SourceDocumentsRemainOpenAfterCompare()
    {
        var wordType = Type.GetTypeFromProgID("Word.Application");
        if (wordType == null)
        {
            Skip.Test("Microsoft Word is not installed");
        }

        // Track the WINWORD this test spawns so it can be reaped in the finally: word.Quit can
        // leave the process alive while COM RCWs for the child documents are still rooted, which
        // otherwise pollutes the process-counting ProcessCleanupTests.
        var existingPids = Word.GetWordProcessIds();
        dynamic word = Activator.CreateInstance(wordType)!;
        var process = Word.WaitForNewWordProcess(existingPids, TimeSpan.FromSeconds(5));
        try
        {
            word.DisplayAlerts = 0;
            word.Options.SaveInterval = 0;

            var doc1 = Word.Open(word, ProjectFiles.input_temp_docx.FullPath);
            var doc2 = Word.Open(word, ProjectFiles.input_target_docx.FullPath);

            var compare = Word.LaunchCompare(word, doc1, doc2);

            word.Visible = true;

            // Non-quiet mode: ShowSourceDocuments should be set to both (3)
            Word.ApplyQuiet(false, word);
            await Assert.That((int) word.ActiveWindow.ShowSourceDocuments).IsEqualTo(3);

            // Quiet mode: ShowSourceDocuments should be set to none (0)
            Word.ApplyQuiet(true, word);
            await Assert.That((int) word.ActiveWindow.ShowSourceDocuments).IsEqualTo(0);

            compare.Saved = true;
            word.Quit(SaveChanges: false);
        }
        finally
        {
            Marshal.ReleaseComObject(word);
            if (process is { HasExited: false })
            {
                try
                {
                    process.Kill();
                }
                catch
                {
                    // Process may have exited between the check and the kill.
                }
            }

            process?.Dispose();
        }
    }
}
