[NotInParallel("MicrosoftWord")]
public class ProtectedDocumentTests
{
    [Test]
    public async Task OpenRemovesEnforcedProtection()
    {
        // A document with enforced read-only protection is the shape that makes
        // Word.CompareDocuments hang (Word blocks on an un-suppressable "remove protection"
        // prompt). Word.Open must strip that protection from the in-memory copy so the compare
        // proceeds and Word actually becomes visible.
        var wordType = Type.GetTypeFromProgID("Word.Application");
        if (wordType == null)
        {
            Skip.Test("Microsoft Word is not installed");
        }

        var protectedPath = Path.Combine(Path.GetTempPath(), $"msworddiff-protected-{Guid.NewGuid():N}.docx");
        dynamic word = Activator.CreateInstance(wordType)!;
        try
        {
            word.DisplayAlerts = 0;

            // Build a read-only-protected document from the unprotected fixture.
            var source = word.Documents.Open(
                ProjectFiles.input_temp_docx.FullPath,
                ConfirmConversions: false,
                ReadOnly: false,
                AddToRecentFiles: false);
            // WdProtectionType.wdAllowOnlyReading = 3
            source.Protect(Type: 3, NoReset: false, Password: "");
            source.SaveAs2(protectedPath);
            source.Close(SaveChanges: false);

            var opened = Word.Open(word, protectedPath);
            try
            {
                // WdProtectionType.wdNoProtection = -1
                await Assert.That((int) opened.ProtectionType).IsEqualTo(-1);
            }
            finally
            {
                opened.Close(SaveChanges: false);
            }
        }
        finally
        {
            word.Quit(SaveChanges: false);
            Marshal.ReleaseComObject(word);
            if (File.Exists(protectedPath))
            {
                File.Delete(protectedPath);
            }
        }
    }
}
