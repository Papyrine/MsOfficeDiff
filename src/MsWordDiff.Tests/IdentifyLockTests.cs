[NotInParallel("IdentifyLock")]
public class IdentifyLockTests
{
    [Test]
    public async Task SerializesConcurrentAcquires()
    {
        using var first = await Word.AcquireIdentifyLock();

        var secondTask = Word.AcquireIdentifyLock();
        await Task.Delay(300);
        await Assert.That(secondTask.IsCompleted).IsFalse()
            .Because("The second acquire should block while the first holds the lock");

        first.Dispose();

        using var second = await secondTask;
        await Assert.That(second).IsNotNull();
    }
}
