namespace DesktopSteps;

// Preserves the reservation order of input-hook events even when their asynchronous
// accessibility captures finish out of order or reenter the UI dispatcher.
internal sealed class OrderedCaptureQueue
{
    private readonly SortedDictionary<long, Action> completed = [];
    private readonly object gate = new();
    private long reserved;
    private long next = 1;
    private bool draining;

    public long Reserve() => Interlocked.Increment(ref reserved);

    // Completion and draining run on the recorder's dispatcher thread.
    public void Complete(long sequence, Action action)
    {
        Enqueue(sequence, action);
        Drain();
    }

    public void Enqueue(long sequence, Action action)
    {
        lock (gate) completed.Add(sequence, action);
    }

    public void Drain()
    {
        if (draining) return;
        draining = true;
        try
        {
            while (true)
            {
                Action? ready;
                lock (gate)
                {
                    if (!completed.Remove(next, out ready)) break;
                    next++;
                }
                ready();
            }
        }
        finally { draining = false; }
    }
}
