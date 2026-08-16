using System.Collections.Concurrent;
using System.Runtime.InteropServices;

namespace ExcelVbaMcp.Excel;

/// <summary>
/// Owns the only STA thread used for Excel. Work is synchronous intentionally: an Excel
/// call must not resume on a thread-pool thread after acquiring a COM reference.
/// </summary>
internal sealed class ExcelComDispatcher : IDisposable
{
    private readonly BlockingCollection<IWorkItem> queue = new();
    private readonly Thread thread;
    private readonly TimeSpan shutdownTimeout;
    private int disposed;

    public ExcelComDispatcher()
        : this(TimeSpan.FromSeconds(5))
    {
    }

    internal ExcelComDispatcher(TimeSpan shutdownTimeout)
    {
        ArgumentOutOfRangeException.ThrowIfLessThanOrEqual(shutdownTimeout, TimeSpan.Zero);
        this.shutdownTimeout = shutdownTimeout;
        thread = new Thread(Run)
        {
            IsBackground = true,
            Name = "Excel COM STA dispatcher",
        };
        thread.SetApartmentState(ApartmentState.STA);
        thread.Start();
    }

    public Task<T> InvokeAsync<T>(Func<T> operation, CancellationToken cancellationToken)
    {
        ObjectDisposedException.ThrowIf(Volatile.Read(ref disposed) != 0, this);
        ArgumentNullException.ThrowIfNull(operation);
        cancellationToken.ThrowIfCancellationRequested();

        WorkItem<T> workItem = new(operation, cancellationToken);
        try
        {
            queue.Add(workItem, cancellationToken);
        }
        catch (InvalidOperationException) when (Volatile.Read(ref disposed) != 0)
        {
            workItem.Cancel();
            throw new ObjectDisposedException(nameof(ExcelComDispatcher));
        }

        return workItem.Completion;
    }

    public void Dispose()
    {
        if (Interlocked.Exchange(ref disposed, 1) != 0)
        {
            return;
        }

        queue.CompleteAdding();
        // Excel exposes no safe way to interrupt a blocked COM call. Do not make MCP host
        // shutdown wait forever; the background STA thread owns the queue until process exit.
        if (thread.Join(shutdownTimeout))
        {
            queue.Dispose();
        }
    }

    private void Run()
    {
        int initializationResult = OleInitialize(IntPtr.Zero);
        bool initialized = initializationResult is 0 or 1;
        try
        {
            if (!initialized)
            {
                Exception initializationException = Marshal.GetExceptionForHR(initializationResult)
                    ?? new COMException("The Excel COM apartment could not be initialized.", initializationResult);
                foreach (IWorkItem workItem in queue.GetConsumingEnumerable())
                {
                    workItem.Fail(initializationException);
                }
            }
            else
            {
                foreach (IWorkItem workItem in queue.GetConsumingEnumerable())
                {
                    if (Volatile.Read(ref disposed) != 0)
                    {
                        workItem.Cancel();
                    }
                    else
                    {
                        workItem.Execute();
                    }
                }
            }
        }
        finally
        {
            if (initialized)
            {
                OleUninitialize();
            }
        }
    }

    private interface IWorkItem
    {
        void Cancel();

        void Execute();

        void Fail(Exception exception);
    }

    private sealed class WorkItem<T> : IWorkItem
    {
        private readonly TaskCompletionSource<T> completion = new(TaskCreationOptions.RunContinuationsAsynchronously);
        private readonly Func<T> operation;
        private readonly CancellationToken cancellationToken;
        private readonly CancellationTokenRegistration cancellationRegistration;

        public WorkItem(Func<T> operation, CancellationToken cancellationToken)
        {
            this.operation = operation;
            this.cancellationToken = cancellationToken;
            cancellationRegistration = cancellationToken.Register(
                static state => ((WorkItem<T>)state!).Cancel(),
                this);
        }

        public Task<T> Completion => completion.Task;

        public void Cancel()
        {
            completion.TrySetCanceled(cancellationToken);
            cancellationRegistration.Dispose();
        }

        public void Execute()
        {
            if (cancellationToken.IsCancellationRequested)
            {
                Cancel();
                return;
            }

            try
            {
                completion.TrySetResult(operation());
            }
            catch (Exception exception)
            {
                completion.TrySetException(exception);
            }
            finally
            {
                cancellationRegistration.Dispose();
            }
        }

        public void Fail(Exception exception)
        {
            completion.TrySetException(exception);
            cancellationRegistration.Dispose();
        }
    }

    [DllImport("ole32.dll")]
    private static extern int OleInitialize(IntPtr reserved);

    [DllImport("ole32.dll")]
    private static extern void OleUninitialize();
}
