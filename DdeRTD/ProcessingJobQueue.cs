using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Windows.Threading;

namespace DdeRTD
{
	class ProcessingJobQueue
	{
		static Dispatcher processingDispatcher = null;
		static Thread processingThread = new Thread(ProcessingThreadMain);
		static ManualResetEvent processingDispatcherReady = new ManualResetEvent(false);

		static ProcessingJobQueue()
		{
			processingThread.Start();
		}

		static void ProcessingThreadMain()
		{
			processingDispatcher = Dispatcher.CurrentDispatcher;
			processingDispatcherReady.Set();
			Dispatcher.Run();
		}

		public static void Add(Action task)
		{
			processingDispatcherReady.WaitOne();
			processingDispatcher.BeginInvoke(task);
		}
	}
}
