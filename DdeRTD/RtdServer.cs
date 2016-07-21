using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Threading;

namespace DdeRTD
{
	public interface IRtdServerTicker
	{
		void Subscribe(RtdServerUpdateNotify notify);
		void Unsubscribe();

		object Value
		{
			get;
		}
	}

	delegate void RtdServerTickerIsSubscribedChangedEvent(RtdServerTicker ticker, bool subscribed);

	class RtdServerTicker : IRtdServerTicker
	{
		RtdServerUpdateNotify notify;

		public RtdServerTicker()
		{
			this.value = ExcelError.Value;
		}

		public RtdServerTicker(object value)
		{
			this.value = value;
		}

		public void Unsubscribe()
		{
			this.notify = null;
			if(this.IsSubscribedChanged != null)
				IsSubscribedChanged(this, false);
		}

		public void Subscribe(RtdServerUpdateNotify notify)
		{
			this.notify = notify;
			if(this.IsSubscribedChanged != null)
				IsSubscribedChanged(this, true);
		}

		object value;
		public object Value
		{
			get
			{
				return value;
			}

			set
			{
				if(value != this.value)
				{
					this.value = value;
					if(notify != null)
						notify();
				}
			}
		}

		public bool IsSubscribed
		{
			get
			{
				return notify != null;
			}
		}

		public event RtdServerTickerIsSubscribedChangedEvent IsSubscribedChanged;
	}

	public delegate IRtdServerTicker RtdServerTickerRouter(string[] args);

	public interface IRtdServerTickerRouter
	{
		IRtdServerTicker TryGetTicker(string[] args);
	}

	public delegate void RtdServerUpdateNotify();

	[Guid("4d7c6548-26f9-4cfd-95e4-1ff83ae582f7")]
	[ProgId("dde")]
	public class RtdServer : IRtdServer
	{
		Dictionary<int, IRtdServerTicker> tickers = new Dictionary<int, IRtdServerTicker>();

		bool notified = false;
		System.Action updateNotify;

		HashSet<int> staleTickers = new HashSet<int>();

		Dispatcher dispatcher;

		List<IRtdServerTickerRouter> routers = new List<IRtdServerTickerRouter>();

		public RtdServer()
		{
			this.routers.Add(new DDETickerRouter());
		}

		public dynamic ConnectData(int TopicID, ref Array Strings, ref bool GetNewValues)
		{
			GetNewValues = true;

			string[] args = new string[Strings.Length];
			for(int k = 0; k < args.Length; k++)
				args[k] = (string) Strings.GetValue(k);

			IRtdServerTicker ticker = null;
			foreach(var router in routers)
			{
				ticker = router.TryGetTicker(args);
				if(ticker != null)
					break;
			}

			this.tickers[TopicID] = ticker;

			if(ticker == null)
			{
				return ExcelError.Name;
			}
			else
			{
				RtdServerUpdateNotify notify = new RtdServerUpdateNotify(delegate()
				{
                    lock (this.staleTickers)
                    {
                        this.staleTickers.Add(TopicID);
                    }
					
					if(!this.notified)
					{
						this.notified = true;
						this.dispatcher.BeginInvoke(updateNotify, null);
					}
				});

				ticker.Subscribe(notify);
				return ticker.Value;
			}
		}

		public void DisconnectData(int TopicID)
		{
			IRtdServerTicker ticker;
			if(tickers.TryGetValue(TopicID, out ticker))
			{
				if(ticker != null)
				{
					ticker.Unsubscribe();
					tickers.Remove(TopicID);
                    lock (this.staleTickers)
                    {
                        staleTickers.Remove(TopicID);
                    }
				}
			}
		}

		public int Heartbeat()
		{
			return 1;
		}

		public Array RefreshData(ref int TopicCount)
		{
			this.notified = false;

            int[] stale;
            lock (this.staleTickers)
            {
                stale = this.staleTickers.ToArray();
                this.staleTickers.Clear();
            }

			TopicCount = stale.Length;
			object[,] data = new object[2, TopicCount];
			int n = 0;
			foreach(var id in stale)
			{
				data[0, n] = id;
				data[1, n] = tickers[id].Value;
				n++;
			}

			return data;
		}

		public int ServerStart(IRTDUpdateEvent CallbackObject)
		{
			updateNotify = new System.Action(delegate() {
				CallbackObject.UpdateNotify();
			});
			dispatcher = Dispatcher.CurrentDispatcher;
			return 1;
		}

		public void ServerTerminate()
		{
			foreach(var kv in this.tickers)
			{
				if(kv.Value != null)
					kv.Value.Unsubscribe();
			}
		}

		[ComRegisterFunctionAttribute]
		public static void RegisterFunction(Type t)
		{
			Microsoft.Win32.Registry.ClassesRoot.CreateSubKey(@"CLSID\{" + t.GUID.ToString().ToUpper() + @"}\Programmable");
			var key = Microsoft.Win32.Registry.ClassesRoot.OpenSubKey(@"CLSID\{" + t.GUID.ToString().ToUpper() + @"}\InprocServer32", true);
			if(key != null)
				key.SetValue("", System.Environment.SystemDirectory + @"\mscoree.dll", Microsoft.Win32.RegistryValueKind.String);
		}

		[ComUnregisterFunctionAttribute]
		public static void UnregisterFunction(Type t)
		{
			Microsoft.Win32.Registry.ClassesRoot.DeleteSubKey(@"CLSID\{" + t.GUID.ToString().ToUpper() + @"}\Programmable");
		}
	}
}
