using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Windows.Threading;

namespace DdeRTD
{
	public delegate void DDEConnectCallback(IDDEConversation conversation);
	public delegate void DDETransactionCallback(string itemName, bool success, string data);

	public interface IDDEConversationListener
	{
		void OnConnect(IDDEConversation conversation);
		void OnAdviseReceived(string item);
		void OnDisconnect();
	}

	public interface IDDEConversation
	{
		void AdviseStart(string item, DDETransactionCallback callback);
		void AdviseStop(string item, DDETransactionCallback callback);
		void Request(string item, DDETransactionCallback callback);
		void Disconnect();
	}

	public class DDEClient
	{
		class Conversation : IDDEConversation
		{
			DispatcherTimer timer;

			public readonly DDEClient Client;
			public IntPtr Handle
			{
				get;
				protected set;
			}
			public readonly string Service;
			public readonly string Topic;

			IDDEConversationListener listener = null;

			class Transaction
			{
				public Transaction(IntPtr id, DDETransactionCallback callback)
				{
					this.Id = id;
					this.Timestamp = DateTime.UtcNow;
					this.Callback = callback;
				}

				public readonly IntPtr Id;
				public readonly DateTime Timestamp;
				public readonly DDETransactionCallback Callback;
			}

			Dictionary<IntPtr, Transaction> transactions = new Dictionary<IntPtr, Transaction>();

			public Conversation(DDEClient client, string service, string topic, IntPtr hConv, IDDEConversationListener listener)
			{
				this.timer = new DispatcherTimer(TimeSpan.FromSeconds(2), DispatcherPriority.Background, this.OnTimer, client.dispatcher);

				this.Client = client;
				this.Service = service;
				this.Topic = topic;
				this.Handle = hConv;
				this.listener = listener;
			}

			void OnTimer(object sender, EventArgs e)
			{
				if(Handle == IntPtr.Zero)
					return;

				DateTime now = DateTime.UtcNow;
				var fiveSeconds = TimeSpan.FromSeconds(5);

				Transaction timedOut = null;

				foreach(var transaction in this.transactions.Values)
				{
					if((now - transaction.Timestamp) >= fiveSeconds)
					{
						timedOut = transaction;
						break;
					}
				}

				if(timedOut != null)
				{
					Log.Emit("dde-transaction-timeout", "hConv", Handle, "id", timedOut.Id);
					this.Disconnect();
				}
			}

			void Transact(string itemName, int wFmt, int wType, DDETransactionCallback callback)
			{
				Client.dispatcher.BeginInvoke(new Action(delegate()
				{
					if(Handle == IntPtr.Zero)
						return;

					int res = 0;
					using(var hsz = new DdeStringHandle(Client.idInst, itemName))
					{
						using(var hData = new DdeDataHandle(DDEML.DdeClientTransaction(null, 0, Handle, hsz.Handle, wFmt, wType, DDEML.TIMEOUT_ASYNC, ref res)))
						{
							Log.Emit("dde-transact", "hConv", this.Handle, "item", itemName, "wFmt", wFmt, "wType", DDEML.GetTypeString(wType), "callback", callback != null, "id", res);
							if(hData.Handle == IntPtr.Zero)
							{
								var errno = DDEML.DdeGetLastError(Client.idInst);
								Log.Emit("dde-transact-failed", "error", DDEML.GetErrorString(errno) + " (" + errno.ToString() + ")");
								if(callback != null)
									callback(itemName, false, null);
								this.Disconnect();
							}
							else
							{
								if(callback != null)
								{
									if(res == 0)
										callback(itemName, false, null);
									else if(callback != null & res != 0)
									{
										var id = new IntPtr(res);
										transactions.Add(id, new Transaction(id, callback));
									}
								}
							}
						}
					}
				}));
			}

			public void TransactionCompleted(IntPtr transactionId, string itemName, IntPtr hData)
			{
				Transaction trans;

				if(transactions.TryGetValue(transactionId, out trans))
					transactions.Remove(transactionId);
				else
					trans = null;

				if(trans != null)
				{
					bool successful = (hData != IntPtr.Zero);
					string data = null;
					if(hData != IntPtr.Zero && hData != DDEML.IntPtrOne)
					{
						int length = DDEML.DdeGetData(hData, null, 0, 0);
						var buffer = new byte[length];
						length = DDEML.DdeGetData(hData, buffer, buffer.Length, 0);
						data = Encoding.Unicode.GetString(buffer);
					}
					trans.Callback(itemName, successful, data);
				}
			}

			public void AdviseReceived(string itemName)
			{
				this.listener.OnAdviseReceived(itemName);
			}

			public void Disconnected()
			{
				this.Client.conversations.Remove(this.Handle);
				this.Handle = IntPtr.Zero;
				this.listener.OnDisconnect();
			}

			public void Disconnect()
			{
				if(Thread.CurrentThread != this.Client.thread)
					this.Client.dispatcher.BeginInvoke(new Action(this.Disconnect));
				else
				{
					DDEML.DdeDisconnect(Handle);
					Log.Emit("dde-disconnect", "hConv", Handle);
					this.Disconnected();
				};
			}

			public void AdviseStart(string item, DDETransactionCallback callback)
			{
				this.Transact(item, 13, DDEML.XTYP_ADVSTART | DDEML.XTYPF_NODATA | DDEML.XTYPF_ACKREQ, callback);
			}

			public void AdviseStop(string item, DDETransactionCallback callback)
			{
				this.Transact(item, 13, DDEML.XTYP_ADVSTOP, callback);
			}

			public void Request(string item, DDETransactionCallback callback)
			{
 				this.Transact(item, 13, DDEML.XTYP_REQUEST, callback);
			}
		}

		// members
		int idInst = 0;
		Thread thread;
		Dispatcher dispatcher = null;
		Dictionary<IntPtr, Conversation> conversations = new Dictionary<IntPtr, Conversation>();

		public string InitializationMessage
		{
			get;
			protected set;
		}

		public DDEClient()
		{
			AutoResetEvent are = new AutoResetEvent(false);
			thread = new Thread(this.ThreadMain);
			thread.Start(are);
			are.WaitOne();
		}

		public void TryConnect(string service, string topic, IDDEConversationListener listener)
		{
			var op = this.dispatcher.BeginInvoke(new Action(delegate()
			{
				if(idInst == 0)
					listener.OnConnect(null);

				using(var hszService = new DdeStringHandle(idInst, service))
				using(var hszTopic = new DdeStringHandle(idInst, topic))
				{
					var hConv = DDEML.DdeConnect(idInst, hszService.Handle, hszTopic.Handle, IntPtr.Zero);
					if(hConv == IntPtr.Zero)
					{
						var errno = DDEML.DdeGetLastError(idInst);
						var error = DDEML.GetErrorString(errno) + " (" + errno.ToString() + ")";
						Log.Emit("dde-connect-failed", "service", service, "topic", topic, "error", error);
						listener.OnConnect(null);
					}
					else
					{
						Log.Emit("dde-connect", "service", service, "topic", topic, "hConv", hConv);
						var conv = new Conversation(this, service, topic, hConv, listener);
						this.conversations.Add(hConv, conv);
						listener.OnConnect(conv);
					}
				}
			}));
		}

		void ThreadMain(object readyEvent)
		{
			var ddeCallback = new DDEML.DdeCallback(this.DdeCallback);

			try
			{
				dispatcher = Dispatcher.CurrentDispatcher;

				int res = DDEML.DdeInitialize(ref idInst, ddeCallback, DDEML.APPCMD_CLIENTONLY, 0);
				if(res != DDEML.DMLERR_NO_ERROR)
				{
					var err_no = DDEML.DdeGetLastError(idInst);
					string error = DDEML.GetErrorString(err_no) + " (" + err_no.ToString() + ")";
					Log.Emit("dde-init-err", "error", error);
					InitializationMessage = error;
					idInst = 0;
				}
				else
				{
					Log.Emit("dde-init-ok", "id", idInst);
					InitializationMessage = "OK";
				}

				(readyEvent as AutoResetEvent).Set();

				Dispatcher.Run();
			}
			catch(Exception e)
			{
				Log.Emit("dde-thread-exception", "exception", e);

				foreach(var kv in conversations)
				{
					try
					{
						kv.Value.Disconnect();
					}
					catch(Exception e2)
					{
						Log.Emit("dde-thread-exception", "exception", e2);
					}
				}
			}

			GC.KeepAlive(ddeCallback);
		}

		IntPtr DdeCallback(int uType, int uFmt, IntPtr hConv, IntPtr hsz1, IntPtr hsz2, IntPtr hData, IntPtr dwData1, IntPtr dwData2)
		{
			string str1 = DDEML.GetString(idInst, hsz1);
			string str2 = DDEML.GetString(idInst, hsz2);
			Log.Emit("dde-cb", "uType", DDEML.GetTypeString(uType), "uFmt", uFmt, "hConv", hConv, "hsz1", str1, "hsz2", str2, "hData", hData, "dwData1", dwData1, "dwData2", dwData2);

			if(hConv == IntPtr.Zero)
				return IntPtr.Zero;

			Conversation conversation;
			if(!conversations.TryGetValue(hConv, out conversation))
			{
				Log.Emit("dde-cb-conv-not-found");
				return IntPtr.Zero;
			}

			switch(uType)
			{
			case DDEML.XTYP_ADVDATA:
				{
					conversation.AdviseReceived(str2);
					return new IntPtr(DDEML.DDE_FACK);
				}

			case DDEML.XTYP_XACT_COMPLETE:
				{
					conversation.TransactionCompleted(dwData1, str2, hData);
					return IntPtr.Zero;
				}

			case DDEML.XTYP_DISCONNECT:
				{
					conversation.Disconnected();
					return IntPtr.Zero;
				}

			default:
				return IntPtr.Zero;
			}
		}
	}
}
