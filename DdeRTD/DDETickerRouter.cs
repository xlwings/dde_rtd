using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Diagnostics;
using System.Threading;
using System.Windows.Threading;

namespace DdeRTD
{
	class DDETickerRouter : IRtdServerTickerRouter
	{
		class Topic : IDDETopicListener
		{
			Dictionary<string, Item> items = new Dictionary<string, Item>();

			public class Item : IDDEItemListener
			{
				public readonly string ItemName;

				Topic owner;

				public Item(Topic owner, string itemName)
				{
					this.owner = owner;
					this.ItemName = itemName;

					this.error.IsSubscribedChanged += new RtdServerTickerIsSubscribedChangedEvent(ticker_IsSubscribedChanged);
					this.updated.IsSubscribedChanged += new RtdServerTickerIsSubscribedChangedEvent(ticker_IsSubscribedChanged);
					this.advising.IsSubscribedChanged += new RtdServerTickerIsSubscribedChangedEvent(ticker_IsSubscribedChanged);
					this.data.IsSubscribedChanged += new RtdServerTickerIsSubscribedChangedEvent(ticker_IsSubscribedChanged);
					this.cellCount.IsSubscribedChanged += new RtdServerTickerIsSubscribedChangedEvent(ticker_IsSubscribedChanged);
					this.cellParseMsec.IsSubscribedChanged += new RtdServerTickerIsSubscribedChangedEvent(ticker_IsSubscribedChanged);
				}

				RtdServerTicker data = new RtdServerTicker();
				RtdServerTicker advising = new RtdServerTicker(false);
				RtdServerTicker updated = new RtdServerTicker(0);
				RtdServerTicker error = new RtdServerTicker("OK");
				RtdServerTicker cellCount = new RtdServerTicker(0);
				RtdServerTicker cellParseMsec = new RtdServerTicker(0);
				RtdServerTicker cellUpdated = new RtdServerTicker(DateTime.MinValue);
				RtdServerTicker cellUpdatedCount = new RtdServerTicker(0);
				RtdServerTicker cellUpdatedCells = new RtdServerTicker("");

				class CellTicker : IRtdServerTicker
				{
					Item item;
					RtdServerUpdateNotify notify;
					public readonly int Row;
					public readonly int Col;

					public CellTicker(Item item, int irow, int icol)
					{
						this.item = item;
						this.Row = irow;
						this.Col = icol;

						UpdateValue();
					}

					public bool UpdateValue()
					{
						object oldValue = Value;
						if(item.cells == null)
							Value = ExcelError.Value;
						else if(Row >= item.cells.Length)
							Value = ExcelError.Ref;
						else if(Col >= item.cells[Row].Length)
							Value = ExcelError.Ref;
						else
							Value = item.cells[Row][Col];
						if(!object.Equals(oldValue, Value))
						{
							if(notify != null)
								this.notify();
							return true;
						}
						else
							return false;
					}

					public void Subscribe(RtdServerUpdateNotify notify)
					{
						this.notify = notify;
					}

					public void Unsubscribe()
					{
						this.notify = null;
						lock(this.item.cellTickers)
						{
							this.item.cellTickers.Remove(this);
						}
						this.item.cellCount.Value = this.item.cellTickers.Count;
						item.CountSubscriptions();
					}

					public object Value
					{
						get;
						protected set;
					}
				}

				HashSet<CellTicker> cellTickers = new HashSet<CellTicker>();
				string[][] cells = null;

				public IRtdServerTicker GetField(string field)
				{
					switch(field)
					{
					case null:
						return data;
					case "§advising":
						return advising;
					case "§updated":
						return updated;
					case "§error":
						return error;
					case "§cellCount":
						return cellCount;
					case "§cellParseMsec":
						return cellParseMsec;
					case "§cellUpdatedCount":
						return cellUpdatedCount;
					case "§cellUpdated":
						return cellUpdated;
					case "§cellUpdatedCells":
						return cellUpdatedCells;
					}

					return null;
				}

				public IRtdServerTicker GetCell(string row, string col)
				{
					int irow, icol;
					if(!Int32.TryParse(row, out irow) || !Int32.TryParse(col, out icol))
						return null;

					var ticker = new CellTicker(this, irow, icol);
					lock(this.cellTickers)
					{
						this.cellTickers.Add(ticker);
					}
					this.cellCount.Value = this.cellTickers.Count;
					return ticker;
				}

				public void SetError(string msg)
				{
					this.error.Value = DateTime.Now.ToString("[HH:mm:ss] ") + msg;
				}

				void CountSubscriptions()
				{
					int count = cellTickers.Count;
					if(data.IsSubscribed)
						count++;
					if(error.IsSubscribed)
						count++;
					if(advising.IsSubscribed)
						count++;
					if(updated.IsSubscribed)
						count++;
					if(cellCount.IsSubscribed)
						count++;
					if(cellParseMsec.IsSubscribed)
						count++;

					if(count == 0)
						this.owner.RemoveItem(this);
				}

				void ticker_IsSubscribedChanged(RtdServerTicker ticker, bool subscribed)
				{
					CountSubscriptions();
				}

				bool cellsUpToDate = false;
				void ParseCells()
				{
					if(this.cellTickers.Count > 0 && !this.cellsUpToDate)
					{
						string str = this.data.Value as string;
						this.cellsUpToDate = true;
						if(str != null)
						{
							Stopwatch sw = new Stopwatch();
							sw.Start();
							var s1 = str.Split('\n');
							this.cellsUpToDate = false;
							cells = new string[s1.Length][];
							for(int k = 0; k < s1.Length; k++)
								cells[k] = s1[k].Split('\t');
							StringBuilder sb = new StringBuilder();
							int updatedCount = 0;
							lock(this.cellTickers)
							{
								foreach(var t in this.cellTickers)
								{
									if(t.UpdateValue())
									{
										if(sb.Length > 0)
											sb.Append("|");
										sb.Append(t.Row).Append(",").Append(t.Col);
										updatedCount++;
									}
								}
							}
							if(updatedCount > 0)
							{
								this.cellUpdated.Value = DateTime.Now;
								this.cellUpdatedCells.Value = sb.ToString();
								this.cellUpdatedCount.Value = updatedCount;
							}
							sw.Stop();
							this.cellParseMsec.Value = sw.ElapsedMilliseconds;
						}
					}
				}

				void ClearCells()
				{
					this.cells = null;
					this.cellsUpToDate = true;
					lock(this.cellTickers)
					{
						foreach(var t in this.cellTickers)
							t.UpdateValue();
					}
				}

				void IDDEItemListener.OnData(string data)
				{
					if(data != null)
					{
						this.data.Value = data;
						this.advising.Value = true;

						DateTime now = DateTime.Now;
						this.updated.Value = now;
						this.owner.updated.Value = now;
						this.owner.updatedItem.Value = this.ItemName;

						this.cellsUpToDate = false;
						ProcessingJobQueue.Add(ParseCells);
					}
					else
					{
						this.data.Value = ExcelError.Value;
						this.advising.Value = false;

						this.cellsUpToDate = false;
						ProcessingJobQueue.Add(ClearCells);
					}
				}
			}

			DDETickerRouter owner;

			public readonly string ServiceName;
			public readonly string TopicName;

			public Topic(DDETickerRouter owner, string serviceName, string topicName)
			{
				this.owner = owner;
				this.ServiceName = serviceName;
				this.TopicName = topicName;
			}

			public IRtdServerTicker GetField(string field)
			{
				switch(field)
				{
				case "§connected":
					return this.connected;
				case "§error":
					return this.error;
				case "§itemCount":
					return this.itemCount;
				case "§updated":
					return this.updated;
				case "§updatedItem":
					return this.updatedItem;
				}

				return null;
			}

			public Item GetItem(string itemName)
			{
				Item item;
				if(!items.TryGetValue(itemName, out item))
				{
					items[itemName] = item = new Item(this, itemName);
					this.owner.tracker.AddItemListener(this.ServiceName, this.TopicName, itemName, item);
					this.itemCount.Value = this.items.Count;
				}
				return item;
			}

			public void RemoveItem(Item item)
			{
				this.owner.tracker.RemoveItemListener(this.ServiceName, this.TopicName, item.ItemName, item);
				items.Remove(item.ItemName);
				this.itemCount.Value = this.items.Count;
			}

			RtdServerTicker connected = new RtdServerTicker(false);
			RtdServerTicker error = new RtdServerTicker("OK");
			RtdServerTicker itemCount = new RtdServerTicker(0);
			RtdServerTicker updatedItem = new RtdServerTicker("");
			RtdServerTicker updated = new RtdServerTicker(DateTime.MinValue);

			void IDDETopicListener.OnConnect()
			{
				this.connected.Value = true;
			}

			void IDDETopicListener.OnDisconnect()
			{
				this.connected.Value = false;
			}
		};

		Dictionary<Tuple<string, string>, Topic> topics = new Dictionary<Tuple<string, string>, Topic>();
		RtdServerTicker initialized;
		
		DDEClient client;
		DDETracker tracker;
		
		public DDETickerRouter()
		{
			client = new DDEClient();
			this.initialized = new RtdServerTicker(client.InitializationMessage);

			tracker = new DDETracker(client);
		}

		Topic GetTopic(string serviceName, string topicName)
		{
			var st = Tuple.Create(serviceName, topicName);
			Topic topic;
			if(!topics.TryGetValue(st, out topic))
			{
				topics[st] = topic = new Topic(this, serviceName, topicName);
				this.tracker.AddTopicListener(serviceName, topicName, topic);
			}
			return topic;
		}

		public IRtdServerTicker TryGetTicker(string[] args)
		{
			if(args.Length == 1 && args[0] == "§initialized")
			{
				return initialized;
			}
			else if(args.Length == 3 && args[2].StartsWith("§"))
			{
				var topic = GetTopic(args[0], args[1]);
				return topic.GetField(args[2]);
			}
			else if(args.Length == 3 || args.Length == 4 || args.Length == 5)
			{
				var topic = GetTopic(args[0], args[1]);
				var item = topic.GetItem(args[2]);
				switch(args.Length)
				{
				case 3:
					return item.GetField(null);
				case 4:
					return item.GetField(args[3]);
				case 5:
					return item.GetCell(args[3], args[4]);
				}
				return null;
			}
			else
				return null;
		}
	}
}
