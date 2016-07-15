using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Threading;
using System.Runtime.InteropServices;
using System.Threading;
using System.Diagnostics;

namespace DdeRTD
{
	interface IDDETopicListener
	{
		void OnConnect();
		void OnDisconnect();
	}

	interface IDDEItemListener
	{
		void OnData(string data);
	}

	class DDETracker
	{
		class Topic : IDDEConversationListener
		{
			public class Item
			{
				public enum RequestState
				{
					NoRequestPending,
					RequestPending,
					StaleRequestPending
				};

				Topic owner;
				public readonly string ItemName;
				public RequestState State = RequestState.NoRequestPending;
				protected string data;

				public Item(Topic owner, string itemName)
				{
					this.data = null;
					this.owner = owner;
					this.ItemName = itemName;
				}

				HashSet<IDDEItemListener> listeners = new HashSet<IDDEItemListener>();

				public void AddListener(IDDEItemListener listener)
				{
					this.listeners.Add(listener);
					listener.OnData(this.data);
				}

				public void RemoveListener(IDDEItemListener listener)
				{
					this.listeners.Remove(listener);
					if(listeners.Count == 0)
						this.owner.RemoveItem(this);
				}

				public void OnData(string data)
				{
					if(data != null && this.State == Item.RequestState.StaleRequestPending && this.owner.Request(ItemName))
					{
						this.State = Item.RequestState.RequestPending;
					}
					else
					{
						this.State = RequestState.NoRequestPending;
						this.data = data;
						foreach(var listener in listeners)
							listener.OnData(data);
					}
				}

				public void Refresh()
				{
					if(State == Item.RequestState.NoRequestPending)
					{
						if(this.owner.Request(ItemName))
							State = Item.RequestState.RequestPending;
					}
					else
						State = Item.RequestState.StaleRequestPending;
				}
			}

			DDETracker owner;
			Dictionary<string, Item> items = new Dictionary<string, Item>();
			HashSet<IDDETopicListener> listeners = new HashSet<IDDETopicListener>();
			public readonly string ServiceName;
			public readonly string TopicName;

			public Topic(DDETracker owner, string serviceName, string topicName)
			{
				this.owner = owner;
				this.ServiceName = serviceName;
				this.TopicName = topicName;
			}

			public Item GetItem(string itemName)
			{
				Item item;
				if(!items.TryGetValue(itemName, out item))
				{
					items[itemName] = item = new Item(this, itemName);
					if(conversation != null)
						conversation.AdviseStart(itemName, this.OnAdviseStartCompleted);
				}
				return item;
			}

			void RemoveItem(Item item)
			{
				if(conversation != null)
					conversation.AdviseStop(item.ItemName, null);

				this.items.Remove(item.ItemName);
				if(this.items.Count == 0)
				{
					if(conversation != null)
					{
						conversation.Disconnect();
						conversation = null;
					}
				}

				CheckRemoveTopic();
			}

			public void AddListener(IDDETopicListener listener)
			{
				this.listeners.Add(listener);
				if(this.conversation != null)
					listener.OnConnect();
			}

			public void RemoveListener(IDDETopicListener listener)
			{
				this.listeners.Remove(listener);
				CheckRemoveTopic();
			}

			void CheckRemoveTopic()
			{
				if(this.items.Count == 0 && this.listeners.Count == 0)
					this.owner.RemoveTopic(this);
			}

			IDDEConversation conversation = null;
			bool connecting = false;

			public void TryConnect()
			{
				if(!connecting && conversation == null && items.Count > 0)
				{
					connecting = true;
					this.owner.client.TryConnect(this.ServiceName, this.TopicName, this);
				}
			}

			bool Request(string itemName)
			{
				if(this.conversation != null)
				{
					this.conversation.Request(itemName, this.OnRequestCompleted);
					return true;
				}
				else
					return false;
			}

			// object entry point
			void OnAdviseStartCompleted(string itemName, bool success, string data)
			{
				lock(this.owner.mutex)
				{
					Item item;
					if(this.items.TryGetValue(itemName, out item))
						item.Refresh();
				}
			}

			// object entry point
			void OnRequestCompleted(string itemName, bool success, string data)
			{
				lock(this.owner.mutex)
				{
					Item item;
					if(this.items.TryGetValue(itemName, out item))
					{
						if(success)
							item.OnData(data);
						else
							item.OnData(null);
					}
				}
			}

			// object entry point
			void IDDEConversationListener.OnAdviseReceived(string itemName)
			{
				lock(this.owner.mutex)
				{
					Item item;
					if(this.items.TryGetValue(itemName, out item))
						item.Refresh();
				}
			}

			// object entry point
			void IDDEConversationListener.OnDisconnect()
			{
				lock(this.owner.mutex)
				{
					conversation = null;
					foreach(var item in this.items.Values)
						item.OnData(null);
					foreach(var listener in this.listeners)
						listener.OnDisconnect();
				}
			}

			// object entry point
			void IDDEConversationListener.OnConnect(IDDEConversation conversation)
			{
				lock(this.owner.mutex)
				{
					this.conversation = conversation;
					this.connecting = false;
					if(conversation != null)
					{
						foreach(var listener in this.listeners)
							listener.OnConnect();
						foreach(var item in this.items.Values)
							conversation.AdviseStart(item.ItemName, this.OnAdviseStartCompleted);
					}
				}
			}
		}

		Dictionary<Tuple<string, string>, Topic> topics = new Dictionary<Tuple<string, string>, Topic>();
		List<Topic> topicsList = new List<Topic>();
		DDEClient client;
		object mutex = new object();

		Thread connectThread;

		void ConnectThreadMain()
		{
			try
			{
				while(true)
				{
					Thread.Sleep(5000);
					int k = topicsList.Count;
					while(true)
					{
						// entry point
						lock(this.mutex)
						{
							if(k > topicsList.Count)
								k = topicsList.Count;
							k--;
							if(k < 0)
								break;
							else
								topicsList[k].TryConnect();
						}
					}
				}
			}
			catch(Exception e)
			{
				Log.Emit("dde-connect-thread-exception", "exception", e);
			}
		}

		public DDETracker(DDEClient client)
		{
			this.client = client;

			this.connectThread = new Thread(ConnectThreadMain);
			this.connectThread.Start();
		}

		Topic GetTopic(string service, string topicName)
		{
			var st = Tuple.Create(service, topicName);
			Topic topic;
			if(!topics.TryGetValue(st, out topic))
			{
				topics[st] = topic = new Topic(this, service, topicName);
				this.topicsList.Add(topic);
			}
			return topic;
		}

		void RemoveTopic(Topic topic)
		{
			var st = Tuple.Create(topic.ServiceName, topic.TopicName);
			this.topics.Remove(st);
			this.topicsList.Remove(topic);
		}

		// entry point
		public void AddTopicListener(string service, string topic, IDDETopicListener listener)
		{
			lock(this.mutex)
			{
				var t = GetTopic(service, topic);
				t.AddListener(listener);
			}
		}

		// entry point
		public void AddItemListener(string service, string topicName, string itemName, IDDEItemListener listener)
		{
			lock(this.mutex)
			{
				var topic = GetTopic(service, topicName);
				var item = topic.GetItem(itemName);
				item.AddListener(listener);
			}
		}

		// entry point
		public void RemoveItemListener(string service, string topicName, string itemName, IDDEItemListener listener)
		{
			lock(this.mutex)
			{
				var topic = GetTopic(service, topicName);
				var item = topic.GetItem(itemName);
				item.RemoveListener(listener);
			}
		}
	}
}