using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using Newtonsoft.Json;
using System.Reflection;

namespace DdeRTD
{
	public static class Log
	{
		static StreamWriter outputFile = null;

		static Log()
		{
			string asmLocation = Assembly.GetExecutingAssembly().Location;
			//string asmFolder = Path.GetDirectoryName(asmLocation);
			//string asmName = Path.GetFile(asmLocation);
			string path = asmLocation + ".txt"; // Path.Combine(asmFolder, asmName + ".txt");
			if(File.Exists(path))
			{
				outputFile = new StreamWriter(path, true);
				outputFile.Write("\n-\n");
				outputFile.Flush();
			}
		}

		static JsonSerializer jsonSerializer = new JsonSerializer();

		public static void Emit(string msg, params object[] data)
		{
			if(outputFile != null)
			{
				outputFile.Write(DateTime.Now.ToString("o"));
				outputFile.Write(" " + msg + " ");
				try
				{
					JsonTextWriter json = new JsonTextWriter(outputFile);
					int ndata = data.Length / 2;
					json.WriteStartObject();
					for(int k = 0; k < ndata; k++)
					{
						json.WritePropertyName((string) data[2 * k]);
						jsonSerializer.Serialize(json, data[2 * k + 1]);
					}
					json.WriteEndObject();
				}
				catch(Exception e)
				{
					outputFile.WriteLine("...<" + e.ToString() + ">");
					outputFile.Flush();
					return;
				}
				outputFile.WriteLine();
				outputFile.Flush();
			}
		}
	}
}
