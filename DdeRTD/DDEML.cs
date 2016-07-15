using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Runtime.InteropServices;

namespace DdeRTD
{
	static class DDEML
	{
		public static readonly IntPtr IntPtrOne = new IntPtr(1);

		static Dictionary<int, string> DmlErrStrings = new Dictionary<int, string>()
	   {
	      { DDEML.DMLERR_ADVACKTIMEOUT, "DMLERR_ADVACKTIMEOUT" },
	      { DDEML.DMLERR_BUSY, "DMLERR_BUSY" },
	      { DDEML.DMLERR_DATAACKTIMEOUT, "DMLERR_DATAACKTIMEOUT" },
	      { DDEML.DMLERR_DLL_NOT_INITIALIZED, "DMLERR_DLL_NOT_INITIALIZED" },
	      { DDEML.DMLERR_DLL_USAGE, "DMLERR_DLL_USAGE" },
	      { DDEML.DMLERR_EXECACKTIMEOUT, "DMLERR_EXECACKTIMEOUT" },
	      { DDEML.DMLERR_INVALIDPARAMETER, "DMLERR_INVALIDPARAMETER" },
	      { DDEML.DMLERR_LOW_MEMORY, "DMLERR_LOW_MEMORY" },
	      { DDEML.DMLERR_MEMORY_ERROR, "DMLERR_MEMORY_ERROR" },
	      { DDEML.DMLERR_NO_CONV_ESTABLISHED, "DMLERR_NO_CONV_ESTABLISHED" },
	      { DDEML.DMLERR_NOTPROCESSED, "DMLERR_NOTPROCESSED" },
	      { DDEML.DMLERR_POKEACKTIMEOUT, "DMLERR_POKEACKTIMEOUT" },
	      { DDEML.DMLERR_POSTMSG_FAILED, "DMLERR_POSTMSG_FAILED" },
	      { DDEML.DMLERR_REENTRANCY, "DMLERR_REENTRANCY" },
	      { DDEML.DMLERR_SERVER_DIED, "DMLERR_SERVER_DIED" },
	      { DDEML.DMLERR_SYS_ERROR, "DMLERR_SYS_ERROR" },
	      { DDEML.DMLERR_UNADVACKTIMEOUT, "DMLERR_UNADVACKTIMEOUT" },
	      { DDEML.DMLERR_UNFOUND_QUEUE_ID, "DMLERR_UNFOUND_QUEUE_ID" }
	   };

		public static string GetErrorString(int errno)
		{
			string s;
			if(DmlErrStrings.TryGetValue(errno, out s))
				return s;
			else
				return "<unknown DDEML error>";
		}

		static Dictionary<int, string> DmlTypeStrings = new Dictionary<int, string>()
		{
			{ DDEML.XTYP_ADVSTART, "XTYP_ADVSTART" },
			{ DDEML.XTYP_CONNECT, "XTYP_CONNECT" },
			{ DDEML.XTYP_ADVDATA, "XTYP_ADVDATA" },
			{ DDEML.XTYP_ADVREQ, "XTYP_ADVREQ" },
			{ DDEML.XTYP_REQUEST, "XTYP_REQUEST" },
			{ DDEML.XTYP_WILDCONNECT, "XTYP_WILDCONNECT" },
			{ DDEML.XTYP_EXECUTE, "XTYP_EXECUTE" },
			{ DDEML.XTYP_POKE, "XTYP_POKE" },
			{ DDEML.XTYP_ADVSTOP, "XTYP_ADVSTOP" },
			{ DDEML.XTYP_CONNECT_CONFIRM, "XTYP_CONNECT_CONFIRM" },
			{ DDEML.XTYP_DISCONNECT, "XTYP_DISCONNECT" },
			{ DDEML.XTYP_ERROR, "XTYP_ERROR" },
			{ DDEML.XTYP_MONITOR, "XTYP_MONITOR" },
			{ DDEML.XTYP_XACT_COMPLETE, "XTYP_XACT_COMPLETE" },
			{ DDEML.XTYP_REGISTER, "XTYP_REGISTER" },
			{ DDEML.XTYP_UNREGISTER, "XTYP_UNREGISTER" },
			{ DDEML.XTYP_ADVSTART | DDEML.XTYPF_NODATA | DDEML.XTYPF_ACKREQ, "DDEML.XTYP_ADVSTART|DDEML.XTYPF_NODATA|DDEML.XTYPF_ACKREQ" }
		};

		public static string GetTypeString(int uType)
		{
			string ret;
			if(DmlTypeStrings.TryGetValue(uType, out ret))
				return ret;
			else
				return uType.ToString();
		}


		public const int MAX_STRING_SIZE = 255;

		public const int APPCMD_CLIENTONLY = unchecked((int) 0x00000010);
		public const int APPCMD_FILTERINITS = unchecked((int) 0x00000020);
		public const int APPCMD_MASK = unchecked((int) 0x00000FF0);
		public const int APPCLASS_STANDARD = unchecked((int) 0x00000000);
		public const int APPCLASS_MONITOR = unchecked((int) 0x00000001);
		public const int APPCLASS_MASK = unchecked((int) 0x0000000F);

		public const int CBR_BLOCK = unchecked((int) 0xFFFFFFFF);

		public const int CBF_FAIL_SELFCONNECTIONS = unchecked((int) 0x00001000);
		public const int CBF_FAIL_CONNECTIONS = unchecked((int) 0x00002000);
		public const int CBF_FAIL_ADVISES = unchecked((int) 0x00004000);
		public const int CBF_FAIL_EXECUTES = unchecked((int) 0x00008000);
		public const int CBF_FAIL_POKES = unchecked((int) 0x00010000);
		public const int CBF_FAIL_REQUESTS = unchecked((int) 0x00020000);
		public const int CBF_FAIL_ALLSVRXACTIONS = unchecked((int) 0x0003f000);
		public const int CBF_SKIP_CONNECT_CONFIRMS = unchecked((int) 0x00040000);
		public const int CBF_SKIP_REGISTRATIONS = unchecked((int) 0x00080000);
		public const int CBF_SKIP_UNREGISTRATIONS = unchecked((int) 0x00100000);
		public const int CBF_SKIP_DISCONNECTS = unchecked((int) 0x00200000);
		public const int CBF_SKIP_ALLNOTIFICATIONS = unchecked((int) 0x003c0000);

		public const int CF_TEXT = 1;

		public const int CP_WINANSI = 1004;
		public const int CP_WINUNICODE = 1200;

		public const int DDE_FACK = unchecked((int) 0x8000);
		public const int DDE_FBUSY = unchecked((int) 0x4000);
		public const int DDE_FDEFERUPD = unchecked((int) 0x4000);
		public const int DDE_FACKREQ = unchecked((int) 0x8000);
		public const int DDE_FRELEASE = unchecked((int) 0x2000);
		public const int DDE_FREQUESTED = unchecked((int) 0x1000);
		public const int DDE_FAPPSTATUS = unchecked((int) 0x00ff);
		public const int DDE_FNOTPROCESSED = unchecked((int) 0x0000);

		public const int DMLERR_NO_ERROR = unchecked((int) 0x0000);
		public const int DMLERR_FIRST = unchecked((int) 0x4000);
		public const int DMLERR_ADVACKTIMEOUT = unchecked((int) 0x4000);
		public const int DMLERR_BUSY = unchecked((int) 0x4001);
		public const int DMLERR_DATAACKTIMEOUT = unchecked((int) 0x4002);
		public const int DMLERR_DLL_NOT_INITIALIZED = unchecked((int) 0x4003);
		public const int DMLERR_DLL_USAGE = unchecked((int) 0x4004);
		public const int DMLERR_EXECACKTIMEOUT = unchecked((int) 0x4005);
		public const int DMLERR_INVALIDPARAMETER = unchecked((int) 0x4006);
		public const int DMLERR_LOW_MEMORY = unchecked((int) 0x4007);
		public const int DMLERR_MEMORY_ERROR = unchecked((int) 0x4008);
		public const int DMLERR_NOTPROCESSED = unchecked((int) 0x4009);
		public const int DMLERR_NO_CONV_ESTABLISHED = unchecked((int) 0x400A);
		public const int DMLERR_POKEACKTIMEOUT = unchecked((int) 0x400B);
		public const int DMLERR_POSTMSG_FAILED = unchecked((int) 0x400C);
		public const int DMLERR_REENTRANCY = unchecked((int) 0x400D);
		public const int DMLERR_SERVER_DIED = unchecked((int) 0x400E);
		public const int DMLERR_SYS_ERROR = unchecked((int) 0x400F);
		public const int DMLERR_UNADVACKTIMEOUT = unchecked((int) 0x4010);
		public const int DMLERR_UNFOUND_QUEUE_ID = unchecked((int) 0x4011);
		public const int DMLERR_LAST = unchecked((int) 0x4011);

		public const int DNS_REGISTER = unchecked((int) 0x0001);
		public const int DNS_UNREGISTER = unchecked((int) 0x0002);
		public const int DNS_FILTERON = unchecked((int) 0x0004);
		public const int DNS_FILTEROFF = unchecked((int) 0x0008);

		public const int EC_ENABLEALL = unchecked((int) 0x0000);
		public const int EC_ENABLEONE = unchecked((int) 0x0080);
		public const int EC_DISABLE = unchecked((int) 0x0008);
		public const int EC_QUERYWAITING = unchecked((int) 0x0002);

		public const int HDATA_APPOWNED = unchecked((int) 0x0001);

		public const int MF_HSZ_INFO = unchecked((int) 0x01000000);
		public const int MF_SENDMSGS = unchecked((int) 0x02000000);
		public const int MF_POSTMSGS = unchecked((int) 0x04000000);
		public const int MF_CALLBACKS = unchecked((int) 0x08000000);
		public const int MF_ERRORS = unchecked((int) 0x10000000);
		public const int MF_LINKS = unchecked((int) 0x20000000);
		public const int MF_CONV = unchecked((int) 0x40000000);

		public const int MH_CREATE = 1;
		public const int MH_KEEP = 2;
		public const int MH_DELETE = 3;
		public const int MH_CLEANUP = 4;

		public const int QID_SYNC = unchecked((int) 0xFFFFFFFF);
		public const int TIMEOUT_ASYNC = unchecked((int) 0xFFFFFFFF);

		public const int XTYPF_NOBLOCK = unchecked((int) 0x0002);
		public const int XTYPF_NODATA = unchecked((int) 0x0004);
		public const int XTYPF_ACKREQ = unchecked((int) 0x0008);
		public const int XCLASS_MASK = unchecked((int) 0xFC00);
		public const int XCLASS_BOOL = unchecked((int) 0x1000);
		public const int XCLASS_DATA = unchecked((int) 0x2000);
		public const int XCLASS_FLAGS = unchecked((int) 0x4000);
		public const int XCLASS_NOTIFICATION = unchecked((int) 0x8000);
		public const int XTYP_ERROR = unchecked((int) (0x0000 | XCLASS_NOTIFICATION | XTYPF_NOBLOCK));
		public const int XTYP_ADVDATA = unchecked((int) (0x0010 | XCLASS_FLAGS));
		public const int XTYP_ADVREQ = unchecked((int) (0x0020 | XCLASS_DATA | XTYPF_NOBLOCK));
		public const int XTYP_ADVSTART = unchecked((int) (0x0030 | XCLASS_BOOL));
		public const int XTYP_ADVSTOP = unchecked((int) (0x0040 | XCLASS_NOTIFICATION));
		public const int XTYP_EXECUTE = unchecked((int) (0x0050 | XCLASS_FLAGS));
		public const int XTYP_CONNECT = unchecked((int) (0x0060 | XCLASS_BOOL | XTYPF_NOBLOCK));
		public const int XTYP_CONNECT_CONFIRM = unchecked((int) (0x0070 | XCLASS_NOTIFICATION | XTYPF_NOBLOCK));
		public const int XTYP_XACT_COMPLETE = unchecked((int) (0x0080 | XCLASS_NOTIFICATION));
		public const int XTYP_POKE = unchecked((int) (0x0090 | XCLASS_FLAGS));
		public const int XTYP_REGISTER = unchecked((int) (0x00A0 | XCLASS_NOTIFICATION | XTYPF_NOBLOCK));
		public const int XTYP_REQUEST = unchecked((int) (0x00B0 | XCLASS_DATA));
		public const int XTYP_DISCONNECT = unchecked((int) (0x00C0 | XCLASS_NOTIFICATION | XTYPF_NOBLOCK));
		public const int XTYP_UNREGISTER = unchecked((int) (0x00D0 | XCLASS_NOTIFICATION | XTYPF_NOBLOCK));
		public const int XTYP_WILDCONNECT = unchecked((int) (0x00E0 | XCLASS_DATA | XTYPF_NOBLOCK));
		public const int XTYP_MONITOR = unchecked((int) (0x00F0 | XCLASS_NOTIFICATION | XTYPF_NOBLOCK));
		public const int XTYP_MASK = unchecked((int) 0x00F0);
		public const int XTYP_SHIFT = unchecked((int) 0x0004);

		public delegate IntPtr DdeCallback(
			 int uType, int uFmt, IntPtr hConv, IntPtr hsz1, IntPtr hsz2, IntPtr hData, IntPtr dwData1, IntPtr dwData2);

		[DllImport("kernel32.dll")]
		public static extern int GetCurrentThreadId();

		[DllImport("user32.dll", EntryPoint = "DdeAbandonTransaction", CharSet = CharSet.Ansi)]
		public static extern bool DdeAbandonTransaction(int idInst, IntPtr hConv, int idTransaction);

		[DllImport("user32.dll", EntryPoint = "DdeAccessData", CharSet = CharSet.Ansi)]
		public static extern IntPtr DdeAccessData(IntPtr hData, ref int pcbDataSize);

		[DllImport("user32.dll", EntryPoint = "DdeAddData", CharSet = CharSet.Ansi)]
		public static extern IntPtr DdeAddData(IntPtr hData, byte[] pSrc, int cb, int cbOff);

		[DllImport("user32.dll", EntryPoint = "DdeClientTransaction", CharSet = CharSet.Ansi)]
		public static extern IntPtr DdeClientTransaction(
			 IntPtr pData, int cbData, IntPtr hConv, IntPtr hszItem, int wFmt, int wType, int dwTimeout, ref int pdwResult);

		[DllImport("user32.dll", EntryPoint = "DdeClientTransaction", CharSet = CharSet.Ansi)]
		public static extern IntPtr DdeClientTransaction(
			 byte[] pData, int cbData, IntPtr hConv, IntPtr hszItem, int wFmt, int wType, int dwTimeout, ref int pdwResult);

		[DllImport("user32.dll", EntryPoint = "DdeCmpStringHandles", CharSet = CharSet.Ansi)]
		public static extern int DdeCmpStringHandles(IntPtr hsz1, IntPtr hsz2);

		[DllImport("user32.dll", EntryPoint = "DdeConnect", CharSet = CharSet.Ansi)]
		public static extern IntPtr DdeConnect(int idInst, IntPtr hszService, IntPtr hszTopic, IntPtr pCC);

		[DllImport("user32.dll", EntryPoint = "DdeConnectList", CharSet = CharSet.Ansi)]
		public static extern IntPtr DdeConnectList(int idInst, IntPtr hszService, IntPtr hszTopic, IntPtr hConvList, IntPtr pCC);

		[DllImport("user32.dll", EntryPoint = "DdeCreateDataHandle", CharSet = CharSet.Ansi)]
		public static extern IntPtr DdeCreateDataHandle(int idInst, byte[] pSrc, int cb, int cbOff, IntPtr hszItem, int wFmt, int afCmd);

		[DllImport("user32.dll", EntryPoint = "DdeCreateStringHandle", CharSet = CharSet.Ansi)]
		public static extern IntPtr DdeCreateStringHandle(int idInst, string psz, int iCodePage);

		[DllImport("user32.dll", EntryPoint = "DdeDisconnect", CharSet = CharSet.Ansi)]
		public static extern bool DdeDisconnect(IntPtr hConv);

		[DllImport("user32.dll", EntryPoint = "DdeDisconnectList", CharSet = CharSet.Ansi)]
		public static extern bool DdeDisconnectList(IntPtr hConvList);

		[DllImport("user32.dll", EntryPoint = "DdeEnableCallback", CharSet = CharSet.Ansi)]
		public static extern bool DdeEnableCallback(int idInst, IntPtr hConv, int wCmd);

		[DllImport("user32.dll", EntryPoint = "DdeFreeDataHandle", CharSet = CharSet.Ansi)]
		public static extern bool DdeFreeDataHandle(IntPtr hData);

		[DllImport("user32.dll", EntryPoint = "DdeFreeStringHandle", CharSet = CharSet.Ansi)]
		public static extern bool DdeFreeStringHandle(int idInst, IntPtr hsz);

		[DllImport("user32.dll", EntryPoint = "DdeGetData", CharSet = CharSet.Ansi)]
		public static extern int DdeGetData(IntPtr hData, [Out] byte[] pDst, int cbMax, int cbOff);

		[DllImport("user32.dll", EntryPoint = "DdeGetLastError", CharSet = CharSet.Ansi)]
		public static extern int DdeGetLastError(int idInst);

		[DllImport("user32.dll", EntryPoint = "DdeImpersonateClient", CharSet = CharSet.Ansi)]
		public static extern bool DdeImpersonateClient(IntPtr hConv);

		[DllImport("user32.dll", EntryPoint = "DdeInitialize", CharSet = CharSet.Ansi)]
		public static extern int DdeInitialize(ref int pidInst, DdeCallback pfnCallback, int afCmd, int ulRes);

		[DllImport("user32.dll", EntryPoint = "DdeKeepStringHandle", CharSet = CharSet.Ansi)]
		public static extern bool DdeKeepStringHandle(int idInst, IntPtr hsz);

		[DllImport("user32.dll", EntryPoint = "DdeNameService", CharSet = CharSet.Ansi)]
		public static extern IntPtr DdeNameService(int idInst, IntPtr hsz1, IntPtr hsz2, int afCmd);

		[DllImport("user32.dll", EntryPoint = "DdePostAdvise", CharSet = CharSet.Ansi)]
		public static extern bool DdePostAdvise(int idInst, IntPtr hszTopic, IntPtr hszItem);

		[DllImport("user32.dll", EntryPoint = "DdeQueryConvInfo", CharSet = CharSet.Ansi)]
		public static extern int DdeQueryConvInfo(IntPtr hConv, int idTransaction, IntPtr pConvInfo);

		[DllImport("user32.dll", EntryPoint = "DdeQueryNextServer", CharSet = CharSet.Ansi)]
		public static extern IntPtr DdeQueryNextServer(IntPtr hConvList, IntPtr hConvPrev);

		[DllImport("user32.dll", EntryPoint = "DdeQueryString", CharSet = CharSet.Ansi)]
		public static extern int DdeQueryString(int idInst, IntPtr hsz, StringBuilder psz, int cchMax, int iCodePage);

		public static string GetString(int idInst, IntPtr hsz)
		{
			if(hsz == IntPtr.Zero)
				return null;

			int length = DdeQueryString(idInst, hsz, null, 0, CP_WINANSI);
			StringBuilder psz = new StringBuilder(length+1);
			DdeQueryString(idInst, hsz, psz, length + 1, CP_WINANSI);
			return psz.ToString();
		}

		[DllImport("user32.dll", EntryPoint = "DdeReconnect", CharSet = CharSet.Ansi)]
		public static extern IntPtr DdeReconnect(IntPtr hConv);

		[DllImport("user32.dll", EntryPoint = "DdeSetUserHandle", CharSet = CharSet.Ansi)]
		public static extern bool DdeSetUserHandle(IntPtr hConv, int id, IntPtr hUser);

		[DllImport("user32.dll", EntryPoint = "DdeUnaccessData", CharSet = CharSet.Ansi)]
		public static extern bool DdeUnaccessData(IntPtr hData);

		[DllImport("user32.dll", EntryPoint = "DdeUninitialize", CharSet = CharSet.Ansi)]
		public static extern bool DdeUninitialize(int idInst);

		[StructLayout(LayoutKind.Sequential)]
		public struct HSZPAIR
		{
			public IntPtr hszSvc;
			public IntPtr hszTopic;
		}

		[StructLayout(LayoutKind.Sequential)]
		public struct CONVINFO
		{
			public int cb;
			public IntPtr hUser;
			public IntPtr hConvPartner;
			public IntPtr hszSvcPartner;
			public IntPtr hszServiceReq;
			public IntPtr hszTopic;
			public IntPtr hszItem;
			public int wFmt;
			public int wType;
			public int wStatus;
			public int wConvst;
			public int wLastError;
			public IntPtr hConvList;
			public CONVCONTEXT ConvCtxt;
			public IntPtr hwnd;
			public IntPtr hwndPartner;

		} // struct

		[StructLayout(LayoutKind.Sequential)]
		public struct CONVCONTEXT
		{
			public int cb;
			public int wFlags;
			public int wCountryID;
			public int iCodePage;
			public int dwLangID;
			public int dwSecurity;
			[MarshalAs(UnmanagedType.ByValArray, SizeConst = 12)]
			public byte[] filler;

		} // struct

		[StructLayout(LayoutKind.Sequential)]
		public struct MONCBSTRUCT
		{
			public int cb;
			public int dwTime;
			public IntPtr hTask;
			public IntPtr dwRet;
			public int wType;
			public int wFmt;
			public IntPtr hConv;
			public IntPtr hsz1;
			public IntPtr hsz2;
			public IntPtr hData;
			public IntPtr dwData1;
			public IntPtr dwData2;
			public CONVCONTEXT cc;
			public int cbData;
			[MarshalAs(UnmanagedType.ByValArray, SizeConst = 32)]
			public byte[] Data;

		} // struct

		[StructLayout(LayoutKind.Sequential)]
		public struct MONCONVSTRUCT
		{
			public int cb;
			public bool fConnect;
			public int dwTime;
			public IntPtr hTask;
			public IntPtr hszSvc;
			public IntPtr hszTopic;
			public IntPtr hConvClient;
			public IntPtr hConvServer;

		} // struct

		[StructLayout(LayoutKind.Sequential)]
		public struct MONERRSTRUCT
		{
			public int cb;
			public int wLastError;
			public int dwTime;
			public IntPtr hTask;

		} // struct

		[StructLayout(LayoutKind.Sequential)]
		public struct MONHSZSTRUCT
		{
			public int cb;
			public int fsAction;
			public int dwTime;
			public IntPtr hsz;
			public IntPtr hTask;
			public IntPtr str;

		} // struct

		[StructLayout(LayoutKind.Sequential)]
		public struct MONLINKSTRUCT
		{
			public int cb;
			public int dwTime;
			public IntPtr hTask;
			public bool fEstablished;
			public bool fNoData;
			public IntPtr hszSvc;
			public IntPtr hszTopic;
			public IntPtr hszItem;
			public int wFmt;
			public bool fServer;
			public IntPtr hConvClient;
			public IntPtr hConvServer;

		} // struct

		[StructLayout(LayoutKind.Sequential)]
		public struct MONMSGSTRUCT
		{
			public int cb;
			public IntPtr hwndTo;
			public int dwTime;
			public IntPtr hTask;
			public int wMsg;
			public IntPtr wParam;
			public IntPtr lParam;
			public DDEML_MSG_HOOK_DATA dmhd;

		} // struct

		[StructLayout(LayoutKind.Sequential)]
		public struct DDEML_MSG_HOOK_DATA
		{
			public IntPtr uiLo;
			public IntPtr uiHi;
			public int cbData;
			[MarshalAs(UnmanagedType.ByValArray, SizeConst = 32)]
			public byte[] Data;

		} // struct


	} // class

	class DdeStringHandle : IDisposable
	{
		public DdeStringHandle(int idInst, string str)
		{
			this.idInst = idInst;
			Handle = DDEML.DdeCreateStringHandle(idInst, str, DDEML.CP_WINANSI);
		}

		protected int idInst;
		public readonly IntPtr Handle;

		public void Dispose()
		{
			if(Handle != IntPtr.Zero)
				DDEML.DdeFreeStringHandle(idInst, Handle);
		}
	}

	class DdeDataHandle : IDisposable
	{
		public DdeDataHandle(IntPtr hData)
		{
			Handle = hData;
		}

		public readonly IntPtr Handle;

		public void Dispose()
		{
			if(Handle != IntPtr.Zero && Handle != DDEML.IntPtrOne)
				DDEML.DdeFreeDataHandle(Handle);
		}
	}
}
