using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Text;
using System.Reflection;
using System.Web.Script.Serialization;

namespace JustJournal
{
	/// <summary>
	/// Summary description for JustJournal.
	/// </summary>
	public static class JustJournalCore
	{
        public const string Server = "www.justjournal.com";

        private static bool _loggedIn;
		private static string _userName;
		private static string _password;
		
        private static bool _enableSpellCheck;
		private static bool _autoSpellCheck;
		private static bool _detectMusic;
		private static bool _useSsl = true;
		private static string _debug;
// ReSharper disable InconsistentNaming
		private static readonly ArrayList moods = new ArrayList(125);
// ReSharper restore InconsistentNaming
		private static bool _winampPaused;
		private static bool _winampStopped;

		/* Properties */

		public static ArrayList Moods
		{
			get
			{
				return moods;
			}
		}

		public static bool EnableSsl
		{
			get
			{
				return _useSsl;
			}
			set
			{
				_useSsl = value;
			}
		}

		public static string Debug
		{
			get
			{
				return _debug;
			}
			set
			{
				_debug = value;
			}
		}

		public static bool EnableSpellCheck
		{
			get
			{
				return _enableSpellCheck;
			}
			set
			{
				_enableSpellCheck = value;
			}
		}

		public static bool AutoSpellCheck
		{
			get
			{
				return _autoSpellCheck;
			}
			set
			{
				_autoSpellCheck = value;
			}
		}

		public static bool EnableMusicDetection
		{
			get
			{
				return _detectMusic;
			}
			set
			{
				_detectMusic = value;
			}
		}

		public static string Version
		{
			get
			{
                return "JustJournal/" + AssemblyVersion() + " Win";
			}
		}

		public static bool LoggedIn
		{   
			get
			{
				return _loggedIn;
			}
			set
			{
				_loggedIn = value;
			}
		}

		public static string UserName
		{
			get
			{
				return _userName;
			}

			set
			{
				_userName = value;
			}
		}

		public static string Password
		{
			get
			{
				return _password;
			}

			set
			{
				_password = value;
			}
		}

		public static bool WinampPaused
		{
			get
			{
				return _winampPaused;
			}
			set
			{
				_winampPaused = value;
			}
		}
		
		public static bool WinampStopped
		{
			get
			{
				return _winampStopped;
			}
			set
			{
				_winampStopped = value;
			}
		}

		/* Methods */

		// Login status values from /api/login and /api/login/2fa
		public const string LoginOk = "JJ.LOGIN.OK";
		public const string LoginFail = "JJ.LOGIN.FAIL";
		public const string LoginTwoFactorNeeded = "JJ.LOGIN.2FA";

		private static readonly JavaScriptSerializer Json = new JavaScriptSerializer();

		// The session cookie from /api/login, kept in memory only.
		private static CookieContainer _cookies = new CookieContainer();

		private static string _twoFactorMethod;

		/// <summary>
		/// How the second factor is delivered (EMAIL or TOTP) after Login() returns LoginTwoFactorNeeded.
		/// </summary>
		public static string TwoFactorMethod
		{
			get
			{
				return _twoFactorMethod;
			}
		}

		public static string BaseUrl
		{
			get
			{
				return (_useSsl ? "https://" : "http://") + Server;
			}
		}

		static JustJournalCore()
		{
			// .NET 4.5 doesn't offer TLS 1.2 by default
			ServicePointManager.SecurityProtocol |= SecurityProtocolType.Tls12;
			// the web server answers POSTs sent with Expect: 100-continue with a 403
			ServicePointManager.Expect100Continue = false;
		}

		/// <summary>
		/// Log in with UserName and Password.
		/// </summary>
		/// <returns>LoginOk, LoginTwoFactorNeeded (finish with LoginTwoFactor), or LoginFail</returns>
		public static string Login()
		{
			// a new login starts a new session, but a failed attempt (wrong password, rate
			// limit, network error) keeps the session we already had
			var previousCookies = _cookies;
			_cookies = new CookieContainer();
			_twoFactorMethod = null;

			var login = new Dictionary<string, object>
				{
					{"username", _userName},
					{"password", _password}
				};

			string loginStatus = LoginFail;
			try
			{
				HttpStatusCode status;
				var resp = SendJson("POST", BaseUrl + "/api/login", login, out status);
				loginStatus = LoginStatus(resp);

				if (loginStatus == LoginTwoFactorNeeded && resp.ContainsKey("twoFactorMethod"))
					_twoFactorMethod = Convert.ToString(resp["twoFactorMethod"]);
			}
			finally
			{
				if (loginStatus != LoginOk && loginStatus != LoginTwoFactorNeeded)
					_cookies = previousCookies;
			}

			if (loginStatus != LoginFail)
				_loggedIn = loginStatus == LoginOk;
			return loginStatus;
		}

		/// <summary>
		/// Finish a login that returned LoginTwoFactorNeeded.
		/// </summary>
		/// <param name="code">code from the e-mail, authenticator app or recovery list</param>
		/// <param name="retry">true when only the code was wrong; false when the login must start over</param>
		/// <returns>LoginOk or LoginFail</returns>
		public static string LoginTwoFactor(string code, out bool retry)
		{
			var body = new Dictionary<string, object>
				{
					{"code", code},
					{"rememberDevice", false}
				};

			HttpStatusCode status;
			var resp = SendJson("POST", BaseUrl + "/api/login/2fa", body, out status);
			var loginStatus = LoginStatus(resp);

			// 401 JJ.LOGIN.FAIL is a wrong code; 429 means too many wrong codes, and
			// JJ.LOGIN.NONE means the pending login expired
			retry = loginStatus == LoginFail && status == HttpStatusCode.Unauthorized;

			_loggedIn = loginStatus == LoginOk;
			return loginStatus == LoginOk ? LoginOk : LoginFail;
		}

		private static string LoginStatus(Dictionary<string, object> resp)
		{
			if (resp != null && resp.ContainsKey("status"))
				return Convert.ToString(resp["status"]);
			return LoginFail;
		}

		/// <summary>
		/// Send a JSON request with the session cookie and parse the JSON object that comes back.
		/// Error statuses are returned rather than thrown so callers can read the error body.
		/// </summary>
		/// <returns>the parsed object, or an empty one when the body isn't a JSON object</returns>
		public static Dictionary<string, object> SendJson(string method, string uriString, object body,
		                                                  out HttpStatusCode status)
		{
			var resp = Send(method, uriString, body, out status);
			try
			{
				return Json.Deserialize<Dictionary<string, object>>(resp) ?? new Dictionary<string, object>();
			}
			catch (ArgumentException)
			{
				return new Dictionary<string, object>();
			}
			catch (InvalidOperationException)
			{
				return new Dictionary<string, object>();
			}
		}

		private static string Send(string method, string uriString, object body, out HttpStatusCode status)
		{
			var request = (HttpWebRequest) WebRequest.Create(uriString);
			request.Method = method;
			request.UserAgent = Version;
			// the login and mood endpoints only answer a wildcard Accept header
			request.Accept = "*/*";
			request.CookieContainer = _cookies;

			if (body != null)
			{
				byte[] data = Encoding.UTF8.GetBytes(Json.Serialize(body));
				request.ContentType = "application/json";
				request.ContentLength = data.Length;
				using (Stream stream = request.GetRequestStream())
					stream.Write(data, 0, data.Length);
			}

			HttpWebResponse response;
			try
			{
				response = (HttpWebResponse) request.GetResponse();
			}
			catch (WebException ex)
			{
				response = ex.Response as HttpWebResponse;
				if (response == null)
					throw;
			}

			using (response)
			using (var reader = new StreamReader(response.GetResponseStream(), Encoding.UTF8))
			{
				status = response.StatusCode;
				var resp = reader.ReadToEnd();
				_debug = resp;
				return resp;
			}
		}

		public static void RetrieveMoods()
		{
            List<Dictionary<string, object>> list = null;
            try
            {
                HttpStatusCode status;
                var resp = Send("GET", "https://" + Server + "/api/mood", null, out status);
                if (status == HttpStatusCode.OK)
                    list = Json.Deserialize<List<Dictionary<string, object>>>(resp);
            }
            // list stays null and the alert below covers each of these
            catch (IOException)
            {
            }
            catch (ArgumentException)
            {
                // not JSON
            }
            catch (InvalidOperationException)
            {
                // includes WebException
            }

            if (list == null)
            {
               new Alert("Error loading Moods from JJ Server. Please restart just journal and try again later.", "");
               return;
            }

            moods.Clear();  // clear the list
            foreach (var mood in list)
            {
                if (mood.ContainsKey("id") && mood.ContainsKey("title"))
                    moods.Add(new Mood(Convert.ToString(mood["title"]), Convert.ToString(mood["id"])));
            }
		}

		// static class

	    private static String AssemblyVersion()
        {
            Assembly assembly = Assembly.GetCallingAssembly();
            AssemblyName assemblyname = assembly.GetName();
            Version assemblyver = assemblyname.Version;
            return assemblyver.ToString();
        }


	}
}
