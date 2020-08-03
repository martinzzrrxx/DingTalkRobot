using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Runtime.Remoting.Messaging;
using System.Security.Cryptography;
using System.Text;

namespace dingtalk {
  class Program {

    private const string DEFAULT_INI_PATH = "Robot.ini";

    static void Main(string[] args) {

      string configFile = "";
      try {
        string exePath = System.Reflection.Assembly.GetEntryAssembly().Location;
        string dir = new FileInfo(exePath).DirectoryName;
        configFile = Path.Combine(dir, DEFAULT_INI_PATH);
      } catch (Exception ex) {
        Console.WriteLine("Default config file " + DEFAULT_INI_PATH + " doesn't find");
        configFile = "";  
      }

      string   webHookURL      = "";
      string   secretKey       = "";
      string   textMsg         = "";
      string   markdownMsg     = "";
      string   json            = "";
      string   dailyBuildMsg   = "";
      string[] mobiles         = new string[0];

      int i = 0;
      while (i < args.Length) {

        switch (args[i].ToLower()) {
          case "-f":
            configFile = args[++i];
            break;
          case "-u":
            webHookURL = args[++i];
            break;
          case "-s":
            secretKey = args[++i];
            break;
          case "-t":
            textMsg = args[++i];
            break;
          case "-m":
            markdownMsg = args[++i];
            break;
          case "-a":
            mobiles = Utils.Split(args[++i], ",");
            break;
          case "-j":
            json = args[i++];
            break;
          case "-d":
            dailyBuildMsg = args[i++];
            break;
          default:
            throw new Exception("Unknown arguments: \"" + args[i] + "\"");
        }
        i++;
      }

      try {
        Config config = Config.LoadConfig(configFile);

        if (webHookURL.Length > 0) config.WebHookUrl = webHookURL;
        if (secretKey.Length > 0) config.SecretKey = secretKey;
        if (textMsg.Length > 0) config.TextMsg = textMsg;
        if (markdownMsg.Length > 0) config.MarkdownMsg = markdownMsg;
        if (json.Length > 0) config.Json = json;
        if (mobiles.Length > 0) config.Mobiles = mobiles;

        config.Validate();

        Robot robot = new Robot(config);
        robot.Send();
      } catch (Exception ex) {
        // ignore all exception
        Console.WriteLine("DING TALK ROBOT ERROR: " + ex.Message + " \n stacktrace: " + ex.StackTrace);
      }
    }

    class Config {

      public string WebHookUrl = "";
      public string SecretKey = "";
      public string TextMsg = "";
      public string MarkdownMsg = "";
      public string Json = ""; // file or json data
      public string[] Mobiles = new string[0];

      public Dictionary<string, Dictionary<string, string>> Sections = new Dictionary<string, Dictionary<string, string>>();

      public static Config LoadConfig(string path) {
        string[] lines = File.ReadAllLines(path);
        if (lines == null || lines.Length == 0)
          throw new Exception("Invalid config file: " + path);

        string section = ""; // if the section is not empty, we are parsing the SECTION
        Config config = new Config();
        for (int i = 0; i < lines.Length; i++) {
          string item = lines[i].Trim();
          if (item.Length == 0 || item.StartsWith(";")) continue;

          if (item.StartsWith("[") && item.EndsWith("]")) {
            section = item.Substring(1, item.Length - 2);
            if (section.Length == 0 || section.Trim().Length == 0) throw new Exception("Section name cannot be empty");
            if (config.Sections.ContainsKey(section)) throw new Exception("Duplicate section \"" + section + "\"");
            config.Sections.Add(section, new Dictionary<string, string>());
            continue;
          }

          int idx = item.IndexOf("=");
          if (idx <= 0) throw new Exception("Invalid config entry: \"" + lines[i] + "\"");

          string[] kv = new string[2];
          kv[0] = item.Substring(0, idx).ToUpper();
          kv[1] = item.Substring(idx + 1);


          if (section.Length > 0) {
            config.Sections[section].Add(kv[0], kv[1]);
            continue;
          }

          switch (kv[0]) {
            case "WEBHOOKURL":
              config.WebHookUrl = kv[1];
              break;
            case "SECRETKEY":
              config.SecretKey = kv[1];
              break;
            case "TEXT":
              config.TextMsg = kv[1];
              break;
            case "MARKDOWN":
              config.MarkdownMsg = kv[1];
              break;
            case "JSON":
              config.Json = kv[1];
              break;
            case "MOBILES":
              config.Mobiles = Utils.Split(kv[1], ",");
              break;
            default:
              Console.WriteLine("Unknow config entry: \"" + lines[i] + "\"");
              break;
          }
        }

        return config;
      }
      public void Validate() {
        if (WebHookUrl.Length == 0) throw new Exception("WebHookUrl is required");
        if (SecretKey.Length == 0) throw new Exception("SecretKey is required");
        if (TextMsg.Length == 0 && MarkdownMsg.Length == 0 && Json.Length == 0 && DialyBuildMsg.Length == 0) throw new Exception("Markdown/Text/Json/DialyBuildMsg is required");
        // if Mobiles is null, we will at all people
      }

    }

    // ref: https://ding-doc.dingtalk.com/doc#/serverapi2/qf2nxq

    class Robot {
      private static readonly DateTime EPOCH = new DateTime(1970, 1, 1, 0, 0, 0, DateTimeKind.Utc);


      private const string SUCCESS_RESPONSE = "{\"errcode\":0,\"errmsg\":\"ok\"}";

      private const string TEXT_MSG = @"
{
  ""msgtype"": ""text"", 
  ""text"": {
      ""content"": ""$TEXT$""
  }, 
  ""at"": {
      ""atMobiles"": [
          $MOBILES$
      ], 
      ""isAtAll"": $ISATALL$
  }
}
";
      private const string MARKDOWN_MSG = @"
{
     ""msgtype"": ""markdown"",
     ""markdown"": {
         ""title"":""$TITLE$"",
         ""text"": ""$MARKDOWN$\n""
     },
     ""at"": {
          ""atMobiles"": [
              $MOBILES$
          ],
          ""isAtAll"": $ISATALL$
      }
  }

";

      private Config m_Config = null;

      public Robot(Config config) {
        m_Config = config;
      }

      public void Send() {
        string json = BuildJSONMsg();
        long timestamp = CurrentTimeMillis();
        string signature = Sign(timestamp);

        WebRequest request = WebRequest.Create(m_Config.WebHookUrl + "&timestamp=" + timestamp + "&sign=" + signature);
        request.Method = "POST";
        request.ContentType = "application/json; charset=utf-8";
        request.Timeout = 60 * 1000;

        byte[] bytes = System.Text.Encoding.UTF8.GetBytes(json);
        using (Stream stream = request.GetRequestStream()) {
          stream.Write(bytes, 0, bytes.Length);
        }

        WebResponse response = request.GetResponse();
        using (Stream stream = response.GetResponseStream()) {
          StreamReader reader = new StreamReader(stream);
          string result = reader.ReadToEnd();

          if (result.Equals(SUCCESS_RESPONSE)) {
            Console.WriteLine("Send DingTalk Robot Message Successfully");
          } else {
            throw new Exception(result);
          }
        }
      }

      private string GetMsgReceiver(string msg) {
        if (m_Config.Sections.ContainsKey("Subscription")) {
          foreach (KeyValuePair<string, string> entry in m_Config.Sections["Subscription"]) {
            string[] keywords = Utils.Split(entry.Value, ",");
            for (int i = 0; i < keywords.Length; i++) {
              if (msg.StartsWith(keywords[i]) && msg[keywords[i].Length] == ' ')
                return entry.Key;
            }
          }
        }
        return "";
      }

      private string BuildJSONMsg() {

        if (m_Config.TextMsg.Length > 0) {
          string receiver = GetMsgReceiver(m_Config.TextMsg);
          string msg = TEXT_MSG;
          msg = msg.Replace("$TEXT$", receiver.Length > 0 ? ("@" + receiver + "\n" + m_Config.TextMsg) : m_Config.TextMsg);

          string mobiles = "";
          for (int i = 0; i < m_Config.Mobiles.Length; i++) {
            mobiles += m_Config.Mobiles[i];
            if (i < m_Config.Mobiles.Length - 1) mobiles += ",";
            mobiles += "\n";
          }
          msg = msg.Replace("$MOBILES$", mobiles);
          msg = msg.Replace("$ISATALL$", mobiles.Length > 0 ? "false" : "true");
          return msg;
        } else if (m_Config.MarkdownMsg.Length > 0) {
          string msg = MARKDOWN_MSG;
          msg = msg.Replace("$TITLE$", "TITLE"); // just hard code the title for now
          msg = msg.Replace("$MARKDOWN$", m_Config.MarkdownMsg);

          string mobiles = "";
          for (int i = 0; i < m_Config.Mobiles.Length; i++) {
            mobiles += m_Config.Mobiles[i];
            if (i < m_Config.Mobiles.Length - 1) mobiles += ",";
            mobiles += "\n";
          }
          msg = msg.Replace("$MOBILES$", mobiles);
          msg = msg.Replace("$ISATALL$", mobiles.Length > 0 ? "false" : "true");
          return msg;
        } else if (m_Config.Json.Length > 0) {
          if (File.Exists(m_Config.Json)) {
            return File.ReadAllText(m_Config.Json);
          }
          return m_Config.Json;
        } else {
          // Should never reach here
        }
        return "";
      }


      private string Sign(long timestamp) {
        using (HMACSHA256 hmac = new HMACSHA256(System.Text.Encoding.Default.GetBytes(m_Config.SecretKey))) {
          byte[] hashValue = hmac.ComputeHash(System.Text.Encoding.Default.GetBytes(timestamp + "\n" + m_Config.SecretKey));

          return Uri.EscapeDataString(System.Convert.ToBase64String(hashValue));
        }
      }
      private long CurrentTimeMillis() {
        return (long)(DateTime.UtcNow - EPOCH).TotalMilliseconds;
      }
      
      private List<String> FormatNSoftwareReport(String m) {
        List<String> ret = new List<string>();
        int pos = m.IndexOf("<table");
        int posEnd = m.IndexOf(("</table>"));
        while (pos < posEnd) {
          int rowStart = m.IndexOf("<tr>", pos);
          int rowEnd = m.IndexOf("</tr>", pos);
          pos = rowEnd + "</tr>".Length;

          String rowStr = m.Substring(rowStart + "<tr>".Length, rowEnd - rowStart - "<tr>".Length);
          rowStr = rowStr.Replace("</td>", "");
          String[] cols = rowStr.Split(new string[] { "<td>" }, StringSplitOptions.RemoveEmptyEntries);
          if (cols.Length != 2) continue;
          cols[0] = RemoveTag(cols[0]);
          cols[1] = RemoveTag(cols[1].Replace("<br>", "\r\n"));
          pos = cols[1].IndexOf("Microsoft (R) Program");
          if (pos > 0) cols[1] = cols[1].Substring(pos);
          while (cols[1].Contains("\r\n\r\n")) {
            cols[1] = cols[1].Replace("\r\n\r\n", "\r\n");
          }
          ret.AddRange(cols);
        }
        return ret;
      }
      private String RemoveTag(String input) {
        StringBuilder sb = new StringBuilder();
        bool inTag = false;
        for (int i = 0; i < input.Length; i++) {
          char ch = input[i];
          if (ch == '<')               { inTag = true;  continue; } 
          else if (inTag && ch == '>') { inTag = false; continue; }
          if (!inTag) sb.Append(ch);
        }
        return sb.ToString();
      }
    }

    class Utils {
      public static string[] Split(string str, string delimiter) {
        if (str == null || str.Trim().Length == 0) return new string[0];

        List<string> ret = new List<string>();
        string[] parts = str.Split(new string[] { delimiter }, StringSplitOptions.RemoveEmptyEntries);
        for (int i = 0; i < parts.Length; i++) {
          ret.Add(parts[i].Trim());
        }
        return ret.ToArray();
      }
    }
  }




}
