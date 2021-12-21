using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace DingTalkRobot {
  class Program {

    private static string m_WebHookURL    = "";
    private static string m_SecretKey     = "";
    private static string m_TextMsg       = "";
    private static string m_MarkdownTitle = "";
    private static string m_MarkdownMsg   = "";
    private static string m_Json          = "";
    private static string m_NReport       = "";
    private static string m_RobotName     = "";
    private static string[] m_Ats         = new string[0];
    private static Dictionary<string, string> m_ProdOwners = new Dictionary<string, string>();

    static void Main(string[] args) {

      try {
        string exePath = System.Reflection.Assembly.GetEntryAssembly().Location;
        string dir = new FileInfo(exePath).DirectoryName;
        IniLoad(Path.Combine(dir, "config.ini"));
      } catch (Exception ex) {
        Console.WriteLine("Default config file config.ini doesn't find");
      }

      int i = 0;
      while (i < args.Length) {

        switch (args[i].ToLower()) {
          case "-f":
            IniLoad( args[++i]);
            break;
          case "-u":
            m_WebHookURL = args[++i];
            break;
          case "-s":
            m_SecretKey = args[++i];
            break;
          case "-t":
            m_TextMsg = args[++i];
            break;
          case "-m":
            m_MarkdownMsg = args[++i];
            break;
          case "-mt":
            m_MarkdownTitle = args[++i];
            break;
          case "-j":
            m_Json = args[++i];
            break;
          case "-a":
            ParseAt(args[++i]);
            break;
          case "-n":
            m_NReport = args[++i];
            break;
          default:
            throw new Exception("Unknown arguments: \"" + args[i] + "\" " + args[++i]);
        }
        i++;
      }
      if (m_WebHookURL.Length == 0) throw new Exception("WebHookUrl is required");
      if (m_SecretKey.Length == 0) throw new Exception("SecretKey is required");
      if (m_TextMsg.Length == 0 && m_MarkdownMsg.Length == 0 && m_Json.Length == 0 && m_NReport.Length == 0) 
        throw new Exception("Markdown/Text/Json/NReport is required");

      DingTalkRobot robot = new DingTalkRobot(m_WebHookURL, m_SecretKey);
      if (m_NReport.Length > 0) {
        try {
          Dictionary<string, List<string>> reports = ParseNSoftReports(m_NReport);
          foreach (KeyValuePair<string, List<string>> report in reports) {
            string version = report.Key;
            List<string> msgs = report.Value;
            for (i = 0; i < msgs.Count; i += 2) {
              robot.SendText("[" + version + "] " + (m_RobotName.Length > 0 ? "[" + m_RobotName + "] " : "") +
                msgs[i] + "\n\n" + msgs[i + 1],
                new string[] { GetProdOwner(msgs[i]) }
              );
            }

          }
        } catch (Exception ex) {
          Console.WriteLine("Failed to send NReport: " + ex.Message);
        }
      } 
      if (m_Json.Length > 0) {
        try {
          robot.SendRaw(m_Json);
        } catch (Exception ex) {
          Console.WriteLine("Failed to send JSON: " + ex.Message);
        }
      } 
      if (m_MarkdownMsg.Length > 0) {
        try {
          robot.SendMarkDown(m_MarkdownTitle, m_MarkdownMsg, m_Ats);
        } catch (Exception ex) {
          Console.WriteLine("Failed to send Markdown: " + ex.Message);
        }
      }
      if (m_TextMsg.Length > 0) {
        try {
          robot.SendText(m_TextMsg, m_Ats);
        } catch (Exception ex) {
          Console.WriteLine("Failed to send Text: " + ex.Message);
        }
      }
    }

    public static void IniLoad(string file) {

      string[] lines = File.ReadAllLines(file);
      for (int i = 0; i < lines.Length; i++) {
        if (lines[i].Trim().Length == 0) continue;
        if (lines[i].StartsWith(";")) continue;

        if (lines[i].Contains("=")) {
          int pos = lines[i].IndexOf("=");
          string key = lines[i].Substring(0, pos).Trim();
          string val = lines[i].Substring(pos + 1).Trim();

          switch (key.ToUpper()) {
            case "WEBHOOKURL":
              m_WebHookURL = val;
              break;
            case "SECRETKEY":
              m_SecretKey = val;
              break;
            case "ROBOTNAME":
              m_RobotName = val;
              break;
            default:
              if (IsPhoneNum(key)) {
                String[] vals = val.Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
                for (int j = 0; j < vals.Length; j++) {
                  vals[j] = vals[j].Trim();
                  m_ProdOwners.Add(vals[j], key);
                }
              } else {
                Console.WriteLine("Ignore invalid config \"" + key + "\"");
              } 
              break;
          }
        }
      }
    }
    private static bool IsPhoneNum(String num) {
      num = num.Trim();
      if (num.Length != 11) return false;
      if (num[0] != '1') return false;
      for (int i = 0; i < num.Length; i++) {
        if (num[i] >= '0' && num[i] <= '9') continue;
        return false;
      }
      return true;
    }

    private static void ParseAt(String ats) {
      List<String> ret = new List<String>();
      String[] vals = ats.Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
      for (int j = 0; j < vals.Length; j++) {
        vals[j] = vals[j].Trim();
        if (IsPhoneNum(vals[j])) ret.Add(vals[j]);
      }
      m_Ats = ret.ToArray();
    }
    

    private static Dictionary<String, List<string>> ParseNSoftReports(String m) {
      Dictionary<String, List<string>> reports = new Dictionary<string, List<string>>();

      if (File.Exists(m)) m = File.ReadAllText(m);

      int pos = 0;
      pos = m.IndexOf("</h2>", pos); //<h2>nsoftware - v20</h2>
      while (pos > 0) {

        int startPos = pos;
        while (startPos - 1 > 0 && m[startPos - 1] != ' ') startPos--;
        string version = m.Substring(startPos, 3).Trim();

        List<string> ret = new List<string>();
        pos = m.IndexOf("<table", pos);
        int posEnd = m.IndexOf("</table>", pos);
        while (pos < posEnd) {
          int rowStart = m.IndexOf("<tr>", pos);
          int rowEnd = m.IndexOf("</tr>", pos);
          pos = rowEnd + "</tr>".Length;

          string rowStr = m.Substring(rowStart + "<tr>".Length, rowEnd - rowStart - "<tr>".Length);
          rowStr = rowStr.Replace("</td>", "");
          string[] cols = rowStr.Split(new string[] { "<td>" }, StringSplitOptions.RemoveEmptyEntries);
          if (cols.Length != 2 || cols[1].Equals("Success!")) continue;
          cols[0] = RemoveTag(cols[0]);
          cols[1] = RemoveTag(cols[1].Replace("<br>", "\n"));
          cols[1] = cols[1].Replace("\r\n", "\n");
          int pos2 = cols[1].IndexOf("Microsoft (R) Program");
          if (pos2 > 0) {
            cols[1] = cols[1].Substring(pos2);
          }
          while (cols[1].Contains("\n\n")) {
            cols[1] = cols[1].Replace("\n\n", "\n");
          }
          cols[1] = cols[1].Replace("\\", "\\\\").Replace("\"", "\\\"");
          ret.AddRange(cols);
        }

        if (ret.Count > 0) {
          reports.Add(version, ret);
        }

        pos = m.IndexOf("</h2>", pos); //<h2>nsoftware - v20</h2>
      }

      return reports;
    }

    private static string RemoveTag(string input) {
      StringBuilder sb = new StringBuilder();
      bool inTag = false;
      for (int i = 0; i < input.Length; i++) {
        char ch = input[i];
        if (ch == '<') { inTag = true; continue; } else if (inTag && ch == '>') { inTag = false; continue; }
        if (!inTag) sb.Append(ch);
      }
      return sb.ToString();
    }

    private static string GetProdOwner(string prod) {
      try {
        int pos = prod.IndexOf("[");
        string prodName = prod.Substring(0, pos).Trim();
        return m_ProdOwners[prodName];
      } catch (Exception ex) { }

      return "15399015948";
    }
  }
}
