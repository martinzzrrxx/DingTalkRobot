using System;
using System.IO;
using System.Net;
using System.Security.Cryptography;

namespace DingTalkRobot {

  // ref: https://ding-doc.dingtalk.com/doc#/serverapi2/qf2nxq
  
  public class DingTalkRobot {
    private static readonly DateTime EPOCH = new DateTime(1970, 1, 1, 0, 0, 0, DateTimeKind.Utc);
    private const string SUCCESS_RESPONSE = "{\"errcode\":0,\"errmsg\":\"ok\"}";

    private string m_WebHookURL = "";
    private string m_SecretKey = "";

    public DingTalkRobot(String url, String secrete) {
      m_WebHookURL = url;
      m_SecretKey = secrete;
    }
    public void SendText(String text, params String[] ats) {
      String arg0 = "";
      String arg2 = "";
      for (int i = 0; i < ats.Length; i++) {
        if (i > 0) { arg0 += " "; arg2 += ","; }
        arg0 += "@" + ats[i];
        arg2 += ats[i];
      }

      SendRaw(String.Format(@"{{
          ""msgtype"": ""text"", 
          ""text"": {{
              ""content"": ""{0} {1}""
          }},
          ""at"": {{
              ""atMobiles"": [{2}], 
              ""isAtAll"": {3}
          }}
        }}", arg0, text, arg2, arg0.Length > 0 ? "false" : "true"));

    }
    public void SendMarkDown(String title, string text, params String[] ats) {
      String arg1 = "";
      String arg3 = "";
      for (int i = 0; i < ats.Length; i++) {
        if (i > 0) { arg1 += ","; arg3 += ","; }
        arg1 += "@" + ats[i];
        arg3 += ats[i];
      }

      SendRaw(String.Format(@"{{
            ""msgtype"": ""markdown"",
            ""markdown"": {{
                ""title"":""{0}"",
                ""text"": ""{1} {2}""
            }},
            ""at"": {{
                ""atMobiles"": [{3}],
                ""isAtAll"": {4}
            }}
        }}", title, arg1, text, arg3, arg1.Length > 0 ? "false" : "true"));
    }
    public void SendRaw(String json) {
      json = json.Replace("\r\n", "\n");
      long timestamp = (long)(DateTime.UtcNow - EPOCH).TotalMilliseconds;
      string signature = "";
      using (HMACSHA256 hmac = new HMACSHA256(System.Text.Encoding.Default.GetBytes(m_SecretKey))) {
        byte[] hashValue = hmac.ComputeHash(System.Text.Encoding.Default.GetBytes(timestamp + "\n" + m_SecretKey));

        signature = Uri.EscapeDataString(Convert.ToBase64String(hashValue));
      }

      WebRequest request = WebRequest.Create(m_WebHookURL + "&timestamp=" + timestamp + "&sign=" + signature);
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

  }
}
