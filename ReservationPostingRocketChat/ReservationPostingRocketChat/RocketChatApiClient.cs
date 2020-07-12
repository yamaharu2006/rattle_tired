using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net.Http;
using System.Windows.Threading;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace ReservationPostingRocketChat
{
    static class RocketChatApiClient
    {
        static HttpClient client = new HttpClient();

        static public string UserID;
        static public string AuthToken;
        static public string HostName;

        static object lockRocketChatApiInfo = new object();

        public static void Init(string userId, string authToken, string hostName)
        {
            lock (lockRocketChatApiInfo)
            {
                UserID = userId;
                AuthToken = authToken;
                HostName = hostName;
            }
        }

        public static string GetRoomInfo(string roomName)
        {
            var request = new HttpRequestMessage(HttpMethod.Get, @"/api/v1/rooms.info");
            AddAuthHeader(ref request);

            request.Content = new FormUrlEncodedContent(new Dictionary<string, string>()
            {
                { "roomName", roomName },
            });

            var response = client.SendAsync(request).Result;
            var responseBody = response.Content.ReadAsStringAsync().Result;



            return "";
        }

        public static void PostMessage(string roomId, string context)
        {
            var request = new HttpRequestMessage(HttpMethod.Post, @"/api/v1/chat.postMessage");
            AddAuthHeader(ref request);

            request.Content = new FormUrlEncodedContent(new Dictionary<string, string>()
            {
                { "roomId", roomId },
                { "text", context },
            });
            
            var response = client.SendAsync(request).Result;
            var responseBody = response.Content.ReadAsStringAsync().Result;
        }

        static void AddAuthHeader(ref HttpRequestMessage request)
        {
            lock (lockRocketChatApiInfo)
            {
                request.Headers.Add(@"X-Auth-Token", UserID);
                request.Headers.Add(@"X-User-Id", AuthToken);
            }
        }

    }
}
