using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReservationPostingRocketChat
{
    class PostingEvent
    {
        public string RoomName { get; set; }
        public DateTime PostingTime { get; set; }
        public string StrPostingTime
        {
            get { return PostingTime.ToString(); }
        }
        public string PostingContext { get; set; }

        public PostingEvent(string roomName, string postingTime, string postingContext)
        {
            this.RoomName = roomName;
            this.PostingTime = DateTime.Parse(postingTime);
            this.PostingContext = postingContext;
        }

        public bool Exec()
        {
            var roomId = RocketChatApiClient.GetRoomInfo(RoomName);
            RocketChatApiClient.PostMessage(roomId, PostingContext);
            return true;
        }
    }
}
