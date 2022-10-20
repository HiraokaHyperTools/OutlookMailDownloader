using LiteDB;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OutlookMailDownloader.Models
{
    public class VisitedMessage
    {
        [BsonId]
        public string? Id { get; set; }
    }
}
