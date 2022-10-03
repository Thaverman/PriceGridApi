using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace PriceGridApi.Models.PriceGridModels
{
    public class Competitor
    {
        public Competitor()
        {

        }

        public Competitor(int id, string url, string fullName, string status, string type)
        {
            this.id = id;
            this.url = url ?? throw new ArgumentNullException(nameof(url));
            this.fullName = fullName ?? throw new ArgumentNullException(nameof(fullName));
            this.status = status ?? throw new ArgumentNullException(nameof(status));
            this.type = type ?? throw new ArgumentNullException(nameof(type));
        }

        public int id { get; set; }
        public string url { get; set; }
        public string fullName { get; set; }
        public string status { get; set; }
        public string type { get; set; }
    }
}