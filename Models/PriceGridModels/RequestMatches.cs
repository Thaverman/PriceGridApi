using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace PriceGridApi.Models.PriceGridModels
{
    public class RequestMatches
    {
        public int Id { get; set; }
        public string VerificationStatus { get; set; }

        public Item Item { get; set; }

        public CompetitorItem CompetitorItem { get; set; }
        public Competitor Competitor { get; set; }

        public RequestMatches()
        {

        }
    }
}