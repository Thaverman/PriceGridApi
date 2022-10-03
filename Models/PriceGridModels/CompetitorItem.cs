using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace PriceGridApi.Models.PriceGridModels
{
    public class CompetitorItem
    {
        public CompetitorItem()
        {
        }

        public int Id { get; set; }
        public Competitor Competitor { get; set; }
        public List<Object> Options { get; set; }

        public string Name { get; set; }
        public string UpdateStatus { get; set; }
        public decimal Price { get; set; }
    }
}