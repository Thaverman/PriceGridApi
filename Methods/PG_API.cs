using PriceGridApi.Models.PriceGridModels;
using RestSharp;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Web;


namespace PriceGridApi.Methods
{
    public class PG_API
    {
        // IF Rest SHARP is updated this will break 
        public static void RequestCompetitorInformation(string PG_Store_Id) // Details about each compeditor 
        {
            string FullString = "https://api.pricegrid.com/ws/rest/competitor/" + PG_Store_Id;
            var client = new RestClient(FullString)
            {
                Timeout = -1
            };
            var request = new RestRequest(Method.GET);
            request.AddHeader("Authorization", "Basic VGhhdmVybWFuQHN0b3Jlc3VwcGx5LmNvbTohbG92ZW15akBiMTIz");
            request.AddHeader("Cookie", "JSESSIONID=8A365596CC70C94B5C1A36449E99BA3E");
            IRestResponse response = client.Execute(request);

            Debug.WriteLine(response.Content);
        }
        public static List<Competitor> RequestCompetitors() //Gathers all of the Compeditors on our account
        {
            List<Competitor> competitors = new List<Competitor>();
            var client = new RestClient("https://api.pricegrid.com/ws/rest/competitors?offset=:offset&max=:max")
            {
                Timeout = -1
            };
            var request = new RestRequest(Method.GET);
            request.AddHeader("Authorization", "Basic VEhhdmVybWFuQHN0b3Jlc3VwcGx5LmNvbTohbG92ZW15akBiMTIz");
            request.AddHeader("Cookie", "JSESSIONID=94D02AC0B5B944587E39E61F6A8407AD");
            IRestResponse response = client.Execute(request);

            var requestMatchesResponse = client.Deserialize<List<Competitor>>(response);
            foreach (Competitor item in requestMatchesResponse.Data)
            {
                competitors.Add(item);
            }

            return competitors;
        }
        public static List<Item> RequestItemInformation(string PG_Item_Id) // Gets matches to a an item based on the item ID given to the items by price grid
        {
            List<Item> items = new List<Item>();
            string FullString = "https://api.pricegrid.com/ws/rest/item/" + PG_Item_Id;
            var client = new RestClient(FullString);
            client.Timeout = -1;
            var request = new RestRequest(Method.GET);
            request.AddHeader("Authorization", "Basic VGhhdmVybWFuQHN0b3Jlc3VwcGx5LmNvbTohbG92ZW15akBiMTIz");
            request.AddHeader("Cookie", "JSESSIONID=8A365596CC70C94B5C1A36449E99BA3E");
            IRestResponse response = client.Execute<List<Item>>(request);
            var requestMatchesResponse = client.Deserialize<List<Item>>(response);
            foreach (Item item in requestMatchesResponse.Data)
            {
                items.Add(item);
            }
            return items;

        }
        public static List<Item> RequestItems() // Gathers all of the items we have on PriceGrid
        {
            List<Item> items = new List<Item>();
            var client = new RestClient("https://api.pricegrid.com/ws/rest/items?offset=:offset&max=:max");
            client.Timeout = -1;
            var request = new RestRequest(Method.GET);
            request.AddHeader("Authorization", "Basic VGhhdmVybWFuQHN0b3Jlc3VwcGx5LmNvbTohbG92ZW15akBiMTIz");
            request.AddHeader("Cookie", "JSESSIONID=94D02AC0B5B944587E39E61F6A8407AD");
            IRestResponse response = client.Execute(request);

            var requestMatchesResponse = client.Deserialize<List<Item>>(response);
            foreach (Item item in requestMatchesResponse.Data)
            {
                items.Add(item);
            }
            return items;
        }
        //public static List<RequestMatches> RequestCompetitorMatches(string CompetitorId) // Gathers all of the items we have on PriceGrid that match a specific competitor
        //{
        //    List<RequestMatches> requestMatches = new List<RequestMatches>();
        //    string FullString = "https://api.pricegrid.com/ws/rest/competitor/" + CompetitorId + "/matches?offset=:offset&max=:max";
        //    var client = new RestClient(FullString);
        //    client.Timeout = -1;
        //    var request = new RestRequest(Method.GET);
        //    request.AddHeader("Authorization", "Basic VEhhdmVybWFuQHN0b3Jlc3VwcGx5LmNvbTohbG92ZW15akBiMTIz");
        //    request.AddHeader("Cookie", "JSESSIONID=94D02AC0B5B944587E39E61F6A8407AD");
        //    IRestResponse response = client.Execute<List<Competitor>>(request);

        //    var requestMatchesResponse = client.Deserialize<List<RequestMatches>>(response);
        //    foreach (RequestMatches item in requestMatchesResponse.Data)
        //    {
        //        requestMatches.Add(item);
        //    }
        //    return requestMatches;
        //}
        public static List<RequestMatches> RequestCompetitorMatches(string CompetitorId, int offset) // Gathers all of the items we have on PriceGrid that match a specific competitor using an offset if needed
        {
            List<RequestMatches> requestMatches = new List<RequestMatches>();
            string FullString = "https://api.pricegrid.com/ws/rest/competitor/" + CompetitorId + "/matches?offset=" + offset + "&max=:max";
            var client = new RestClient(FullString);
            client.Timeout = -1;
            var request = new RestRequest(Method.GET);
            request.AddHeader("Authorization", "Basic VEhhdmVybWFuQHN0b3Jlc3VwcGx5LmNvbTohbG92ZW15akBiMTIz");
            request.AddHeader("Cookie", "JSESSIONID=94D02AC0B5B944587E39E61F6A8407AD");
            IRestResponse response = client.Execute<List<Competitor>>(request);

            var requestMatchesResponse = client.Deserialize<List<RequestMatches>>(response);
            foreach (RequestMatches item in requestMatchesResponse.Data)
            {
                requestMatches.Add(item);
            }
            return requestMatches;
        }
        public static List<RequestMatches> RequestItemIdMatches(string PG_Item_Id) // returns a list of items that match our accounts item from the itemId 
        {
            List<RequestMatches> requestMatches = new List<RequestMatches>();

            string FullString = "https://api.pricegrid.com/ws/rest/item/" + PG_Item_Id + "/matches?offset=&max=:max";
            var client = new RestClient(FullString);
            client.Timeout = -1;
            var request = new RestRequest(Method.GET);
            request.AddHeader("Authorization", "Basic VEhhdmVybWFuQHN0b3Jlc3VwcGx5LmNvbTohbG92ZW15akBiMTIz");
            request.AddHeader("Cookie", "JSESSIONID=8A365596CC70C94B5C1A36449E99BA3E");
            IRestResponse response = client.Execute<List<RequestMatches>>(request);
            //Debug.WriteLine(response.Content);
            var requestMatchesResponse = client.Deserialize<List<RequestMatches>>(response);

            foreach (RequestMatches item in requestMatchesResponse.Data)
            {
                requestMatches.Add(item);
            }
            return requestMatches;
        }
        public static List<RequestMatches> RequestSkuMatches(string SKU) // Request Matches
        {
            List<RequestMatches> requestMatches = new List<RequestMatches>();

            string FullString = "https://api.pricegrid.com/ws/rest/matches?query=" + SKU + "&offset=:0&max=:1000";
            var client = new RestClient(FullString);
            client.Timeout = -1;
            var request = new RestRequest(Method.GET);
            request.AddHeader("Authorization", "Basic VEhhdmVybWFuQHN0b3Jlc3VwcGx5LmNvbTohbG92ZW15akBiMTIz");
            request.AddHeader("Cookie", "JSESSIONID=94D02AC0B5B944587E39E61F6A8407AD");
            request.AlwaysMultipartFormData = true;
            IRestResponse response = client.Execute<List<RequestMatches>>(request);

            var requestMatchesResponse = client.Deserialize<List<RequestMatches>>(response);
            foreach (RequestMatches item in requestMatchesResponse.Data)
            {
                requestMatches.Add(item);
            }
            return requestMatches;
        }



    }
}