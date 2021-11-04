using RestSharp;
using System.IO;
using Newtonsoft.Json;
using System.Collections.Generic;

namespace ChampionsLeague2021
{
    public class RestHelper
    {
        public RestClient _restClient;

        public RestRequest _restRequest;

        public string _baseUrl = "";

        public RestClient SetRestClient(string resourceUrl)
        {
            var fullPath = Path.Combine(_baseUrl, resourceUrl);
            _restClient = new RestClient(fullPath);
            return _restClient;

        }

        public RestRequest CreateGetRequest(Method method)
        {
            _restRequest = new RestRequest(method);
            _restRequest.AddHeader("Accept", "application/json");
            //_restRequest.AddParameter()
            return _restRequest;
        }

        public IRestResponse GetResponse(RestClient restClient, RestRequest restRequest)
        {
            return restClient.Execute(restRequest);
        }

        public T Content<T>(IRestResponse respnse)
        {
            var content = respnse.Content;
            var deserializeObject = JsonConvert.DeserializeObject<List<T>>(content);

            return deserializeObject[0];
        }

        public List<T> Content2<T>(IRestResponse respnse)
        {
            var content = respnse.Content;
            var deserializeObject = JsonConvert.DeserializeObject<List<T>>(content);

            return deserializeObject;
        }
    }
}
