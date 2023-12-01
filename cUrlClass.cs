using System;
using System.IO;
using System.Net.Http;
using System.Threading.Tasks;
using static System.Net.Mime.MediaTypeNames;

namespace iPos4DS_DTTest
{
    public class cUrlClass : IDisposable
    {
        const string _cUrlProdSecureHTTP = "http://dashboard.fairbanc.app/api/documents";
        const string _cUrlTestSecureHTTP = "https://sandbox.fairbanc.app/api/documents";
        const string _cUrlProdUnsecureHTTP = "http://dashboard.fairbanc.app/api/documents";
        const string _cUrlTestUnsecureHTTP = "http://sandbox.fairbanc.app/api/documents";
        const string _prodToken = "2S0VtpYzETxDrL6WClmxXXnOcCkNbR5nUCCLak6EHmbPbSSsJiTFTPNZrXKk2S0VtpYzETxDrL6WClmx";
        const string _testToken = "KQtbMk32csiJvm8XDAx2KnRAdbtP3YVAnJpF8R5cb2bcBr8boT3dTvGc23c6fqk2NknbxpdarsdF3M4V";

        private string _httpresponses;

        private readonly string _token;
        private readonly string _urlAddress;
        private readonly string _targetFile;
        private readonly HttpClient? _httpClient;

        public string TargetFile => _targetFile;

        public string cUrlProdSecureHTTP => _urlAddress;

        public string Httpresponses => _httpresponses;

        public cUrlClass(char isSecureHttp = 'Y', char isSandbox = 'Y', string token = "", string fileandpathname = "")
        {
            _urlAddress = isSecureHttp  switch  
            { 
                'Y' => isSandbox == 'Y' ? _cUrlTestSecureHTTP : _cUrlProdSecureHTTP, 
                _ => isSandbox == 'Y' ? _cUrlTestUnsecureHTTP : _cUrlProdUnsecureHTTP
            };

            _token = isSandbox switch
            {
                'Y' => _testToken,
                _ => _prodToken
            };
            _targetFile = fileandpathname;
            _httpClient = new HttpClient();

        }

        public string SendRequest()
        {
            using var multipartFormDataContent = new MultipartFormDataContent();
            multipartFormDataContent.Add(new StringContent(_token), "api_token");
            multipartFormDataContent.Add(new ByteArrayContent(File.ReadAllBytes(_targetFile)), "file", Path.GetFileName(_targetFile));

            using var httpRequestMessage = new HttpRequestMessage(HttpMethod.Post, _urlAddress)
            {
                Content = multipartFormDataContent
            };

            using var httpResponseMessage =  _httpClient.Send(httpRequestMessage);
            Thread.Sleep(5000);
            httpResponseMessage.EnsureSuccessStatusCode();

            var strResponseBody = httpResponseMessage.Content.ReadAsStream();
            var array = httpResponseMessage.ToString().Split(':', ',');
            _httpresponses = array[1].Trim();
            return _httpresponses;
        }

        public void Dispose()
        {
            if (_httpClient is not null)
            {
                _httpClient.Dispose();
            }
        }
    }
}
