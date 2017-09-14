using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace ProvPersonalSite.REST
{
    class Program
    {
        static void Main(string[] args)
        {
        }

        public void CreatePersonalSiteUsingRest(string tenantAdminUrl, string userName, string password, string emailIDs)
        {
            //Validate Arguments
            Uri _uri = new Uri(tenantAdminUrl);
            string _emailData = emailIDs;
            if (_creds == null)
            {
                _creds = new SharePointOnlineCredentials(userName, Utilities.StringToSecure(password));
                string authCookie = _creds.GetAuthenticationCookie(_uri);
                _cookies = new CookieContainer();
                _cookies.Add(new Cookie("FedAuth", authCookie.TrimStart("SPOIDCRL=".ToCharArray()), "", _uri.Authority));

                Uri _apiContextInfo = new Uri(_uri, API_CONTEXTINFO);

                HttpWebRequest _endpointRequest = (HttpWebRequest)HttpWebRequest.Create(_apiContextInfo.OriginalString);
                _endpointRequest.Method = "POST";
                _endpointRequest.Accept = "text/xml;charset=utf-8";
                _endpointRequest.CookieContainer = _cookies;
                _endpointRequest.ContentLength = 0;
                _endpointRequest.Credentials = _creds;

                //We need to get the form digest since we are submitting or writing data via rest.
                //We need to make a call to _api/contextinnfo to retreive the digestvalue and attach the form digest value to all 
                //requests. http://msdn.microsoft.com/en-us/library/office/jj164022(v=office.15).aspx

                HttpWebResponse _endpointResponse = (HttpWebResponse)_endpointRequest.GetResponse();
                XmlNamespaceManager _xmlnspm = new XmlNamespaceManager(new NameTable());
                _xmlnspm.AddNamespace("d", "http://schemas.microsoft.com/ado/2007/08/dataservices");
                StreamReader _contextinfoReader = new StreamReader(_endpointResponse.GetResponseStream(), System.Text.Encoding.UTF8);
                var _formDigestXML = new XmlDocument();
                _formDigestXML.LoadXml(_contextinfoReader.ReadToEnd());
                var _formDigestNode = _formDigestXML.SelectSingleNode("//d:FormDigestValue", _xmlnspm);
                _formDigest = _formDigestNode.InnerXml;
            }

            Byte[] _payLoad = System.Text.Encoding.ASCII.GetBytes(_emailData);
            HttpWebRequest _spoMySiteRequest = (HttpWebRequest)HttpWebRequest.Create(new Uri(_uri, API_CREATEPERSONALSITE));
            _spoMySiteRequest.Method = "POST";
            _spoMySiteRequest.ContentType = "application/json;odata=verbose";
            _spoMySiteRequest.Accept = "application/json;odata=verbose";
            _spoMySiteRequest.Headers.Add("X-FORMS_BASED_AUTH_ACCEPTED", "f");
            _spoMySiteRequest.Headers.Add("X-RequestDigest", _formDigest);
            _spoMySiteRequest.CookieContainer = _cookies;
            _spoMySiteRequest.ContentLength = _emailData.Length;
            _spoMySiteRequest.Credentials = _creds;

            Stream _itemRequestStream = _spoMySiteRequest.GetRequestStream();
            _itemRequestStream.Write(_payLoad, 0, _payLoad.Length);
            _itemRequestStream.Close();

            HttpWebResponse _mysiteCreateResponse = (HttpWebResponse)_spoMySiteRequest.GetResponse();

        }

    }
}
