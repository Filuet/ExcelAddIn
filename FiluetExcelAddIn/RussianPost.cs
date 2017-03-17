using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Serialization;
using System.Xml;
using System.Net;
using System.IO;


public static class RussianPost
{
	private const string login = "oooFiluetPC";
	private const string password = "UtdWluQ3xI";
	private const string server = @"http://voh.russianpost.ru:8080/niips-operationhistory-web-ml/OperationHistory?wsdl";


	public static bool GetStatus(string barcode, ref PostResponse resp)
	{
		bool r = true;
		if (string.IsNullOrEmpty(barcode))
			return true;
		XmlDocument requestXml = CreateRequest(barcode);
		resp = new PostResponse();
		HttpWebRequest request = CreateWebRequest();
		try
		{
			using (Stream stream = request.GetRequestStream())
			{
				requestXml.Save(stream);
			}

			using (WebResponse response = request.GetResponse())
			{
				using (StreamReader rd = new StreamReader(response.GetResponseStream()))
				{
					XmlSerializer deserializer = new XmlSerializer(typeof(PostResponse));
					resp = (PostResponse)deserializer.Deserialize(rd);
				}
			}
		}
		catch
		{
			r = false;
		}
		return r;
	}

	private static XmlDocument CreateRequest(string barcode)
	{
		XmlDocument res = new XmlDocument();
		string t = string.Format(@"<?xml version=""1.0"" encoding=""utf-8""?><soapenv:Envelope xmlns:soapenv=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:data=""http://russianpost.org/operationhistory/data""><soapenv:Header><data:AuthorizationHeader soapenv:mustUnderstand=""1""><data:login>{0}</data:login><data:password>{1}</data:password></data:AuthorizationHeader></soapenv:Header><soapenv:Body><data:OperationHistoryRequest><data:Barcode>{2}</data:Barcode><data:MessageType>0</data:MessageType><!--Optional:--><data:Language>RUS</data:Language></data:OperationHistoryRequest></soapenv:Body></soapenv:Envelope>", login, password, barcode);
		res.LoadXml(t);
		return res;
	}
	private static HttpWebRequest CreateWebRequest()
	{
		HttpWebRequest webRequest = (HttpWebRequest)WebRequest.Create(server);
		webRequest.Headers.Add(@"SOAP:Action");
		webRequest.ContentType = "text/xml";
		webRequest.Accept = "text/xml";
		webRequest.Method = "POST";
		return webRequest;
	}
}

[XmlRoot(ElementName = "Envelope", IsNullable = false, Namespace = "http://schemas.xmlsoap.org/soap/envelope/")]
public class PostResponse
{
	public clBody Body { get; set; }

	public class clBody
	{
		[XmlElement(Namespace = "http://russianpost.org/operationhistory/data")]
		public clOperationHistoryData OperationHistoryData { get; set; }

		public class clOperationHistoryData
		{
			[XmlElement(Namespace = "http://russianpost.org/operationhistory/data")]
			public List<clHistoryRecord> historyRecord { get; set; }

			public class clHistoryRecord
			{
				[XmlElement(Namespace = "http://russianpost.org/operationhistory/data")]
				public clFinanceParameters FinanceParameters { get; set; }

				[XmlElement(Namespace = "http://russianpost.org/operationhistory/data")]
				public clOperationParameters OperationParameters { get; set; }

				[XmlElement(Namespace = "http://russianpost.org/operationhistory/data")]
				public clAddressParameters AddressParameters { get; set; }

				[XmlElement(Namespace = "http://russianpost.org/operationhistory/data")]
				public clItemParameters ItemParameters { get; set; }

				public class clFinanceParameters
				{
					public decimal Payment { get; set; }
					public decimal Value { get; set; }
					public decimal MassRate { get; set; }
					public decimal InsrRate { get; set; }
					public decimal AirRate { get; set; }
					public decimal Rate { get; set; }
				}
				public class clOperationParameters
				{
					public DateTime OperDate { get; set; }
					public clOperType OperType { get; set; }
					public clOperAttr OperAttr { get; set; }

					public class clOperType
					{
						public int Id { get; set; }
						public string Name { get; set; }
					}
					public class clOperAttr
					{
						public int Id { get; set; }
						public string Name { get; set; }
					}
				}
				public class clAddressParameters
				{
					public clAddress DestinationAddress { get; set; }
					public clAddress OperationAddress { get; set; }

					public class clAddress
					{
						public int Index { get; set; }
						public string Description { get; set; }
					}
				}
				public class clItemParameters
				{
					public decimal Mass { get; set; }
				}
			}
		}
	}
}
