﻿using Newtonsoft.Json;
using Newtonsoft.Json.Serialization;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web.Http;

namespace Module4
{
	public static class WebApiConfig
	{
		public static void Register(HttpConfiguration config)
		{
			// Json settings
			config.Formatters.JsonFormatter.SerializerSettings.NullValueHandling = NullValueHandling.Ignore;
			config.Formatters.JsonFormatter.SerializerSettings.ContractResolver = new CamelCasePropertyNamesContractResolver();
			config.Formatters.JsonFormatter.SerializerSettings.Formatting = Formatting.Indented;
			JsonConvert.DefaultSettings = () => new JsonSerializerSettings()
			{
				ReferenceLoopHandling = ReferenceLoopHandling.Ignore,
				ContractResolver = new CamelCasePropertyNamesContractResolver(),
				Formatting = Newtonsoft.Json.Formatting.Indented,
				NullValueHandling = NullValueHandling.Ignore,
			};

			// Web API configuration and services

			// Web API routes
			config.MapHttpAttributeRoutes();

			config.Routes.MapHttpRoute(
					name: "DefaultApi",
					routeTemplate: "api/{controller}/{action}",
					defaults: new { id = RouteParameter.Optional }
			);
		}
	}
}
