using Microsoft.ApplicationInsights;
using Microsoft.ApplicationInsights.DataContracts;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Diagnostics;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;

namespace UCB_SPAMAddinWeb
{
    public class ConfigController : ApiController
    {
        private TelemetryClient telemetry = new TelemetryClient();

        public class DataResponse
        {
            public string Status { get; set; }
            public string Data { get; set; }
        }

        // GET api/<controller>
        public DataResponse GetConfig()
        {
            return new DataResponse()
            {
                Status = "Success",
                Data = "lucas.girardin@outlook.com"
            };
        }

        // GET api/<controller>/5
        public DataResponse Get(string id)
        {
            telemetry.TrackTrace("Get request", SeverityLevel.Information);
           if (id.ToLower().Equals("email"))
            {
                if (!ConfigurationManager.AppSettings["emailAddress"].Equals(string.Empty))
                {
                    return new DataResponse()
                    {
                        Status = "Success",
                        Data = ConfigurationManager.AppSettings["emailAddress"]
                    };
                }
                else
                {
                    telemetry.TrackTrace("No emailAddress application string", SeverityLevel.Error);
                    return new DataResponse()
                    {
                        Status = "Error",
                        Data = string.Format("No emailAddress application string")
                    };
                }
            }

            if (id.ToLower().Equals("bodyemail"))
            {
                if (!ConfigurationManager.AppSettings["bodyEmail"].Equals(string.Empty))
                {
                    return new DataResponse()
                    {
                        Status = "Success",
                        Data = ConfigurationManager.AppSettings["bodyEmail"]
                    };
                }
                else
                {
                    telemetry.TrackTrace("No bodyEmail application string", SeverityLevel.Error);
                    return new DataResponse()
                    {
                        Status = "Error",
                        Data = string.Format("No bodyEmail application string")
                    };
                }
            }

            telemetry.TrackTrace(string.Format("The ID {0} is not supported by Get function", id), SeverityLevel.Error);

            return new DataResponse()
            {
                Status = "Error",
                Data = string.Format("The ID {0} is not supported by Get function", id)
            };
        }

        // POST api/<controller>
        public void Post([FromBody]string value)
        {
        }

        // PUT api/<controller>/5
        public void Put(int id, [FromBody]string value)
        {
        }

        // DELETE api/<controller>/5
        public void Delete(int id)
        {
        }
    }
}