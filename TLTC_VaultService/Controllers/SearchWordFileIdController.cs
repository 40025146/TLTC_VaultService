using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;
using TLTC_VaultService.Lib;
namespace TLTC_VaultService.Controllers
{
    public class SearchWordFileIdController : ApiController
    {
        // GET api/values
        public IEnumerable<string> Get(string cond)
        {
            List<string> result = OpenXMLWord.ParseDirectoryFileContent(@"C:\Users\TLTC\Downloads",cond);

            return result;
        }
    }
}
