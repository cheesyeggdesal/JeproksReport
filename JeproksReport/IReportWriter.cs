using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace JeproksReport
{
    public interface IReportWriter
    {
        void Write(ReportTemplate reportTemplate, Stream stream, Dictionary<string, IEnumerable<object>> datasource = null);
    }
}
