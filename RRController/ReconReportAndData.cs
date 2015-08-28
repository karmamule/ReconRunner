using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ReconRunner.Model;
using System.Data;

namespace ReconRunner.Controller
{
    class ReconReportAndData
    {
        ReconReport reconReport;
        DataTable firstQueryData;
        DataTable secondQueryData;

        public ReconReportAndData(ReconReport recon)
        {
            reconReport = recon;
        }

        public ReconReportAndData(ReconReport recon, DataTable query1Data, DataTable query2Data)
        {
            reconReport = recon;
            firstQueryData = query1Data;
            secondQueryData = query2Data;
        }

        public ReconReport ReconReport
        {
            get { return reconReport; }
            set { reconReport = value; }
        }

        public DataTable FirstQueryData
        {
            get { return firstQueryData; }
            set { firstQueryData = value; }
        }

        public DataTable SecondQueryData
        {
            get { return secondQueryData; }
            set { secondQueryData = value; }
        }
    }
}
