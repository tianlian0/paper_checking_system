using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace paper_checking
{
    static class RunningEnv
    {
        struct running_data
        {
            public int check_way { get; }
            public int check_threshold { get; }
            public string report_path { get; }
            public bool st_table { get; }
            public bool recover { get; }
            public string paper_to_add_path { get; }
            public int check_t_cnt { get; }
            public int convert_t_cnt { get; }
            public bool suport_pdf { get; }
            public bool suport_doc { get; }
            public bool suport_docx { get; }
            public bool suport_txt { get; }
        }


    }
}
