using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReleaseNotesBusinessLogic
{
    public class Issue
    {
        public int Id { get; set; }
        public string Key { get; set; }
        public string Description { get; set; }
        public string Name { get; set; }
        public Project ProjectID { get; set; }
        public List<string> Labels { get; set; }
        public DateTime? ResolutionDate { get; set; }
        public DateTime? CreationDate { get; set; }
        public string Asignee { get; set; }
        public string Creator { get; set; }
        public string Status { get; set; }

    }
}
