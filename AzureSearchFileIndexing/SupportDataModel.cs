using Azure.Search.Documents.Indexes;
using System;

using System.Collections.Generic;

using System.Linq;
using System.Text;

using System.Threading.Tasks;



using System;

using System.Net;




namespace AzureSearchFileIndexing
{
    using System.Text.Json.Serialization;
    using Azure.Search.Documents.Indexes;

    using Azure.Search.Documents.Indexes.Models;


    public partial class SupportDataModel
    {

        [SimpleField(IsKey = true, IsFilterable = true)]
        public string TicketID { get; set; }

        [SimpleField(IsFilterable = true, IsSortable = true, IsFacetable = true)]

        public string TicketType { get; set; }

        [SimpleField(IsFilterable = true, IsSortable = true, IsFacetable = true)]
        public string TicketSubject {  get; set;     }

        [SimpleField(IsFilterable = true, IsSortable = true, IsFacetable = true)]
        public string TicketDescription { get; set; }

        [SimpleField(IsFilterable = true, IsSortable = true, IsFacetable = true)]
        public string TicketStatus { get; set; }

        [SimpleField(IsFilterable = true, IsSortable = true, IsFacetable = true)]
        public string TicketPriority { get; set; }

        [SimpleField(IsFilterable = true, IsSortable = true, IsFacetable = true)]
        public string TicketChannel { get; set; }


    }
}

