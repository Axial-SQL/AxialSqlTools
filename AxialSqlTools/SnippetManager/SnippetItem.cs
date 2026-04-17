using System;

namespace AxialSqlTools
{
    public class SnippetItem
    {
        public string Id { get; set; }
        public string Prefix { get; set; }
        public string Description { get; set; }
        public string Body { get; set; }

        public SnippetItem()
        {
            Id = Guid.NewGuid().ToString();
            Prefix = string.Empty;
            Description = string.Empty;
            Body = string.Empty;
        }

        public SnippetItem(string prefix, string description, string body)
        {
            Id = Guid.NewGuid().ToString();
            Prefix = prefix ?? string.Empty;
            Description = description ?? string.Empty;
            Body = body ?? string.Empty;
        }
    }
}
