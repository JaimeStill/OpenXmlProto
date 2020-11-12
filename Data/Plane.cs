using System.Collections.Generic;

namespace OpenXmlProto.Data
{
    public class Plane
    {
        public string Name { get; set; }

        public IEnumerable<Person> People { get; set; }        
    }
}