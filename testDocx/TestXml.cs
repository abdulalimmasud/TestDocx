using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace testDocx
{
    public static class TestXml
    {
        public static Stream GetStream(string path)
        {
            using(Stream stream = System.IO.File.Open(path, FileMode.OpenOrCreate))
            {
                return stream;
            }
        }
    }
}
