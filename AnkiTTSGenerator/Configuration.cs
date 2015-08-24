using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AnkiTTSGenerator
{
    enum ProgramFunction
    {
        ListInstalledVoices,
        GenerateTTS
    }

    class Configuration
    {
        public ProgramFunction Function { get; set; }

        public string SelectedVoice { get; set; }

        public int MP3BitRate { get; set; }

        public string SourceXmlFileName { get; set; }
        
        public string RowElementName { get; set; }

        public string IDElementName { get; set; }

        public List<string> TextElementNames { get; set; }

        public string Mp3FilePrefix { get; set; }

        public string DestinationDirectory { get; set; }
    }
}
