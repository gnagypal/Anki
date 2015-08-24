namespace AnkiTTSGenerator
{
    using System;
    using System.Collections.Generic;

    using System.IO;
    using System.Speech.Synthesis;
    using System.Xml.Linq;

    using Fclp;

    using NAudio.Lame;
    using NAudio.Wave;

    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                Console.WriteLine("AnkiTTSGenerator started...\n");

                var p = new Program();
                p.RunProgram(args);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
            finally
            {
                Console.WriteLine("\nAnkiTTSGenerator finished.");
            }
        }

        private void RunProgram(string[] args)
        {
            var p = new FluentCommandLineParser<Configuration>();
            var result = this.ParseCommandLine(args, p);

            if (result.HasErrors == false)
            {
                switch (p.Object.Function)
                {
                    case ProgramFunction.GenerateTTS:
                        this.GenerateTTS(p.Object);
                        break;

                    case ProgramFunction.ListInstalledVoices:
                        this.ListInstalledVoices();
                        break;
                }
            }
            else if (result.HelpCalled)
            {
                Console.WriteLine("heeeelp");
            }
            else
            {
                Console.WriteLine(result.ErrorText);
                p.HelpOption.ShowHelp(p.Options);
            }
        }

        private ICommandLineParserResult ParseCommandLine(string[] args, FluentCommandLineParser<Configuration> p)
        {
            p.SetupHelp("?", "help").Callback(text => Console.WriteLine(text));

            p.Setup(options => options.DestinationDirectory).As('d', "destinationDirectory").WithDescription(" ");

            p.Setup(options => options.Function)
                .As('f', "function")
                .SetDefault(ProgramFunction.ListInstalledVoices)
                .WithDescription("ListInstalledVoices, GenerateTTS ");

            p.Setup(options => options.IDElementName).As('i', "IDElementName").SetDefault("ID").WithDescription(" ");
            p.Setup(options => options.MP3BitRate).As('b', "bitRate").SetDefault(64).WithDescription(" ");

            p.Setup(options => options.RowElementName).As('r', "rowElementName").SetDefault("AnkiCardRow").WithDescription(" ");

            p.Setup(options => options.SelectedVoice).As('v', "voice").WithDescription(" ");
            p.Setup(options => options.SourceXmlFileName).As('s', "sourceXmlFileName").WithDescription(" ");
            p.Setup(options => options.Mp3FilePrefix).As('m', "mp3FilePrefix").WithDescription(" ");
            p.Setup(options => options.TextElementNames).As('t', "textElementNames").WithDescription(" ");

            var result = p.Parse(args);
            return result;
        }

        private void GenerateTTS(Configuration cfg)
        {
            using (var synthesizer = new SpeechSynthesizer())
            {
                this.InitializeSpeechSynthesizer(cfg, synthesizer);

                var xmlDoc = this.LoadXmlDocument(cfg);
                var rows = this.LoadRows(cfg, xmlDoc);

                foreach (var row in rows)
                {
                    string id = this.GetElementValue(cfg, row, cfg.IDElementName);
                    Console.WriteLine(id);
                    
                    byte i = 0;
                    foreach (var textElementName in cfg.TextElementNames)
                    {
                        string textElement = this.GetElementValue(cfg, row, textElementName);
                        char postFix = (char)((byte)'A' + i);
                        string filename = cfg.Mp3FilePrefix + id + postFix + ".mp3";
                        string fullFilename = Path.Combine(cfg.DestinationDirectory, filename);
                        this.SpeakToMp3(cfg, synthesizer, fullFilename, textElement);
                        i++;
                    }
                }
            }
        }

        private void InitializeSpeechSynthesizer(Configuration cfg, SpeechSynthesizer synthesizer)
        {
            synthesizer.SelectVoice(cfg.SelectedVoice);

            //set some settings
            synthesizer.Volume = 100;
            synthesizer.Rate = 0; //medium
        }

        private IEnumerable<XElement> LoadRows(Configuration cfg, XElement xmlDoc)
        {
            var rows = xmlDoc.Descendants(cfg.RowElementName);
            if (rows == null)
            {
                throw new ApplicationException(
                    string.Format(
                        "Can't find '{0}' nodes in the loaded XML file '{1}'.",
                        cfg.RowElementName,
                        cfg.SourceXmlFileName));
            }
            return rows;
        }

        private XElement LoadXmlDocument(Configuration cfg)
        {
            XElement xmlDoc;
            try
            {
                xmlDoc = XElement.Load(cfg.SourceXmlFileName);
            }
            catch (Exception ex)
            {
                throw new ApplicationException(string.Format("Couldn't load the XML file '{0}'.", cfg.SourceXmlFileName), ex);
            }
            return xmlDoc;
        }

        private string GetElementValue(Configuration cfg, XElement row, string elementName)
        {
            var element = row.Element(elementName);
            if (element == null)
            {
                throw new ApplicationException(
                    string.Format(
                        "Can't find '{0}' node in the loaded XML file '{1}'.",
                        elementName,
                        cfg.SourceXmlFileName));
            }

            return element.Value;            
        }

        private void ListInstalledVoices()
        {
            using (SpeechSynthesizer reader = new SpeechSynthesizer())
            {
                var voices = reader.GetInstalledVoices();

                Console.WriteLine("The installed Windows TTS voices:");

                foreach (var v in voices)
                {
                    Console.WriteLine(v.VoiceInfo.Name);
                }
            }
        }

        private void SpeakToMp3(Configuration cfg, SpeechSynthesizer reader, string fileName, string speakText)
        {
            //save to memory stream
            MemoryStream ms = new MemoryStream();
            reader.SetOutputToWaveStream(ms);

            //do speaking
            reader.Speak(speakText);

            //now convert to mp3 using LameEncoder or shell out to audiograbber
            ConvertWavStreamToMp3File(cfg, ref ms, fileName);
        }

        public static void ConvertWavStreamToMp3File(Configuration cfg, ref MemoryStream ms, string savetofilename)
        {
            //rewind to beginning of stream
            ms.Seek(0, SeekOrigin.Begin);

            using (var retMs = new MemoryStream())
            {
                using (var rdr = new WaveFileReader(ms))
                {
                    using (var wtr = new LameMP3FileWriter(savetofilename, rdr.WaveFormat, cfg.MP3BitRate))
                    {
                        rdr.CopyTo(wtr);
                        wtr.Close();
                    }

                    rdr.Close();

                    Console.WriteLine(savetofilename + " ok.");
                }
            }
        }
    }
}
