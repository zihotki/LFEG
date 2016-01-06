using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Linq;

namespace LFEG.Sample
{
    class Program
    {
        private static void Main(string[] args)
        {
            // we can reuse settings so it's a good candidate to be put into DI container
            var builder = ExcelFileGeneratorSettings.Create();

            var generator = builder.CreateGenerator();

            var data = GenerateData().ToArray();

            using (var stream = File.Create("c:\\output.xlsx"))
            {
                var stopwatch = Stopwatch.StartNew();
                generator.GenerateFile(data, stream);
                stopwatch.Stop();

                Console.WriteLine("Elapsed " + stopwatch.ElapsedMilliseconds + " ms.");

                stream.Flush();
            }

            Console.ReadKey();
        }


        private static IEnumerable<Model> GenerateData()
        {
            for (var i = 0; i < 999; i++)
            {
                yield return new Model
                {
                    Id = Guid.NewGuid(),
                    TitleInline = "Title " + i%3,
                    TitleShared = "Title " + i%3,
                    IsEnabled = i%2 == 0,
                    IsEnabled2 = i % 2 == 0,
                    Type = i%2 == 1 ? ModelType.AnotherOne : ModelType.One,
                    UpdatedDate = DateTime.Now.AddDays(i),
                    Width = i*i,
                };
            }
        } 
    }

    class Model
    {
        [IgnoreExcelExport]
        public Guid Id { get; set; }

        [ExcelExport(Caption = "Title hello", Width = 40)]
        public string TitleInline { get; set; }

        [ExcelExport(Intern = true)]
        public string TitleShared { get; set; }

        [ExcelExport(DataFormat = "\"Yes\";;\"No\"")]
        public bool IsEnabled { get; set; }
        
        public bool IsEnabled2 { get; set; }
        public int Width { get; set; }
        public ModelType Type { get; set; }
        public DateTime UpdatedDate { get; set; }

    }

    enum ModelType
    {
        [Description("One")]
        One,
        [Description("Two")]
        Two,
        [Description("Another One")]
        AnotherOne
    }
}
