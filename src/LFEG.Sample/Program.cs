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
            var builder = ExcelFileGeneratorSettings.Create()
                .AddDefaultColumnVisitors()
                .AddColumnVisitor<BooleanYesNoDataProviderVisitor>()
                .AddColumnVisitor<DateTimeDataProviderVisitor>();

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
            for (var i = 0; i < 1000000; i++)
            {
                yield return new Model
                {
                    Id = new Guid(),
                    Title = "Title " + i,
                    IsEnabled = i > 2000,
                    Type = i > 1500 ? ModelType.AnotherOne : ModelType.One,
                    UpdatedDate = DateTime.Now.AddDays(i),
                    Width = i*i
                };
            }
        } 
    }

    class Model
    {
        public Guid Id { get; set; }
        public string Title { get; set; }
        public bool IsEnabled { get; set; }
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
