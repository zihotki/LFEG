using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LFEG.Sample
{
    class Program
    {
        private static void Main(string[] args)
        {
            var generator = new ExcelFileGenerator(new ExcelColumnFactory(new IExcelColumnDataInitializerVisitor[]
            {
                new BooleanYesNoDataInitializerVisitor(),
                new DateTimeDataInitializerVisitor(),
            }
                .Concat(ExcelColumnFactory.DefaultDataInitializerVisitors)
                .ToArray()));

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
            for (var i = 0; i < 5000; i++)
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
