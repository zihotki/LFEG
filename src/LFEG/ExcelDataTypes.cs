namespace LFEG
{
    public static class ExcelDataTypes
    {
        public const string Boolean = "b";
        public const string Date = "d";
        public const string Number = "n";

        // actually it's a formula type but it's much more easy to use it instead of real string type
        // which is 's' because it require a custom value format
        public const string String = "inlineStr"; 
    }
}