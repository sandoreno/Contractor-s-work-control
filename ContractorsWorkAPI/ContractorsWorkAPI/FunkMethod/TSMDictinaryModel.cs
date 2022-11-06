namespace ContractorsWorkAPI.FunkMethod
{
    public class TSMDictinaryModel
    {
        public string Compound { get; set; }

        public string Addition { get; set; }

        public string Conversion_Index { get; set; }

        public string Compilation_Date { get; set; }

        //


        List<TSMDictinaryModel> TSN = new List<TSMDictinaryModel>();
        /// <summary>
        /// Извлекает часть строки используя 2 макроса
        /// </summary>
        /// <returns></returns>
        public string GetStringElement(string makros_start, string makros_end, ref string pasrse_string)
        {
            var makros_length_start = makros_start.Length;
            var makros_length_end = makros_end.Length;

            var index_start = pasrse_string.ToLower().IndexOf(makros_start.ToLower());
            var index_end = pasrse_string.ToLower().IndexOf(makros_end.ToLower());
            var full_index_start = index_start + makros_length_start;
            var full_index_end = index_end + makros_length_end;
            var str = pasrse_string.Substring(full_index_start, index_end - full_index_start);

            pasrse_string = pasrse_string.Remove(0, full_index_start);
            return str;
        }
    }
}
