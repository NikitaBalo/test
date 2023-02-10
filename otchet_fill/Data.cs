namespace otchet_fill
{
    public static class Data
    {
        private static Dictionary<int, object> DataGrid = new Dictionary<int, object>();

        public static void AddData(int i,object obj)
        {
            DataGrid.Add(i, obj);
        }
        public static object GetData(int i)
        {
            return DataGrid[i];
        }
        public static void ClearData()
        {
            DataGrid.Clear();
        }
    }
}
