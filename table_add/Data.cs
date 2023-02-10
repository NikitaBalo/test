using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace table_add
{
    public static class Data
    {
        private static Dictionary<int,object> _data = new Dictionary<int,object>();
        private static List<object> _list = new List<object>();
        
        public static void Add(int i, object obj)
        {
            _data.Add(i,obj);
        }
        public static void Set(int i,object obj)
        {
            _data[i] = obj;
        }
        public static object Get(int i)
        {
            return _data[i];
        }
        public static void Clear()
        {
            _data.Clear();
        }
        public static void Add1(object obj)
        {
            _list.Add(obj);
        }
        public static object Get1(int i)
        {
            return _list[i];
        }
    }
}
