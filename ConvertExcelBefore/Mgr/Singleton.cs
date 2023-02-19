using System;

namespace ConvertExcel
{
    public abstract class Singleton<T> where T : class, new()
    {
        private static readonly T instance = new T();

        protected Singleton()
        {
            if (instance != null)
                throw new Exception(instance+ "created more than once");
        }
        
        static Singleton()
        {
        }

        public static T Instance => instance;
    }
}