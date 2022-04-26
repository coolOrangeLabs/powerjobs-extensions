using System;
using System.Collections.Concurrent;
using System.Threading;

namespace coolOrange.AutoCADElectrical
{
    sealed class Proxy
    {
        private Thread _t;

        private Proxy()
        {
        }

        public BlockingCollection<Action> Collection { get; } = new BlockingCollection<Action>();
        public static Proxy Instance { get { return Nested.instance; } }

        private class Nested
        {
            // Explicit static constructor to tell C# compiler
            // not to mark type as beforefieldinit
            static Nested()
            {
            }

            internal static readonly Proxy instance = new Proxy();
        }

        public void Run()
        {
            if (_t == null)
            {
                _t = new Thread(() =>
                {
                    while (Collection.TryTake(out Action a, -1))
                    {
                        a.Invoke();
                    }
                });
                _t.SetApartmentState(ApartmentState.STA);
                _t.DisableComObjectEagerCleanup();
                _t.Start();
            }
        }
    }
}
