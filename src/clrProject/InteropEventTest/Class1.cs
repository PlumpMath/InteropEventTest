using System.Runtime.InteropServices;
using System.Threading.Tasks;

namespace InteropEventTest
{
    [Guid("E1BC643E-0CCF-4A91-8499-71BC48CAC01D")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    [ComVisible(true)]
    public interface ITheEvents
    {
        void OnHappened(string theMessage);
    }

    [Guid("77F1EEBA-A952-4995-9384-7228F6182C32")]
    [ComVisible(true)]
    public interface IInteropConnection
    {
        void DoEvent(string theMessage);
    }

    [Guid("2EE25BBD-1849-4CA8-8369-D65BF47886A5")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComSourceInterfaces(typeof(ITheEvents))]
    [ComVisible(true)]
    public class InteropConnection : IInteropConnection
    {
        [ComVisible(false)]
        public delegate void Happened(string theMessage);
        public event Happened OnHappened;


        public void DoEvent(string theMessage)
        {

            if (OnHappened != null)
            {
                Task.Factory.StartNew(() => OnHappened(theMessage));
            }
        }
    }
}
