using System;
using System.Runtime.InteropServices;

namespace PirnterUtility.Tool
{
    class Kernel32
    {

        //获取设备文件（获取句柄）
        [DllImport("kernel32.dll", SetLastError = true)]
        //根据要求可在下面设定参数，具体参考参数说明：https://docs.microsoft.com/zh-cn/windows/desktop/api/fileapi/nf-fileapi-createfilea
        public static extern int CreateFile
            (
             string lpFileName,                             // file name 文件名
             uint dwDesiredAccess,                        // access mode 访问模式
             uint dwShareMode,                            // share mode 共享模式
             uint lpSecurityAttributes,                   // SD 安全属性
             uint dwCreationDisposition,                  // how to create 如何创建
             uint dwFlagsAndAttributes,                   // file attributes 文件属性
             uint hTemplateFile                           // handle to template file 模板文件的句柄
            );


        [DllImport("Kernel32.dll", SetLastError = true)]  //接收函数DLL
        public static extern bool ReadFile
            (
                IntPtr hFile,
                byte[] lpBuffer,
                uint nNumberOfBytesToRead,
                ref uint lpNumberOfBytesRead,
                ref OVERLAPPED lpOverlapped
            );

        [DllImport("kernel32.dll", SetLastError = true)] //发送数据DLL
        public static extern Boolean WriteFile
            (
                IntPtr hFile,
                byte[] lpBuffer,
                uint nNumberOfBytesToWrite,
                ref uint nNumberOfBytesWrite,
                ref OVERLAPPED lpOverlapped
            // IntPtr lpOverlapped
            );

        //关闭访问设备句柄，结束进程的时候把这个加上保险点
        [DllImport("kernel32.dll")]
        static public extern int CloseHandle(int hObject);

        //查看数据传输异常函数
        [DllImport("kernel32.dll", EntryPoint = "GetProcAddress", SetLastError = true)]
        public static extern IntPtr GetProcAddress(int hModule, [MarshalAs(UnmanagedType.LPStr)] string lpProcName);


        /*构建一个Overlapped结构，异步通信用，
          internal是错误码，internalHigh是传输字节,这个两个是IO操作完成后需要填写的内容*/
        [StructLayout(LayoutKind.Sequential)]
        public struct OVERLAPPED
        {
            public IntPtr Internal;         //I/O请求的状态代码
            public IntPtr InternalHigh;     //传输I/O请求的字节数
            public int Offset;              //文件读写起始位置
            public int OffsetHigh;          //地址偏移量
            public IntPtr hEvent;           //操作完成时系统将设置为信号状态的事件句柄
        }


        /*监听异步通信函数*/
        [DllImport("Kernel32.dll")]
        public static extern long WaitForSingleObject(IntPtr hHandle, long dwMilliseconds);

        [DllImport("kernel32.dll")]
        public static extern IntPtr CreateEvent(IntPtr lpEventAttributes, int bManualReset, int bInitialState, string lpName);

        [DllImport("kernel32.dll", SetLastError = true)]
        public static extern bool GetOverlappedResult
            (IntPtr hFile,
           //[In] ref System.Threading.NativeOverlapped lpOverlapped,
           ref OVERLAPPED lpOverlapped,
           ref uint lpNumberOfBytesTransferred, 
           bool bWait);

        /*清除緩衝*/
        [DllImport("kernel32.dll")]
        public static extern bool FlushFileBuffers(IntPtr hFile);

        [DllImport("kernel32.dll", SetLastError = true)]
        public static extern bool PurgeComm(int hFile, uint dwFlags);
    }
}
