using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace Eum.Hwp
{
    internal static class HwpComFactory
    {
        public static dynamic CreateHwpObject()
        {
            var type = Type.GetTypeFromProgID("HWPFrame.HwpObject");
            if (type == null)
                throw new InvalidOperationException("한글(HWP) COM이 등록되어 있지 않습니다.");

            return Activator.CreateInstance(type);
        }

        // 한글 COM 객체 릴리즈 메서드 (필요 시 사용)
        //public static void ReleaseComObject(object com)
        //{
        //    //if (com != null && Marshal.IsComObject(com))
        //    //    Marshal.FinalReleaseComObject(com);
        //    if (com == null) return;

        //    try
        //    {
        //        if (Marshal.IsComObject(com)) return;

        //        //Marshal.FinalReleaseComObject(com);
        //        // FinalReleaseComObject 대신 ReleaseComObject 루프 사용
        //        while (Marshal.ReleaseComObject(com) > 0) { }

        //    }            
        //    catch
        //    {
        //        // 기타 예외도 릴리즈 단계에서는 무시
        //    }
        //}
    }
}
