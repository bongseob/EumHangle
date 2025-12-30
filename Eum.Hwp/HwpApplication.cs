using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Eum.Hwp
{
    public sealed class HwpApplication : IDisposable
    {
        private dynamic _hwp;
        private bool _visible;

        public HwpApplication(bool visible = false)
        {
            _visible = visible;
            _hwp = HwpComFactory.CreateHwpObject();
            _hwp.XHwpWindows.Item(0).Visible = visible;
        }

        public HwpDocument NewDocument()
        {
            _hwp.HAction.GetDefault("FileNew", _hwp.HParameterSet.HFileOpenSave.HSet);
            _hwp.HAction.Execute("FileNew", _hwp.HParameterSet.HFileOpenSave.HSet);
            return new HwpDocument(_hwp);
        }

        public HwpDocument OpenDocument(string filePath)
        {
            _hwp.HAction.GetDefault("FileOpen", _hwp.HParameterSet.HFileOpenSave.HSet);
            _hwp.HParameterSet.HFileOpenSave.filename = filePath;
            _hwp.HParameterSet.HFileOpenSave.Format = "HWP"; // 필요 시 "HWPX"
            _hwp.HAction.Execute("FileOpen", _hwp.HParameterSet.HFileOpenSave.HSet);
            return new HwpDocument(_hwp);
        }

        public void SetVisible(bool visible)
        {
            _visible = visible;
            _hwp.XHwpWindows.Item(0).Visible = visible;
        }

        public void Quit()
        {
            try { _hwp.Quit(); }
            catch { /* ignore */ }
        }

        public void Dispose()
        {
            try { Quit(); } catch { }
            //HwpComFactory.ReleaseComObject(_hwp);
            _hwp = null;

            // 선택: 잔여 RCW 정리를 원하면 사용
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
    }
}
