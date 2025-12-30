using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using HWPCONTROLLib;
using System.Runtime.InteropServices.ComTypes;
using HwpObjectLib;
using Microsoft.Win32;
using Eum.Hwp;

namespace WindowsFormsApp1
{
    public partial class HangulTest : Form
    {
        [DllImport("ole32.dll")]
        static extern int CreateBindCtx(uint reserved, out IBindCtx ppbc);
        [DllImport("ole32.dll")]
        private static extern int GetRunningObjectTable(int reserved, out IRunningObjectTable prot);

        private HwpDocument _hwpDocument;
        private HwpApplication _hwpApplication;

        public HangulTest()
        {
            InitializeComponent();
        }

        private void btnNewHwp_Click(object sender, EventArgs e)
        {
            string[,] data = {
              { "구분", "Q1", "Q2", "Q3", "Q4" },
              { "서울", "120", "135", "142", "160" },
              { "부산", "120", "135", "142", "160" },
              { "대구", "120", "135", "142", "160" },
              { "제주", "120", "135", "142", "160" },
              { "합계", "260", "295", "312", "358" }
            };

            using (var app = new HwpApplication(visible: true))
            {
                var doc = app.NewDocument();
                doc.InsertTitleTableAndDescription(
                    "분기별 매출 요약",
                    data,
                    "설명: 위 표는 샘플 데이터이며 지역별 분기 매출 추이를 요약한 것입니다."
                );
                string savePath = System.IO.Directory.GetCurrentDirectory() + @"\output1.hwp";
                doc.SaveAs(savePath);
                doc.Close();
            }
        }

        
        private void Form1_Load(object sender, EventArgs e)
        {
            // 한글 파일 열 때 경고문 출현 방지
            // 레지스트리에 보안모듈 추가
            string checkFile = Application.StartupPath + @"\FilePathCheckerModuleExample.dll";
            if (!System.IO.File.Exists(checkFile))
            {
                MessageBox.Show("보안 모듈 파일이 존재하지 않습니다: " + checkFile, "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            RegistryKey key = Registry.CurrentUser.OpenSubKey(@"SOFTWARE\HNC\HwpAutomation\Modules", true);
            if (key == null)
            {
                key = Registry.CurrentUser.CreateSubKey(@"SOFTWARE\HNC\HwpAutomation\Modules");
            }
            if (key.GetValue("FilePathCheckerModuleExample") == null)
            {
                key.SetValue("FilePathCheckerModuleExample", Application.StartupPath + @"\FilePathCheckerModuleExample.dll");
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            using (var app = new HwpApplication(visible: true))
            {
                string filePath = System.IO.Directory.GetCurrentDirectory() + @"\output2.hwp";
                string savePath = System.IO.Directory.GetCurrentDirectory() + @"\output3.hwp";

                var doc = app.OpenDocument(filePath);
                doc.UpdateTableCellByCaption("간다", 2, 2, "555");
                doc.SaveAs(savePath);
                doc.Close(save: false);
            }
        }

        private void btnChangeHwp1_Click(object sender, EventArgs e)
        {
            string[,] newTable = {
              { "구분", "Q1", "Q2", "Q3", "Q4" },
              { "여수", "130", "140", "150", "170" },
              { "순천", "85",  "92",  "98",  "120" },
              { "광양", "65",  "78",  "82",  "95"  },
              { "합계", "280", "310", "330", "385" }
            };

            using (var app = new HwpApplication(visible: true))
            {
                string filePath = System.IO.Directory.GetCurrentDirectory() + @"\org.hwp";
                string savePath = System.IO.Directory.GetCurrentDirectory() + @"\output2.hwp";

                var doc = app.OpenDocument(filePath);

                doc.ReplaceAll("[문서제목]", "수정된 문서제목");

                doc.ReplaceAll("[표1제목]", "표1제목을 수정");
                doc.ReplaceAll("[표1컬럼1]", "컬럼1");
                doc.ReplaceAll("[표1컬럼2]", "컬럼2");
                doc.ReplaceAll("[표1컬럼3]", "컬럼3");
                doc.ReplaceAll("[표1컬럼4]", "컬럼4");
                doc.ReplaceAll("[표1컬럼5]", "컬럼5");

                doc.ReplaceAll("[표2제목]", "표2제목을 수정");
                doc.ReplaceAll("[표2컬럼1]", "컬럼1");
                doc.ReplaceAll("[표2컬럼2]", "컬럼2");
                doc.ReplaceAll("[표2컬럼3]", "컬럼3");
                doc.ReplaceAll("[표2컬럼4]", "컬럼4");
                doc.ReplaceAll("[표2컬럼5]", "컬럼5");

                doc.ReplaceAll("[공통]", "[이것은 공통 항목이 수정된 것입니다.]");

                doc.SaveAs(savePath);
                doc.Close(save: false);
            }
            //UpdateTableCellByCaption(string captionText, int row, int col, string value, bool captionAboveTable = true)
        }

        private void btnSaveAsPdf_Click(object sender, EventArgs e)
        {
            saveFileDialog1.Filter = "PDF files (*.pdf)|*.pdf|All files (*.*)|*.*";
            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                if (!IsDocumentReady()) return;

                try
                {
                    _hwpDocument.SaveAs(saveFileDialog1.FileName, "PDF");
                    MessageBox.Show("Saved as PDF.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                catch (Exception ex)
                {
                    HandleException(ex);
                }
            }

        }
        
        private bool IsDocumentReady()
        {
            if (_hwpDocument == null)
            {
                MessageBox.Show("Please create or open a document first.", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return false;
            }
            return true;
        }

        private void HandleException(Exception ex)
        {
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        private void btnOpenDocument_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                _hwpApplication = new HwpApplication(visible: true);
                _hwpDocument = _hwpApplication.OpenDocument(openFileDialog1.FileName);
            }
        }

        private void btnSaveAsDistribution_Click(object sender, EventArgs e)
        {
            if (!IsDocumentReady()) return;

            saveFileDialog1.Filter = "HWP files (*.hwp)|*.hwp|All files (*.*)|*.*";
            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    _hwpDocument.SaveAsDistribution(saveFileDialog1.FileName, "1234");
                    MessageBox.Show("Saved as distribution document.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                catch (Exception ex)
                {
                    HandleException(ex);
                }
            }
        }

        private void btnInsertText_Click(object sender, EventArgs e)
        {
            if (!IsDocumentReady()) return;
            _hwpDocument.InsertText(txtInsertText.Text);
        }

        private void btnReplaceAll_Click(object sender, EventArgs e)
        {
            if (!IsDocumentReady()) return;
            _hwpDocument.ReplaceAll(txtFindText.Text, txtReplaceText.Text);
        }

        private void btnCreateTable_Click(object sender, EventArgs e)
        {
            if (!IsDocumentReady()) return;
            _hwpDocument.CreateTable(int.Parse(txtTableRows.Text), int.Parse(txtTableRows.Text), 1, 1);
        }
    }
}
