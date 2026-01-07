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
        //[DllImport("ole32.dll")]
        //static extern int CreateBindCtx(uint reserved, out IBindCtx ppbc);
        //[DllImport("ole32.dll")]
        //private static extern int GetRunningObjectTable(int reserved, out IRunningObjectTable prot);

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
            // 한글 파일 열 때 경고문 출현 방지 (레지스트리에 보안모듈 추가)
            RegisterSecurityModule();
        }

        /// <summary>
        /// 한/글 자동화 보안 모듈을 레지스트리에 등록하여 파일 접근 보안 창이 뜨지 않도록 합니다.
        /// </summary>
        private static void RegisterSecurityModule()
        {
            try
            {
                string moduleFileName = "FilePathCheckerModuleExample.dll";
                string executablePath = Application.StartupPath + @"\" + moduleFileName;

                // 레지스트리 키 경로
                string keyPath = @"Software\HNC\HwpAutomation\Modules";

                if (!System.IO.File.Exists(executablePath))
                {
                    MessageBox.Show("보안 모듈 파일이 존재하지 않습니다: " + executablePath, "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                using (RegistryKey key = Registry.CurrentUser.CreateSubKey(keyPath))
                {
                    if (key != null)
                    {
                        key.SetValue("FilePathCheckerModuleExample", executablePath, RegistryValueKind.String);
                        //key.SetValue("FilePathCheckDLL", executablePath, RegistryValueKind.String);
                    }
                }
            }
            catch (Exception ex)
            {
                // 레지스트리 접근 실패 시 오류 메시지를 보여주지만, 앱 실행은 계속됩니다.
                MessageBox.Show($"한/글 보안 모듈 레지스트리 등록 실패: {ex.Message}", "레지스트리 오류", MessageBoxButtons.OK, MessageBoxIcon.Warning);
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
                _hwpDocument = null;
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
                    _hwpDocument.SaveAsDistribution(saveFileDialog1.FileName, "12345");
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
            //_hwpDocument.CreateTable(int.Parse(txtTableRows.Text), int.Parse(txtTableRows.Text), 1, 1);
            //int rows = tableData.GetLength(0);
            //int cols = tableData.GetLength(1);

            _hwpDocument.CreateTable(5, 5, 1, 1);
        }

        private void btnSetTableCell_Click(object sender, EventArgs e)
        {
            if (!IsDocumentReady()) return;
            try
            {
                if (!int.TryParse(txtTableIndex.Text, out int tableIndex) ||
                    !int.TryParse(txtCellRow.Text, out int row) ||
                    !int.TryParse(txtCellCol.Text, out int col))
                {
                    MessageBox.Show("테이블 인덱스, 행, 열에 유효한 숫자를 입력하세요.", "입력 오류", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                _hwpDocument.SetTableCellText(tableIndex, row, col, txtCellText.Text);
            }
            catch (Exception ex)
            {
                HandleException(ex);
            }
        }

        private void btnInsertImage_Click(object sender, EventArgs e)
        {
            if (!IsDocumentReady()) return;

            openFileDialog1.Filter = "Image Files(*.BMP;*.JPG;*.GIF;*.PNG)|*.BMP;*.JPG;*.GIF;*.PNG|All files (*.*)|*.*";
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    _hwpDocument.InsertImage(openFileDialog1.FileName);
                }
                catch (Exception ex)
                {
                    HandleException(ex);
                }
            }
        }

        private void ChageDocument()
        {
            if (!IsDocumentReady()) return;

            // 제목 일자
            제목일자("2025.12.31");

            // 수입지출특이사항
            수입지출특이사항();

            // [표1] 총괄 현황
            총괄현황();

            // [표2] 조기지급 현황
            조기지급현황();

            // 선지급 미상환 잔액
            선지급미상환잔액();

            //// [표3] 세부 현황
            //세부현황();

            //// [표4] 자금예치 현황
            //자금예치현황();

            //// [표5] 당원수지현황당월
            //당월수지현황당월();

            //// [표6] 당기수지현황누적
            //당기수지현황누적();

            //// [표7] 누적수지현황
            //누적수지현황();

            //// [표8~표17] 연도별월별재정현황 
            //연도별월별재정현황(8);  // 표8    
            //연도별월별재정현황(9);  // 표9
            //연도별월별재정현황(10); // 표10
            //연도별월별재정현황(11); // 표11
            //연도별월별재정현황(12); // 표12
            //연도별월별재정현황(13); // 표13
            //연도별월별재정현황(14); // 표14
            //연도별월별재정현황(15); // 표15
            //연도별월별재정현황(16); // 표16
            //연도별월별재정현황(17); // 표17

            //// [표18] 월별보험급여비수입현황비교 
            //월별보험급여비수입현황비교();

            // [표19] 월별보험급여비지출현황비교 
            월별보험급여비지출현황비교();

            // 표를 빠져 나와서 다음 작업을 위해 한 줄 내림
            _hwpDocument.MoveOutOfTable();

            string p1 = "E:\\gungangbohum\\WindowsFormsApp1\\WindowsFormsApp1\\bin\\x86\\Debug\\그림1_월별보험급여비지출현황.png";
            string p2 = "E:\\gungangbohum\\WindowsFormsApp1\\WindowsFormsApp1\\bin\\x86\\Debug\\그림2_일자별보험급여비현황.png";
            string p3 = "E:\\gungangbohum\\WindowsFormsApp1\\WindowsFormsApp1\\bin\\x86\\Debug\\그림3_보험급여비지출현황.png";
            string p4 = "E:\\gungangbohum\\WindowsFormsApp1\\WindowsFormsApp1\\bin\\x86\\Debug\\그림4_보험급여비수입현황.png";

            // 그림1 월별보험급여비지출현황
            그림추가("그림1", p1);

            // 그림2 일자별보험급여비현황
            그림추가("그림2", p2);

            // 그림3 보험급여비지출현황
            그림추가("그림3", p3);

            // 그림4 보험급여비수입현황
            그림추가("그림4", p4);

            // 파일 다른이름으로 저장 및 한글 종료
            파일저장및종료();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            ChageDocument();
        }

        /// <summary>
        /// 제목일자
        /// </summary>
        /// <param name="strdate"></param>
        private void 제목일자(string strdate)
        {
            _hwpDocument.ReplaceAll("[E_제목일자]", strdate);
        }

        /// <summary>
        /// 수입지출특이사항
        /// </summary>
        private void 수입지출특이사항()
        {
            _hwpDocument.ReplaceAll("[E_보험료전월수입]", "7조 25억원");
            _hwpDocument.ReplaceAll("[E_보험료수입차액]", "5,056억 원 증가");
            _hwpDocument.ReplaceAll("[E_보험급여비전월지출]", "7조 8,578억원");
            _hwpDocument.ReplaceAll("[E_보험급여비차액]", "1조 6,971억 원 증가");
        }

        /// <summary>
        /// 선지급미상환잔액
        /// </summary>
        private void 선지급미상환잔액()
        {
            _hwpDocument.ReplaceAll("[E_선지급미상환잔액]", "(장봉섭) 300억 원");
        }
        /// <summary>
        /// [표1] 총괄현황
        /// </summary>
        private void 총괄현황()
        {
            int tableIndex = 1;

            // 총괄 현황 - 타이틀
            _hwpDocument.SetTableCellText(tableIndex, 0, 1, "2024년");
            _hwpDocument.MoveDownCell();
            _hwpDocument.InsertText("11월말", true);
            _hwpDocument.MoveNextCell();
            _hwpDocument.InsertText("12월말", true);
            _hwpDocument.MoveNextCell();
            //_hwpDocument.InsertText("연간 전망", true);
            _hwpDocument.MoveNextCell();
            _hwpDocument.InsertText("11월", true);

            _hwpDocument.MoveNextCell();
            _hwpDocument.InsertText("12월", true);
            _hwpDocument.MoveDownCell();
            //_hwpDocument.MoveLeftCell();
            _hwpDocument.InsertText("당일", true);
            _hwpDocument.MoveNextCell();
            _hwpDocument.InsertText("당월 누적", true);
            _hwpDocument.MoveNextCell();
            _hwpDocument.InsertText("연간 누적", true);

            _hwpDocument.SetTableCellText(tableIndex, 0, 2, "2025년");

            // 총괄현황 - 수입
            _hwpDocument.SetTableCellText(tableIndex, 1, 2, "100");  // 전년 11월말
            for (int i = 0; i < 3; i++)
            {
                _hwpDocument.MoveDownCell();
                _hwpDocument.InsertText((200 + i * 100).ToString(), true);
            }

            _hwpDocument.SetTableCellText(tableIndex, 1, 3, "110");  // 전년 12월말
            for (int i = 0; i < 3; i++)
            {
                _hwpDocument.MoveDownCell();
                _hwpDocument.InsertText((210 + i * 100).ToString(), true);
            }

            _hwpDocument.SetTableCellText(tableIndex, 1, 4, "120");  // 당년 연간 전망
            for (int i = 0; i < 3; i++)
            {
                _hwpDocument.MoveDownCell();
                _hwpDocument.InsertText((220 + i * 100).ToString(), true);
            }

            _hwpDocument.SetTableCellText(tableIndex, 1, 5, "130");  // 당년 전월
            for (int i = 0; i < 3; i++)
            {
                _hwpDocument.MoveDownCell();
                _hwpDocument.InsertText((230 + i * 100).ToString(), true);
            }

            _hwpDocument.SetTableCellText(tableIndex, 1, 6, "140");  // 당년 당월 당일
            for (int i = 0; i < 3; i++)
            {
                _hwpDocument.MoveDownCell();
                _hwpDocument.InsertText((240 + i * 100).ToString(), true);
            }

            _hwpDocument.SetTableCellText(tableIndex, 1, 7, "150");  // 당년 당월 누적
            for (int i = 0; i < 3; i++)
            {
                _hwpDocument.MoveDownCell();
                _hwpDocument.InsertText((250 + i * 100).ToString(), true);
            }

            _hwpDocument.SetTableCellText(tableIndex, 1, 8, "160");  // 당년 연간 누적
            for (int i = 0; i < 3; i++)
            {
                _hwpDocument.MoveDownCell();
                _hwpDocument.InsertText((260 + i * 100).ToString(), true);
            }

            // 총괄현황 - 지출
            _hwpDocument.SetTableCellText(tableIndex, 2, 2, "100");  // 전년 11월말
            for (int i = 0; i < 4; i++)
            {
                _hwpDocument.MoveDownCell();
                _hwpDocument.InsertText((200 + i * 100).ToString(), true);
            }

            _hwpDocument.SetTableCellText(tableIndex, 2, 3, "110");  // 전년 12월말
            for (int i = 0; i < 4; i++)
            {
                _hwpDocument.MoveDownCell();
                _hwpDocument.InsertText((210 + i * 100).ToString(), true);
            }

            _hwpDocument.SetTableCellText(tableIndex, 2, 4, "120");  // 당년 연간 전망
            for (int i = 0; i < 4; i++)
            {
                _hwpDocument.MoveDownCell();
                _hwpDocument.InsertText((220 + i * 100).ToString(), true);
            }

            _hwpDocument.SetTableCellText(tableIndex, 2, 5, "130");  // 당년 전월
            for (int i = 0; i < 4; i++)
            {
                _hwpDocument.MoveDownCell();
                _hwpDocument.InsertText((230 + i * 100).ToString(), true);
            }

            _hwpDocument.SetTableCellText(tableIndex, 2, 6, "140");  // 당년 당월 당일
            for (int i = 0; i < 4; i++)
            {
                _hwpDocument.MoveDownCell();
                _hwpDocument.InsertText((240 + i * 100).ToString(), true);
            }

            _hwpDocument.SetTableCellText(tableIndex, 2, 7, "150");  // 당년 당월 누적
            for (int i = 0; i < 4; i++)
            {
                _hwpDocument.MoveDownCell();
                _hwpDocument.InsertText((250 + i * 100).ToString(), true);
            }

            _hwpDocument.SetTableCellText(tableIndex, 2, 8, "160");  // 당년 연간 누적
            for (int i = 0; i < 4; i++)
            {
                _hwpDocument.MoveDownCell();
                _hwpDocument.InsertText((260 + i * 100).ToString(), true);
            }

            // 총괄현황 - 수지
            _hwpDocument.SetTableCellText(tableIndex, 3, 2, "100");  // 전년 11월말
            _hwpDocument.MoveDownCell();
            _hwpDocument.InsertText("200", true);

            _hwpDocument.SetTableCellText(tableIndex, 3, 3, "110");  // 전년 12월말
            _hwpDocument.MoveDownCell();
            _hwpDocument.InsertText("210", true);

            _hwpDocument.SetTableCellText(tableIndex, 3, 4, "120");  // 당년 연간 전망
            _hwpDocument.MoveDownCell();
            _hwpDocument.InsertText("220", true);

            _hwpDocument.SetTableCellText(tableIndex, 3, 5, "130");  // 당년 전월
            _hwpDocument.MoveDownCell();
            _hwpDocument.InsertText("230", true);

            _hwpDocument.SetTableCellText(tableIndex, 3, 6, "140");  // 당년 당월 당일
            _hwpDocument.MoveDownCell();
            _hwpDocument.InsertText("240", true);

            _hwpDocument.SetTableCellText(tableIndex, 3, 7, "150");  // 당년 당월 누적
            _hwpDocument.MoveDownCell();
            _hwpDocument.InsertText("250", true);

            _hwpDocument.SetTableCellText(tableIndex, 3, 8, "160");  // 당년 연간 누적
            _hwpDocument.MoveDownCell();
            _hwpDocument.InsertText("260", true);
        }

        /// <summary>
        /// [표2] 조기지급현황
        /// </summary>
        private void 조기지급현황()
        {
            int tableIndex = 2;

            _hwpDocument.SetTableCellText(tableIndex, 1, 1, "100개");     // 당일 기관
            _hwpDocument.SetTableCellText(tableIndex, 1, 2, "100억원");   // 당일 금액
            _hwpDocument.SetTableCellText(tableIndex, 1, 3, "1000개");    // 누적 기관
            _hwpDocument.SetTableCellText(tableIndex, 1, 4, "1000억원");  // 누적 금액
        }

        /// <summary>
        /// [표3] 세부현황
        /// </summary>
        private void 세부현황()
        {
            int tableIndex = 3;

            // 세부 현황 - 타이틀
            _hwpDocument.SetTableCellText(tableIndex, 0, 1, "2024년");
            _hwpDocument.MoveDownCell();
            _hwpDocument.InsertText("11월말", true);
            _hwpDocument.MoveNextCell();
            _hwpDocument.InsertText("12월말", true);
            _hwpDocument.MoveNextCell();
            //_hwpDocument.InsertText("연간 전망", true);
            _hwpDocument.MoveNextCell();
            _hwpDocument.MoveUpCell();
            _hwpDocument.InsertText("11월", true);

            _hwpDocument.MoveNextCell();
            _hwpDocument.InsertText("12월", true);
            _hwpDocument.MoveDownCell();
            //_hwpDocument.MoveLeftCell();
            //_hwpDocument.InsertText("당일", true);
            //_hwpDocument.MoveNextCell();
            //_hwpDocument.InsertText("당월 누적", true);
            //_hwpDocument.MoveNextCell();
            //_hwpDocument.InsertText("연간 누적", true);

            _hwpDocument.SetTableCellText(tableIndex, 0, 2, "2025년");

            // 세부현황 - 수입
            _hwpDocument.SetTableCellText(tableIndex, 1, 2, "100");  // 전년 11월말
            for (int i = 0; i < 11; i++)
            {
                _hwpDocument.MoveDownCell();
                _hwpDocument.InsertText((200 + i * 100).ToString(), true);
            }

            _hwpDocument.SetTableCellText(tableIndex, 1, 3, "110");  // 전년 12월말
            for (int i = 0; i < 11; i++)
            {
                _hwpDocument.MoveDownCell();
                _hwpDocument.InsertText((210 + i * 100).ToString(), true);
            }

            _hwpDocument.SetTableCellText(tableIndex, 1, 4, "120");  // 당년 연간 전망
            for (int i = 0; i < 11; i++)
            {
                _hwpDocument.MoveDownCell();
                _hwpDocument.InsertText((220 + i * 100).ToString(), true);
            }

            _hwpDocument.SetTableCellText(tableIndex, 1, 5, "130");  // 당년 전월 연간누적 금액
            for (int i = 0; i < 11; i++)
            {
                _hwpDocument.MoveDownCell();
                _hwpDocument.InsertText((230 + i * 100).ToString(), true);
            }

            _hwpDocument.SetTableCellText(tableIndex, 1, 6, "140");  // 당년 전월 연간누적 집행률
            for (int i = 0; i < 11; i++)
            {
                _hwpDocument.MoveDownCell();
                _hwpDocument.InsertText((240 + i * 100).ToString(), true);
            }

            _hwpDocument.SetTableCellText(tableIndex, 1, 7, "150");  // 당년 전월 연간누적 전년대비
            for (int i = 0; i < 11; i++)
            {
                _hwpDocument.MoveDownCell();
                _hwpDocument.InsertText((250 + i * 100).ToString(), true);
            }

            _hwpDocument.SetTableCellText(tableIndex, 1, 8, "160");  // 당년 당월 전일누적
            for (int i = 0; i < 11; i++)
            {
                _hwpDocument.MoveDownCell();
                _hwpDocument.InsertText((260 + i * 100).ToString(), true);
            }

            _hwpDocument.SetTableCellText(tableIndex, 1, 9, "170");  // 당년 당월 당일
            for (int i = 0; i < 11; i++)
            {
                _hwpDocument.MoveDownCell();
                _hwpDocument.InsertText((270 + i * 100).ToString(), true);
            }

            _hwpDocument.SetTableCellText(tableIndex, 1, 10, "180");  // 당년 당월 당월누적
            for (int i = 0; i < 11; i++)
            {
                _hwpDocument.MoveDownCell();
                _hwpDocument.InsertText((280 + i * 100).ToString(), true);
            }

            _hwpDocument.SetTableCellText(tableIndex, 1, 11, "190");  // 당년 연간누적 F
            for (int i = 0; i < 11; i++)
            {
                _hwpDocument.MoveDownCell();
                _hwpDocument.InsertText((290 + i * 100).ToString(), true);
            }

            _hwpDocument.SetTableCellText(tableIndex, 1, 12, "1901");  // 당년 연간누적 집행률
            for (int i = 0; i < 11; i++)
            {
                _hwpDocument.MoveDownCell();
                _hwpDocument.InsertText((2901 + i * 100).ToString(), true);
            }

            // 세부현황 - 지출
            _hwpDocument.SetTableCellText(tableIndex, 2, 2, "100");  // 전년 11월말
            for (int i = 0; i < 12; i++)
            {
                _hwpDocument.MoveDownCell();
                _hwpDocument.InsertText((200 + i * 100).ToString(), true);
            }

            _hwpDocument.SetTableCellText(tableIndex, 2, 3, "110");  // 전년 12월말
            for (int i = 0; i < 12; i++)
            {
                _hwpDocument.MoveDownCell();
                _hwpDocument.InsertText((210 + i * 100).ToString(), true);
            }

            _hwpDocument.SetTableCellText(tableIndex, 2, 4, "120");  // 당년 연간 전망
            for (int i = 0; i < 12; i++)
            {
                _hwpDocument.MoveDownCell();
                _hwpDocument.InsertText((220 + i * 100).ToString(), true);
            }

            _hwpDocument.SetTableCellText(tableIndex, 2, 5, "130");  // 당년 전월
            for (int i = 0; i < 12; i++)
            {
                _hwpDocument.MoveDownCell();
                _hwpDocument.InsertText((230 + i * 100).ToString(), true);
            }

            _hwpDocument.SetTableCellText(tableIndex, 2, 6, "140");  // 당년 당월 당일
            for (int i = 0; i < 12; i++)
            {
                _hwpDocument.MoveDownCell();
                _hwpDocument.InsertText((240 + i * 100).ToString(), true);
            }

            _hwpDocument.SetTableCellText(tableIndex, 2, 7, "150");  // 당년 당월 누적
            for (int i = 0; i < 12; i++)
            {
                _hwpDocument.MoveDownCell();
                _hwpDocument.InsertText((250 + i * 100).ToString(), true);
            }

            _hwpDocument.SetTableCellText(tableIndex, 2, 8, "160");  // 당년 연간 누적
            for (int i = 0; i < 12; i++)
            {
                _hwpDocument.MoveDownCell();
                _hwpDocument.InsertText((260 + i * 100).ToString(), true);
            }

            _hwpDocument.SetTableCellText(tableIndex, 2, 9, "170");  // 당년 당월 당일
            for (int i = 0; i < 12; i++)
            {
                _hwpDocument.MoveDownCell();
                _hwpDocument.InsertText((270 + i * 100).ToString(), true);
            }

            _hwpDocument.SetTableCellText(tableIndex, 2, 10, "180");  // 당년 당월 당월누적
            for (int i = 0; i < 12; i++)
            {
                _hwpDocument.MoveDownCell();
                _hwpDocument.InsertText((280 + i * 100).ToString(), true);
            }

            _hwpDocument.SetTableCellText(tableIndex, 2, 11, "190");  // 당년 연간누적 F
            for (int i = 0; i < 12; i++)
            {
                _hwpDocument.MoveDownCell();
                _hwpDocument.InsertText((290 + i * 100).ToString(), true);
            }

            _hwpDocument.SetTableCellText(tableIndex, 2, 12, "1901");  // 당년 연간누적 집행률
            for (int i = 0; i < 12; i++)
            {
                _hwpDocument.MoveDownCell();
                _hwpDocument.InsertText((2901 + i * 100).ToString(), true);
            }

            // 세부현황 - 수지
            _hwpDocument.SetTableCellText(tableIndex, 3, 2, "100");  // 전년 11월말
            _hwpDocument.MoveDownCell();
            _hwpDocument.InsertText("200", true);

            _hwpDocument.SetTableCellText(tableIndex, 3, 3, "110");  // 전년 12월말 연간누적
            _hwpDocument.MoveDownCell();
            _hwpDocument.InsertText("210", true);

            _hwpDocument.SetTableCellText(tableIndex, 3, 4, "120");  // 당년 연간 전망
            _hwpDocument.MoveDownCell();
            _hwpDocument.InsertText("220", true);

            _hwpDocument.SetTableCellText(tableIndex, 3, 5, "130");  // 당년 전월 연간누적 금액
            _hwpDocument.MoveDownCell();
            _hwpDocument.InsertText("230", true);

            _hwpDocument.SetTableCellText(tableIndex, 3, 6, "140");  // 당년 전월월 연간누적 집행률
            _hwpDocument.MoveDownCell();
            _hwpDocument.InsertText("240", true);

            _hwpDocument.SetTableCellText(tableIndex, 3, 7, "150");  // 당년 전월 연간누적 전년대비
            _hwpDocument.MoveDownCell();
            _hwpDocument.InsertText("250", true);

            _hwpDocument.SetTableCellText(tableIndex, 3, 8, "160");  // 당년 당월 전일누적
            _hwpDocument.MoveDownCell();
            _hwpDocument.InsertText("260", true);

            _hwpDocument.SetTableCellText(tableIndex, 3, 9, "170");  // 당년 당월 당일
            _hwpDocument.MoveDownCell();
            _hwpDocument.InsertText("270", true);

            _hwpDocument.SetTableCellText(tableIndex, 3, 10, "180");  // 당년 당월 당월누적
            _hwpDocument.MoveDownCell();
            _hwpDocument.InsertText("280", true);

            _hwpDocument.SetTableCellText(tableIndex, 3, 11, "190");  // 당년 연간누적 F
            _hwpDocument.MoveDownCell();
            _hwpDocument.InsertText("290", true);

            _hwpDocument.SetTableCellText(tableIndex, 3, 12, "1901");  // 당년 연간누적 집행률
            _hwpDocument.MoveDownCell();
            _hwpDocument.InsertText("2901", true);

        }

        /// <summary>
        /// [표4] 자금예치현황
        /// </summary>
        private void 자금예치현황()
        {
            int tableIndex = 4;

            _hwpDocument.SetTableCellText(tableIndex, 0, 1, "2024년");
            _hwpDocument.SetTableCellText(tableIndex, 0, 2, "2025년");
            _hwpDocument.MoveDownCell();
            _hwpDocument.InsertText("전월(11월)", true);
            _hwpDocument.MoveNextCell();
            _hwpDocument.InsertText("당월(12월)", true);

            // 자금예치 현황 
            for (int c = 0; c < 12; c++)
            {
                _hwpDocument.SetTableCellText(tableIndex, 1, c + 1, (100 + c * 100).ToString());
                for (int r = 0; r < 10; r++)
                {
                    _hwpDocument.MoveDownCell();
                    _hwpDocument.InsertText((r + 100 + c * 100).ToString(), true);
                }
            }

        }

        /// <summary>
        /// [표5] 당월수지현황당월
        /// </summary>
        private void 당월수지현황당월()
        {
            int tableIndex = 5;

            // 년도
            _hwpDocument.SetTableCellText(tableIndex, 1, 0, "2016년");
            for (int r = 0; r < 9; r++)
            {
                _hwpDocument.MoveDownCell();
                _hwpDocument.InsertText($"{(2017 + r).ToString()}년" , true);
            }

            // 월별현황
            for (int c = 0; c < 12; c++)
            {
                _hwpDocument.SetTableCellText(tableIndex, 1, c + 1, (100 + c * 100).ToString());
                for (int r = 0; r < 9; r++)
                {
                    _hwpDocument.MoveDownCell();
                    _hwpDocument.InsertText((r + 100 + c * 100).ToString(), true);
                }
            }

        }

        /// <summary>
        /// [표6] 당기수지현황누적
        /// </summary>
        private void 당기수지현황누적()
        {
            int tableIndex = 6;

            // 년도
            _hwpDocument.SetTableCellText(tableIndex, 1, 0, "2016년");
            for (int r = 0; r < 9; r++)
            {
                _hwpDocument.MoveDownCell();
                _hwpDocument.InsertText($"{(2017 + r).ToString()}년", true);
            }

            // 월별현황
            for (int c = 0; c < 12; c++)
            {
                _hwpDocument.SetTableCellText(tableIndex, 1, c + 1, (100 + c * 100).ToString());
                for (int r = 0; r < 9; r++)
                {
                    _hwpDocument.MoveDownCell();
                    _hwpDocument.InsertText((r + 100 + c * 100).ToString(), true);
                }
            }

        }

        /// <summary>
        /// [표7] 누적수지현황
        /// </summary>
        private void 누적수지현황()
        {
            int tableIndex = 7;

            // 년도
            _hwpDocument.SetTableCellText(tableIndex, 1, 0, "2016년");
            for (int r = 0; r < 9; r++)
            {
                _hwpDocument.MoveDownCell();
                _hwpDocument.InsertText($"{(2017 + r).ToString()}년", true);
            }

            // 월별현황
            for (int c = 0; c < 12; c++)
            {
                _hwpDocument.SetTableCellText(tableIndex, 1, c + 1, (100 + c * 100).ToString());
                for (int r = 0; r < 9; r++)
                {
                    _hwpDocument.MoveDownCell();
                    _hwpDocument.InsertText((r + 100 + c * 100).ToString(), true);
                }
            }

        }

        /// <summary>
        /// [표8] ~ [표17] 연도별월별재정현황
        /// </summary>
        private void 연도별월별재정현황(int tableIndex)
        {
            //int tableIndex = 8;

            for (int c = 0; c < 12; c++)
            {
                _hwpDocument.SetTableCellText(tableIndex, 1, c + 2, (100 + c * 100).ToString());
                for (int r = 0; r < 29; r++)
                {
                    _hwpDocument.MoveDownCell();
                    _hwpDocument.InsertText((r + 100 + c * 100).ToString(), true);
                }
            }

        }

        /// <summary>
        /// [표18] 월별보험급여비지출현황비교
        /// </summary>
        private void 월별보험급여비지출현황비교()
        {
            int tableIndex = 18;

            // 년도
            _hwpDocument.SetTableCellText(tableIndex, 1, 1, "2016");
            for (int r = 0; r < 9; r++)
            {
                _hwpDocument.MoveDownCell();
                _hwpDocument.InsertText($"{(2017 + r).ToString()}", true);
            }

            // 월별현황
            for (int c = 0; c < 13; c++)
            {
                _hwpDocument.SetTableCellText(tableIndex, 1, c + 3, (100 + c * 100).ToString());
                for (int r = 0; r < 19; r++)
                {
                    _hwpDocument.MoveDownCell();
                    _hwpDocument.InsertText((r + 100 + c * 100).ToString(), true);
                }
            }

        }

        /// <summary>
        /// [표19] 월별보험급여비수입현황비교
        /// </summary>
        private void 월별보험급여비수입현황비교()
        {
            int tableIndex = 19;

            // 년도
            _hwpDocument.SetTableCellText(tableIndex, 1, 1, "2016");
            for (int r = 0; r < 9; r++)
            {
                _hwpDocument.MoveDownCell();
                _hwpDocument.InsertText($"{(2017 + r).ToString()}", true);
            }

            // 월별현황
            for (int c = 0; c < 13; c++)
            {
                _hwpDocument.SetTableCellText(tableIndex, 1, c + 3, (100 + c * 100).ToString());
                for (int r = 0; r < 19; r++)
                {
                    _hwpDocument.MoveDownCell();
                    _hwpDocument.InsertText((r + 100 + c * 100).ToString(), true);
                }
            }

        }

        private void 그림추가(string field, string imgPath)
        {
            _hwpDocument.MoveToField(field);
            _hwpDocument.InsertImage(imgPath);
        }

        private void 파일저장및종료()
        {
            string savePath = System.IO.Directory.GetCurrentDirectory() + @"\일일재정현황2.hwp";
            _hwpDocument.SaveAs(savePath);

            _hwpDocument.Close(save: false);
            _hwpApplication.Quit();
        }
    }
}
