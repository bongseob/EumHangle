using Microsoft.SqlServer.Server;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace Eum.Hwp
{
    public sealed class HwpDocument
    {
        private readonly dynamic _hwp;

        internal HwpDocument(dynamic hwp)
        {
            _hwp = hwp;
        }

        private dynamic CreateAction(string actionId)
        {
            return _hwp.CreateAction(actionId);
        }

        private dynamic CreateSet(dynamic action)
        {
            return action.CreateSet();
        }

        public void SaveAs(string filePath, string format = "HWP")
        {
            _hwp.HAction.GetDefault("FileSaveAs", _hwp.HParameterSet.HFileOpenSave.HSet);
            _hwp.HParameterSet.HFileOpenSave.filename = filePath;
            _hwp.HParameterSet.HFileOpenSave.Format = format; // 필요 시 "HWPX" 등으로 변경
            _hwp.HAction.Execute("FileSaveAs", _hwp.HParameterSet.HFileOpenSave.HSet);
        }

        public void InsertText(string text, bool bSelect=false)
        {
            if (bSelect)
            {
                _hwp.HAction.Run("Cancel");
                _hwp.HAction.Run("MoveLineBegin");
                _hwp.HAction.Run("MoveSelLineEnd");
                _hwp.HAction.Run("Delete");

                _hwp.HAction.GetDefault("InsertText", _hwp.HParameterSet.HInsertText.HSet);
                _hwp.HParameterSet.HInsertText.Text = text;
                _hwp.HAction.Execute("InsertText", _hwp.HParameterSet.HInsertText.HSet);

            }
            else
            {
                _hwp.HAction.GetDefault("InsertText", _hwp.HParameterSet.HInsertText.HSet);
                _hwp.HParameterSet.HInsertText.Text = text;
                _hwp.HAction.Execute("InsertText", _hwp.HParameterSet.HInsertText.HSet);
            }
        }

        public void MoveToField(string fieldName)
        {
            bool text = true;
            bool start = true;
            bool select = true;
            _hwp.MoveToField(fieldName, text, start, select);
        }

        /// <summary>
        /// 지정된 경로의 이미지를 현재 커서 위치에 삽입합니다.
        /// https://forum.developer.hancom.com/t/c-dll-insertbackgroundpicture/41
        /// axHwpCtrl1.InsertBackgroundPicture("SelectedCell","C:/test/test.bmp");
        /// (이미지삽입) 수행 이전에 vHwpCtrl.SetMessageBoxMode(0x00000100); 을 먼저 수행해주어야 합니다.
        /// 다시 설정을 해제하는 vHwpCtrl.SetMessageBoxMode(0x00000000);를 적용
        /// </summary>
        public void InsertImage(string imagePath)
        {
            bool embedded = true;    // 이미지 파일을 문서내에 포함할지 여부 (True/False). 생략하면 true
            int sizeoption = 2;      // 삽입할 그림의 크기 옵션 0: 원본 크기 1: 사용자가 지정한 크기 (mmPicWidth, mmPicHeight 값 사용) 2:문서에 맞게 자동 조절
            bool reverse = false;    // 이미지 반전 유무
            bool watermark = false;  // 워터마크 여부
            int effect = 0;          // 이미지 효과 0: 없음
            int mmPicWidth = 0;      // 그림 가로 크기 (sizeoption이 1일 때 사용, 단위: mm)
            int mmPicHeight = 0;     // 그림 세로 크기 (sizeoption이 1일 때 사용, 단위: mm)
            _hwp.InsertPicture(imagePath, embedded, sizeoption, reverse, watermark, effect, mmPicWidth, mmPicHeight);
        }

        public void DeleteTable(int tableIndex)
        {
            if (!MoveToTableByIndex(tableIndex))
                throw new InvalidOperationException($"표를 찾지 못했습니다. index={tableIndex}");

            _hwp.HAction.GetDefault("DeleteCtrl", _hwp.HParameterSet.HDeleteCtrl.HSet);
            _hwp.HAction.Execute("DeleteCtrl", _hwp.HParameterSet.HDeleteCtrl.HSet);
        }


        /// <summary>
        /// 지정된 표의 특정 셀에 텍스트를 입력합니다.
        /// </summary>
        public void SetTableCellText(int tableIndex, int row, int col, string text)
        {
            /*
            string tableFieldList = _hwp.GetFieldList(0, "tbl");
            string[] tableFields = tableFieldList.Split('\x02');

            if (string.IsNullOrEmpty(tableFieldList) || tableIndex < 0 || tableIndex >= tableFields.Length)
            {
                throw new System.IndexOutOfRangeException($"테이블을 찾을 수 없거나 테이블 인덱스({tableIndex})가 잘못되었습니다. 문서에 있는 테이블 수: {(string.IsNullOrEmpty(tableFieldList) ? 0 : tableFields.Length)}");
            }

            string targetTableField = tableFields[tableIndex];

            if (!_hwp.MoveToField(targetTableField, true, true, true))
            {
                throw new System.Exception($"문서에서 '{targetTableField}' 표를 찾는 데 실패했습니다.");
            }

            _hwp.HAction.GetDefault("TableCellBlock", _hwp.HParameterSet.HTableCellBlock.HSet);
            _hwp.HParameterSet.HTableCellBlock.StartRow = row;
            _hwp.HParameterSet.HTableCellBlock.StartCol = col;
            _hwp.HParameterSet.HTableCellBlock.EndRow = row;
            _hwp.HParameterSet.HTableCellBlock.EndCol = col;
            _hwp.HAction.Execute("TableCellBlock", _hwp.HParameterSet.HTableCellBlock.HSet);

            _hwp.HAction.GetDefault("Delete", _hwp.HParameterSet.HSelection.HSet);
            _hwp.HAction.Execute("Delete", _hwp.HParameterSet.HSelection.HSet);

            InsertText(text);

            _hwp.MoveToField(targetTableField, false, false, false);
            */

            if (tableIndex < 0 || row < 0 || col < 0)
            {
                throw new System.IndexOutOfRangeException($"테이블을 찾을 수 없거나 테이블 인덱스({tableIndex})가 잘못되었습니다.");
            }

            MoveToTableByIndex(tableIndex);
            MoveToCell(row, col);
            InsertText(text, true);

        }

        public void ReplaceAll(string find, string replace)
        {
            ReplaceAll(find, replace, FindReplaceOptions.Default);
        }

        public void ReplaceAll(string find, string replace, FindReplaceOptions options)
        {
            if (options == null) throw new ArgumentNullException(nameof(options));

            _hwp.HAction.GetDefault("AllReplace", _hwp.HParameterSet.HFindReplace.HSet);
            _hwp.HParameterSet.HFindReplace.FindString = find;
            _hwp.HParameterSet.HFindReplace.ReplaceString = replace;
            ApplyFindReplaceOptions(options);
            _hwp.HAction.Execute("AllReplace", _hwp.HParameterSet.HFindReplace.HSet);
        }

        public void Close(bool save = false)
        {
            if (save) Save();

            _hwp.HAction.GetDefault("FileClose", _hwp.HParameterSet.HFileOpenSave.HSet);
            _hwp.HAction.Execute("FileClose", _hwp.HParameterSet.HFileOpenSave.HSet);
        }

        public void Save()
        {
            _hwp.HAction.GetDefault("FileSave", _hwp.HParameterSet.HFileOpenSave.HSet);
            _hwp.HAction.Execute("FileSave", _hwp.HParameterSet.HFileOpenSave.HSet);
        }

        public void InsertTitleTableAndDescription(string title, string[,] tableData, string description)
        {
            if (tableData == null) throw new ArgumentNullException(nameof(tableData));

            // 1) 제목 (글씨 크게)
            InsertText(title);
            ApplyCharSizeToCurrentPara(20);
            InsertParagraph();

            // 2) 표 생성 및 채우기
            int rows = tableData.GetLength(0);
            int cols = tableData.GetLength(1);

            CreateTable(rows, cols);
            FillTable(tableData);

            // 3) 표 밖으로 이동 후 설명 작성
            MoveOutOfTable();
            InsertText(description);
            ApplyCharSizeToCurrentPara(10);
        }


        public void UpdateTableCellByIndex(int tableIndexZeroBased, int row, int col, string value)
        {
            if (row < 0 || col < 0) throw new ArgumentOutOfRangeException();

            if (!MoveToTableByIndex(tableIndexZeroBased))
                throw new InvalidOperationException($"표를 찾지 못했습니다. index={tableIndexZeroBased}");

            // 첫 셀로 이동 후 원하는 셀로 이동
            MoveToFirstCell();
            MoveToCell(row, col);

            // 셀 내용 교체
            ReplaceCurrentCellText(value);
        }

        /// <summary>
        /// 표 캡션(제목)으로 표를 찾은 후, 지정된 셀의 내용을 변경
        /// </summary>
        /// <param name="captionText"></param>
        /// <param name="row"></param>
        /// <param name="col"></param>
        /// <param name="value"></param>
        /// <param name="captionAboveTable"></param>
        /// <exception cref="ArgumentException"></exception>
        /// <exception cref="ArgumentOutOfRangeException"></exception>
        /// <exception cref="InvalidOperationException"></exception>
        public void UpdateTableCellByCaption(string captionText, int row, int col, string value, bool captionAboveTable = true)
        {
            if (string.IsNullOrWhiteSpace(captionText)) throw new ArgumentException("captionText is required.");
            if (row < 0 || col < 0) throw new ArgumentOutOfRangeException();

            if (!MoveToTableByCaption(captionText, captionAboveTable))
                throw new InvalidOperationException($"표 캡션을 찾지 못했습니다. caption={captionText}");

            // 표 안으로 진입
            if (!EnsureCaretInTableCell())
                throw new InvalidOperationException("표 셀 진입 실패");

            // 첫 셀로 이동 후 원하는 셀로 이동
            MoveToFirstCell();
            MoveToCell(row, col);

            // 셀 내용 교체
            ReplaceCurrentCellText(value);
        }

        private bool MoveToTableByCaption(string captionText, bool captionAboveTable)
        {
            TryAction("MoveDocBegin");

            if (!FindTextForward(captionText))
                return false;

            for (int i = 0; i < 30; i++)
            {
                if (captionAboveTable)
                {
                    if (!TryAction("SelectCtrlFront"))
                        break;
                }
                else
                {
                    if (!TryAction("SelectCtrlReverse"))
                        break;
                }

                if (TryGetCurSelectedCtrlId(out string ctrlId) && ctrlId == "tbl")
                {
                    //TryAction("TableCellBlock");
                    return true;
                }
            }

            return false;
        }

        private bool FindTextForward(string text)
        {
            try
            {
                _hwp.HAction.GetDefault("ForwardFind", _hwp.HParameterSet.HFindReplace.HSet);
                _hwp.HParameterSet.HFindReplace.FindString = text;
                _hwp.HParameterSet.HFindReplace.ReplaceString = "";
                ApplyFindReplaceOptions(FindReplaceOptions.ForwardFindDefault);
                _hwp.HAction.Execute("ForwardFind", _hwp.HParameterSet.HFindReplace.HSet);
                return true;
            }
            catch
            {
                return false;
            }
        }

        private void ApplyFindReplaceOptions(FindReplaceOptions options)
        {
            _hwp.HParameterSet.HFindReplace.Direction = options.Direction;
            _hwp.HParameterSet.HFindReplace.MatchCase = options.MatchCase ? 1 : 0;
            _hwp.HParameterSet.HFindReplace.WholeWordOnly = options.WholeWordOnly ? 1 : 0;
            _hwp.HParameterSet.HFindReplace.FindRegExp = options.FindRegExp ? 1 : 0;
            _hwp.HParameterSet.HFindReplace.UseWildCards = options.UseWildCards ? 1 : 0;
            _hwp.HParameterSet.HFindReplace.SeveralWords = options.SeveralWords ? 1 : 0;
            _hwp.HParameterSet.HFindReplace.AllWordForms = options.AllWordForms ? 1 : 0;
            _hwp.HParameterSet.HFindReplace.AutoSpell = options.AutoSpell ? 1 : 0;
            _hwp.HParameterSet.HFindReplace.ReplaceMode = options.ReplaceMode ? 1 : 0;
            _hwp.HParameterSet.HFindReplace.IgnoreFindString = options.IgnoreFindString ? 1 : 0;
            _hwp.HParameterSet.HFindReplace.IgnoreReplaceString = options.IgnoreReplaceString ? 1 : 0;
            _hwp.HParameterSet.HFindReplace.IgnoreMessage = options.IgnoreMessage ? 1 : 0;
        }


        private void MoveToFirstRow()
        {
            // 맨 위/왼쪽 셀로 이동
            for (int i = 0; i < 50; i++) // 안전 루프
            {
                if (!TryAction("TableUpperCell")) break;
            }
        }

        private void MoveToFirstCell()
        {
            // 맨 위/왼쪽 셀로 이동
            for (int i = 0; i < 50; i++) // 안전 루프
            {
                if (!TryAction("TableUpperCell")) break;
            }
            RunActionAny("TableColBegin");
        }

        private void MoveToCell(int row, int col)
        {
            MoveToFirstRow();
            for (int r = 0; r < row; r++)
                RunActionAny("TableLowerCell");

            RunActionAny("TableColBegin");
            for (int c = 0; c < col; c++)
                RunActionAny("TableRightCell");
        }

        private bool MoveToTableByIndex(int index)
        {
            int count = 0;
            dynamic ctrl = _hwp.HeadCtrl;

            while (ctrl != null)
            {
                string id = null;

                try { id = ctrl.CtrlID as string; }
                catch { }

                if (id == "tbl")
                {
                    if (count == index)
                    {
                        var anchor = ctrl.GetAnchorPos(0);
                        _hwp.SetPosBySet(anchor);
                        _hwp.FindCtrl();

                        TryAction("ShapeObjTableSelCell");
                        return true;
                    }
                    count++;
                }

                try { ctrl = ctrl.Next; }
                catch { break; }
            }

            return false;
        }

        /// <summary>
        /// 현재 커서 위치 정보 가져오기
        /// </summary>
        /// <param name="cell"></param>
        /// <returns></returns>
        private int[] GetCursorInfo(string cell)
        {
            try
            {
                _hwp.KeyIndicator(out int seccnt, out int secno, out int prnpageno, out int colno, out int line, out int pos, out short over, out string ctrlname);
                int[] list = new int[6];
                list[0] = seccnt;               // 총 구역
                list[1] = secno;                // 구역 번호
                list[2] = prnpageno;            // 페이지 번호
                list[3] = colno;                // 다단 번호
                list[4] = line;                 // 줄 번호
                list[5] = pos;                  // 칸(컬럼) 번호
                                                // short ins_over = over;      // 삽입, 수정
                return list;
            }
            catch
            {
                return null;
            }
        }

        /// <summary>
        /// 현재 컨트롤 이름 가져오기
        /// </summary>
        /// <returns></returns>
        private string GetCtrlName()
        {
            try
            {
                _hwp.KeyIndicator(out int seccnt, out int secno, out int prnpageno, out int colno, out int line, out int pos, out short over, out string ctrlname);
                return ctrlname;
            }
            catch
            {
                return null;
            }
        }

        /// <summary>
        /// 현재 셀 주소 가져오기 (예: A1, E3 등)
        /// </summary>
        /// <returns></returns>
        private string GetCellAddr()
        {
            try
            {
                string cell = GetCtrlName();
                if (string.IsNullOrEmpty(cell))
                    return null;

                int open = cell.IndexOf('(');
                if (open < 0)
                    return null;

                int close = cell.IndexOf(')', open + 1);
                if (close < 0 || close <= open + 1)
                    return null;

                return cell.Substring(open + 1, close - open - 1);
            }
            catch 
            { 
                return null; 
            }
        }


        /// <summary>
        /// 표 셀 안에 커서가 위치하도록 시도
        /// </summary>
        /// <returns></returns>
        private bool EnsureCaretInTableCell()
        {
            try
            {
                //TryAction("ShapeObjTableSelCell")
                _hwp.Run("ShapeObjTableSelCell");
            }
            catch
            {
                return false;
            }

            return true;
        }

        private void ReplaceCurrentCellText(string value)
        {
            ClearCurrentCellText();
            InsertText(value);
        }

        private void ClearCurrentCellText()
        {
            TryAction("MoveLineBegin");
            TryAction("MoveSelLineEnd");
            TryAction("Delete");
            TryAction("Cancel");
        }

        private bool TryMoveToCtrl(dynamic ctrl)
        {
            // GetAnchorPos 인덱스를 0~2까지 시도
            for (int i = 0; i <= 2; i++)
            {
                if (TryMoveToCtrlByAnchor(ctrl, i))
                    return true;
            }
            return false;
        }

        private bool TryMoveToCtrlByAnchor(dynamic ctrl, int anchorIndex)
        {
            try
            {
                var ctrlType = ctrl.GetType();
                var anchor = ctrlType.InvokeMember(
                    "GetAnchorPos",
                    BindingFlags.InvokeMethod,
                    null,
                    ctrl,
                    new object[] { anchorIndex });

                if (anchor == null) return false;

                // 1) HwpPos 속성(List/Para/Pos) 우선
                if (TryGetAnchorPosProps(anchor, out int list, out int para, out int pos))
                {
                    if (list == 0 && para == 0 && pos == 0)
                        return false;

                    _hwp.SetPos(list, para, pos);
                    return true;
                }

                // 2) Item 인덱서(메서드/프로퍼티) 시도
                if (TryGetAnchorPosItem(anchor, 0, out int p0, out int p1, out int p2))
                {
                    _hwp.SetPos(p0, p1, p2);
                    return true;
                }

                if (TryGetAnchorPosItem(anchor, 1, out p0, out p1, out p2))
                {
                    _hwp.SetPos(p0, p1, p2);
                    return true;
                }

                return false;
            }
            catch
            {
                return false;
            }
        }

        private bool TryGetAnchorPosItem(object anchor, int startIndex, out int p0, out int p1, out int p2)
        {
            p0 = p1 = p2 = 0;

            try
            {
                var t = anchor.GetType();

                object o0 = TryInvokeItem(t, anchor, startIndex + 0);
                object o1 = TryInvokeItem(t, anchor, startIndex + 1);
                object o2 = TryInvokeItem(t, anchor, startIndex + 2);

                if (o0 == null || o1 == null || o2 == null)
                    return false;

                p0 = Convert.ToInt32(o0);
                p1 = Convert.ToInt32(o1);
                p2 = Convert.ToInt32(o2);

                return !(p0 == 0 && p1 == 0 && p2 == 0);
            }
            catch
            {
                return false;
            }
        }

        private object TryInvokeItem(Type t, object target, int index)
        {
            // 1) 메서드 Item(index)
            try
            {
                return t.InvokeMember("Item", BindingFlags.InvokeMethod, null, target, new object[] { index });
            }
            catch { }

            // 2) 프로퍼티 인덱서 Item[index]
            try
            {
                return t.InvokeMember("Item", BindingFlags.GetProperty, null, target, new object[] { index });
            }
            catch { }

            return null;
        }

        private bool TryGetAnchorPosProps(object anchor, out int list, out int para, out int pos)
        {
            list = para = pos = 0;
            try
            {
                var t = anchor.GetType();
                object oList = t.InvokeMember("List", BindingFlags.GetProperty, null, anchor, null);
                object oPara = t.InvokeMember("Para", BindingFlags.GetProperty, null, anchor, null);
                object oPos = t.InvokeMember("Pos", BindingFlags.GetProperty, null, anchor, null);

                if (oList == null || oPara == null || oPos == null)
                    return false;

                list = Convert.ToInt32(oList);
                para = Convert.ToInt32(oPara);
                pos = Convert.ToInt32(oPos);
                return true;
            }
            catch
            {
                return false;
            }
        }


        private bool TryGetCurSelectedCtrlId(out string ctrlId)
        {
            ctrlId = null;

            try
            {
                // dynamic 대신 reflection으로 안전하게 접근
                var hwpType = _hwp.GetType();
                var ctrl = hwpType.InvokeMember("CurSelectedCtrl",
                    BindingFlags.GetProperty, null, _hwp, null);

                if (ctrl == null) return false;

                var ctrlType = ctrl.GetType();
                ctrlId = ctrlType.InvokeMember("CtrlID",
                    BindingFlags.GetProperty, null, ctrl, null) as string;

                return !string.IsNullOrEmpty(ctrlId);
            }
            catch
            {
                return false;
            }
        }

        private void SetFontSize(int size)
        {
            _hwp.HAction.GetDefault("CharShape", _hwp.HParameterSet.HCharShape.HSet);
            _hwp.HParameterSet.HCharShape.Height = size * 100; // 1/100 pt
            _hwp.HAction.Execute("CharShape", _hwp.HParameterSet.HCharShape.HSet);
        }

        // 현재 줄 전체에 글자 크기 적용
        private void ApplyCharSizeToCurrentLine(int size)
        {
            // 현재 줄 전체 선택
            RunActionAny("MoveLineBegin");
            RunActionAny("Select");
            RunActionAny("MoveSelLineEnd");

            SetFontSize(size);

            // 줄 선택 해제
            RunActionAny("Cancel");
        }

        // 현재 문단 전체에 글자 크기 적용
        private void ApplyCharSizeToCurrentPara(int size)
        {
            // 현재 문단 전체 선택
            RunActionAny("MoveParaBegin");
            RunActionAny("Select");
            RunActionAny("MoveParaEnd");

            SetFontSize(size);

            // 문단 선택 해제
            RunActionAny("Cancel");
        }

        // 액션 실행 시도 (실패 시 예외 발생하지 않고 false 반환)
        private bool TryAction(string name)
        {
            try { return _hwp.HAction.Run(name); }
            catch { return false; }

        }

        // 액션명이 명확하지 않을 때 후보군을 순차 시도
        public void RunActionAny(params string[] names)
        {
            foreach (var name in names)
            {
                if (TryAction(name)) return;
            }
            throw new InvalidOperationException($"액션 실행 실패: {string.Join(", ", names)}");
        }

        // 셀 이동: 셀 오른쪽
        public void MoveNextCell()
        {
            TryAction("TableRightCell");
        }

        // 셀 이동: 셀 왼쪽
        public void MoveLeftCell()
        {
            TryAction("TableLeftCell");
        }

        // 셀 이동: 셀 위로
        public void MoveUpCell()
        {
            TryAction("TableUpperCell");
        }

        // 셀 이동: 셀 아래로
        public void MoveDownCell()
        {
            TryAction("TableLowerCell");
        }

        // 다음 행의 시작 셀로 이동
        private void MoveToNextRowStart()
        {
            TryAction("TableLowerCell");
            TryAction("TableColBegin");
        }

        // 문단 나누기
        private void InsertParagraph()
        {
            _hwp.HAction.Run("BreakPara");
        }

        private void EnsureCaretInBody()
        {
            // Clear selection and exit table/object mode if possible.
            TryAction("Cancel");
            TryAction("Close");
            TryAction("CloseEx");
            // Force caret into a guaranteed editable body position.
            TryAction("MoveDocBegin");
            TryAction("MoveDocEnd");
        }

        // 표 생성
        private void CreateTable(int rows, int cols)
        {
            if (rows <= 0 || cols <= 0)
                throw new ArgumentOutOfRangeException("rows/cols must be positive.");

            EnsureCaretInBody();

            try
            {
                _hwp.HAction.GetDefault("TableCreate", _hwp.HParameterSet.HTableCreation.HSet);

                _hwp.HParameterSet.HTableCreation.Rows = rows;
                _hwp.HParameterSet.HTableCreation.Cols = cols;

                _hwp.HParameterSet.HTableCreation.WidthType = 2;    // 너비를 내용에 맞게 자동 조절
                _hwp.HParameterSet.HTableCreation.HeightType = 0;   // 높이를 자동으로 조절

                _hwp.HAction.Execute("TableCreate", _hwp.HParameterSet.HTableCreation.HSet);
            }
            catch (Exception ex)
            {
                string ctrlName = GetCtrlName() ?? "?";
                string cellAddr = GetCellAddr() ?? "?";
                string selectedCtrlId = TryGetCurSelectedCtrlId(out string id) ? id : "?";
                Debug.WriteLine($"[HWP] TableCreate failed. rows={rows}, cols={cols}, ctrlName={ctrlName}, selectedCtrlId={selectedCtrlId}, cell={cellAddr}. {ex}");
                throw;
            }
        }

        // 표 만들기 (줄, 칸, 가로크기, 세로크기)
        public void CreateTable(int rows, int cols, double width, double height)
        {
            try
            {
                CreateTable(rows, cols);
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException("표 생성 실패", ex);
            }   
            //_hwp.HAction.GetDefault("TableCreate", _hwp.HParameterSet.HTableCreation.HSet);

            //_hwp.HParameterSet.HTableCreation.Rows = rows;
            //_hwp.HParameterSet.HTableCreation.Cols = cols;

            //_hwp.HParameterSet.HTableCreation.WidthType = 2; ;
            //_hwp.HParameterSet.HTableCreation.HeightType = 1; ;
            //_hwp.HParameterSet.HTableCreation.WidthValue = _hwp.MiliToHwpUnit(width);
            //_hwp.HParameterSet.HTableCreation.HeightValue = _hwp.MiliToHwpUnit(height);
            //_hwp.HParameterSet.HTableCreation.CreateItemArray("ColWidth", column);           // 칸 수만큼 배열 생성
            //for (int i = 0; i < column; i++)
            //{
            //    _hwp.HParameterSet.HTableCreation.ColWidth.item[i] = _hwp.MiliToHwpUnit((width / column) - 3.6);  //칸마다 폭 크기 설정      // 칸마다 왼쪽 여백 1.8,  오른쪽 여백 1.8을 빼줘야 함
            //}
            //_hwp.HParameterSet.HTableCreation.CreateItemArray("RowHeight", row); ;            // 행 수 만큼 배열 생성
            //for (int i = 0; i < row; i++)
            //{
            //    _hwp.HParameterSet.HTableCreation.RowHeight.item[i] = _hwp.MiliToHwpUnit((height / row) - 1.0);     // 행마다 높이 설정
            //}
            //_hwp.HParameterSet.HTableCreation.TableProperties.Width = _hwp.MiliToHwpUnit(width);
            //_hwp.HParameterSet.HTableCreation.TableProperties.OutsideMarginLeft = _hwp.MiliToHwpUnit(0);
            //_hwp.HParameterSet.HTableCreation.TableProperties.OutsideMarginRight = _hwp.MiliToHwpUnit(0);
            //_hwp.HParameterSet.HTableCreation.TableProperties.OutsideMarginTop = _hwp.MiliToHwpUnit(0);
            //_hwp.HParameterSet.HTableCreation.TableProperties.OutsideMarginBottom = _hwp.MiliToHwpUnit(0);
            //_hwp.HParameterSet.HTableCreation.TableProperties.TreatAsChar = 1;       //  # ;글자처럼 취급

            //_hwp.HAction.Execute("TableCreate", _hwp.HParameterSet.HTableCreation.HSet);
        }


        // 표 채우기
        private void FillTable(string[,] data)
        {
            int rows = data.GetLength(0);
            int cols = data.GetLength(1);

            for (int r = 0; r < rows; r++)
            {
                for (int c = 0; c < cols; c++)
                {
                    InsertText(data[r, c]);

                    if (c < cols - 1)
                        MoveNextCell();
                }

                if (r < rows - 1)
                    MoveToNextRowStart();
            }
        }

        public void MoveOutOfTable()
        {
            RunActionAny("Close", "CloseEx");
        }

        /// <summary>
        /// 배포용으로 저장
        /// </summary>
        /// <param name="filePath"></param>
        /// <param name="password"></param>
        /// <exception cref="InvalidOperationException"></exception>
        public void SaveAsDistribution(string filePath, string password = null)
        {
            SaveAs(filePath);

            var act = _hwp.CreateAction("FileSetSecurity");
            var set = act.CreateSet();
            if (act == null || set == null)
                throw new InvalidOperationException("보안 설정 액션 생성 실패");

            act.GetDefault(set);
            if (!string.IsNullOrEmpty(password) && password.Length > 4)
            {
                set.SetItem("Password", password);
            }
            set.SetItem("NoPrint", true);
            set.SetItem("NoCopy", true);
            act.Execute(set);
        }

        public sealed class FindReplaceOptions
        {
            public int Direction { get; set; } = 2; // 0=down, 1=up, 2=all
            public bool MatchCase { get; set; } = false;
            public bool WholeWordOnly { get; set; } = false;
            public bool FindRegExp { get; set; } = false;
            public bool UseWildCards { get; set; } = false;
            public bool SeveralWords { get; set; } = false;
            public bool AllWordForms { get; set; } = false;
            public bool AutoSpell { get; set; } = false;
            public bool ReplaceMode { get; set; } = false;
            public bool IgnoreFindString { get; set; } = false;
            public bool IgnoreReplaceString { get; set; } = false;
            public bool IgnoreMessage { get; set; } = true;

            public static FindReplaceOptions Default => new FindReplaceOptions
            {
                Direction = 2,
                MatchCase = false,
                WholeWordOnly = false,
                FindRegExp = false,
                UseWildCards = false,
                SeveralWords = false,
                AllWordForms = false,
                AutoSpell = false,
                ReplaceMode = false,
                IgnoreFindString = false,
                IgnoreReplaceString = false,
                IgnoreMessage = true
            };

            public static FindReplaceOptions ForwardFindDefault => new FindReplaceOptions
            {
                Direction = 0,
                MatchCase = false,
                WholeWordOnly = false,
                FindRegExp = false,
                UseWildCards = false,
                SeveralWords = false,
                AllWordForms = false,
                AutoSpell = false,
                ReplaceMode = false,
                IgnoreFindString = false,
                IgnoreReplaceString = true,
                IgnoreMessage = true
            };
        }
    }
}
