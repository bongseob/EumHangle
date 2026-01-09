namespace WindowsFormsApp1
{
    partial class HangulTest
    {
        /// <summary>
        /// 필수 디자이너 변수입니다.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 사용 중인 모든 리소스를 정리합니다.
        /// </summary>
        /// <param name="disposing">관리되는 리소스를 삭제해야 하면 true이고, 그렇지 않으면 false입니다.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form 디자이너에서 생성한 코드

        /// <summary>
        /// 디자이너 지원에 필요한 메서드입니다. 
        /// 이 메서드의 내용을 코드 편집기로 수정하지 마세요.
        /// </summary>
        private void InitializeComponent()
        {
            Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
            Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
            Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
            Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
            Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
            Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
            Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
            Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
            Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
            Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
            Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
            Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
            this.btnNewHwp = new System.Windows.Forms.Button();
            this.btnChangeHwp1 = new System.Windows.Forms.Button();
            this.button3 = new System.Windows.Forms.Button();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.btnSaveAsDistribution = new System.Windows.Forms.Button();
            this.btnSaveAsPdf = new System.Windows.Forms.Button();
            this.btnOpenDocument = new System.Windows.Forms.Button();
            this.btnNewDocument = new System.Windows.Forms.Button();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.btnInsertText = new System.Windows.Forms.Button();
            this.txtInsertText = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.btnReplaceAll = new System.Windows.Forms.Button();
            this.txtReplaceText = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.txtFindText = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.groupBox4 = new System.Windows.Forms.GroupBox();
            this.btnSetTableCell = new System.Windows.Forms.Button();
            this.txtCellText = new System.Windows.Forms.TextBox();
            this.label8 = new System.Windows.Forms.Label();
            this.txtCellCol = new System.Windows.Forms.TextBox();
            this.label7 = new System.Windows.Forms.Label();
            this.txtCellRow = new System.Windows.Forms.TextBox();
            this.label6 = new System.Windows.Forms.Label();
            this.txtTableIndex = new System.Windows.Forms.TextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.btnCreateTable = new System.Windows.Forms.Button();
            this.txtTableCols = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.txtTableRows = new System.Windows.Forms.TextBox();
            this.lblRows = new System.Windows.Forms.Label();
            this.groupBox5 = new System.Windows.Forms.GroupBox();
            this.btnInsertImage = new System.Windows.Forms.Button();
            this.saveFileDialog1 = new System.Windows.Forms.SaveFileDialog();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.button1 = new System.Windows.Forms.Button();
            this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.groupBox3.SuspendLayout();
            this.groupBox4.SuspendLayout();
            this.groupBox5.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
            this.SuspendLayout();
            // 
            // btnNewHwp
            // 
            this.btnNewHwp.Location = new System.Drawing.Point(6, 406);
            this.btnNewHwp.Name = "btnNewHwp";
            this.btnNewHwp.Size = new System.Drawing.Size(135, 32);
            this.btnNewHwp.TabIndex = 0;
            this.btnNewHwp.Text = "새로만들기";
            this.btnNewHwp.UseVisualStyleBackColor = true;
            this.btnNewHwp.Click += new System.EventHandler(this.btnNewHwp_Click);
            // 
            // btnChangeHwp1
            // 
            this.btnChangeHwp1.Location = new System.Drawing.Point(147, 406);
            this.btnChangeHwp1.Name = "btnChangeHwp1";
            this.btnChangeHwp1.Size = new System.Drawing.Size(135, 32);
            this.btnChangeHwp1.TabIndex = 1;
            this.btnChangeHwp1.Text = "수정하기1";
            this.btnChangeHwp1.UseVisualStyleBackColor = true;
            this.btnChangeHwp1.Click += new System.EventHandler(this.btnChangeHwp1_Click);
            // 
            // button3
            // 
            this.button3.Location = new System.Drawing.Point(288, 406);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(135, 32);
            this.button3.TabIndex = 1;
            this.button3.Text = "수정하기2";
            this.button3.UseVisualStyleBackColor = true;
            this.button3.Click += new System.EventHandler(this.button3_Click);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.btnSaveAsDistribution);
            this.groupBox1.Controls.Add(this.btnSaveAsPdf);
            this.groupBox1.Controls.Add(this.btnOpenDocument);
            this.groupBox1.Controls.Add(this.btnNewDocument);
            this.groupBox1.Location = new System.Drawing.Point(5, 6);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(250, 100);
            this.groupBox1.TabIndex = 5;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "파일기능";
            // 
            // btnSaveAsDistribution
            // 
            this.btnSaveAsDistribution.BackColor = System.Drawing.SystemColors.Highlight;
            this.btnSaveAsDistribution.FlatAppearance.BorderColor = System.Drawing.Color.Silver;
            this.btnSaveAsDistribution.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnSaveAsDistribution.ForeColor = System.Drawing.Color.White;
            this.btnSaveAsDistribution.Location = new System.Drawing.Point(130, 60);
            this.btnSaveAsDistribution.Name = "btnSaveAsDistribution";
            this.btnSaveAsDistribution.Size = new System.Drawing.Size(110, 23);
            this.btnSaveAsDistribution.TabIndex = 3;
            this.btnSaveAsDistribution.Text = "배포문서로저장";
            this.btnSaveAsDistribution.UseVisualStyleBackColor = false;
            this.btnSaveAsDistribution.Click += new System.EventHandler(this.btnSaveAsDistribution_Click);
            // 
            // btnSaveAsPdf
            // 
            this.btnSaveAsPdf.BackColor = System.Drawing.SystemColors.Highlight;
            this.btnSaveAsPdf.FlatAppearance.BorderColor = System.Drawing.Color.Silver;
            this.btnSaveAsPdf.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnSaveAsPdf.ForeColor = System.Drawing.Color.White;
            this.btnSaveAsPdf.Location = new System.Drawing.Point(9, 60);
            this.btnSaveAsPdf.Name = "btnSaveAsPdf";
            this.btnSaveAsPdf.Size = new System.Drawing.Size(110, 23);
            this.btnSaveAsPdf.TabIndex = 2;
            this.btnSaveAsPdf.Text = "PDF로 저장";
            this.btnSaveAsPdf.UseVisualStyleBackColor = false;
            this.btnSaveAsPdf.Click += new System.EventHandler(this.btnSaveAsPdf_Click);
            // 
            // btnOpenDocument
            // 
            this.btnOpenDocument.BackColor = System.Drawing.SystemColors.Highlight;
            this.btnOpenDocument.FlatAppearance.BorderColor = System.Drawing.Color.Silver;
            this.btnOpenDocument.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnOpenDocument.ForeColor = System.Drawing.Color.White;
            this.btnOpenDocument.Location = new System.Drawing.Point(130, 22);
            this.btnOpenDocument.Name = "btnOpenDocument";
            this.btnOpenDocument.Size = new System.Drawing.Size(110, 23);
            this.btnOpenDocument.TabIndex = 1;
            this.btnOpenDocument.Text = "문서열기";
            this.btnOpenDocument.UseVisualStyleBackColor = false;
            this.btnOpenDocument.Click += new System.EventHandler(this.btnOpenDocument_Click);
            // 
            // btnNewDocument
            // 
            this.btnNewDocument.BackColor = System.Drawing.SystemColors.Highlight;
            this.btnNewDocument.FlatAppearance.BorderColor = System.Drawing.Color.Silver;
            this.btnNewDocument.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnNewDocument.ForeColor = System.Drawing.Color.White;
            this.btnNewDocument.Location = new System.Drawing.Point(9, 22);
            this.btnNewDocument.Name = "btnNewDocument";
            this.btnNewDocument.Size = new System.Drawing.Size(110, 23);
            this.btnNewDocument.TabIndex = 0;
            this.btnNewDocument.Text = "새문서만들기";
            this.btnNewDocument.UseVisualStyleBackColor = false;
            this.btnNewDocument.Click += new System.EventHandler(this.btnNewHwp_Click);
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.btnInsertText);
            this.groupBox2.Controls.Add(this.txtInsertText);
            this.groupBox2.Controls.Add(this.label1);
            this.groupBox2.Location = new System.Drawing.Point(261, 6);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(250, 100);
            this.groupBox2.TabIndex = 6;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "편집하기";
            // 
            // btnInsertText
            // 
            this.btnInsertText.BackColor = System.Drawing.SystemColors.Highlight;
            this.btnInsertText.FlatAppearance.BorderColor = System.Drawing.Color.Silver;
            this.btnInsertText.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnInsertText.ForeColor = System.Drawing.Color.White;
            this.btnInsertText.Location = new System.Drawing.Point(157, 59);
            this.btnInsertText.Name = "btnInsertText";
            this.btnInsertText.Size = new System.Drawing.Size(75, 23);
            this.btnInsertText.TabIndex = 2;
            this.btnInsertText.Text = "문장추가";
            this.btnInsertText.UseVisualStyleBackColor = false;
            this.btnInsertText.Click += new System.EventHandler(this.btnInsertText_Click);
            // 
            // txtInsertText
            // 
            this.txtInsertText.Location = new System.Drawing.Point(18, 32);
            this.txtInsertText.Name = "txtInsertText";
            this.txtInsertText.Size = new System.Drawing.Size(214, 21);
            this.txtInsertText.TabIndex = 1;
            this.txtInsertText.Text = "Hello, HWP!";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(16, 17);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(85, 12);
            this.label1.TabIndex = 0;
            this.label1.Text = "추가할 문자열:";
            // 
            // groupBox3
            // 
            this.groupBox3.Controls.Add(this.btnReplaceAll);
            this.groupBox3.Controls.Add(this.txtReplaceText);
            this.groupBox3.Controls.Add(this.label3);
            this.groupBox3.Controls.Add(this.txtFindText);
            this.groupBox3.Controls.Add(this.label2);
            this.groupBox3.Location = new System.Drawing.Point(517, 6);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Size = new System.Drawing.Size(250, 130);
            this.groupBox3.TabIndex = 7;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = "찾기/바꾸기";
            // 
            // btnReplaceAll
            // 
            this.btnReplaceAll.BackColor = System.Drawing.SystemColors.Highlight;
            this.btnReplaceAll.FlatAppearance.BorderColor = System.Drawing.Color.Silver;
            this.btnReplaceAll.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnReplaceAll.ForeColor = System.Drawing.Color.White;
            this.btnReplaceAll.Location = new System.Drawing.Point(136, 98);
            this.btnReplaceAll.Name = "btnReplaceAll";
            this.btnReplaceAll.Size = new System.Drawing.Size(96, 23);
            this.btnReplaceAll.TabIndex = 4;
            this.btnReplaceAll.Text = "문자열바꾸기";
            this.btnReplaceAll.UseVisualStyleBackColor = false;
            this.btnReplaceAll.Click += new System.EventHandler(this.btnReplaceAll_Click);
            // 
            // txtReplaceText
            // 
            this.txtReplaceText.Location = new System.Drawing.Point(18, 71);
            this.txtReplaceText.Name = "txtReplaceText";
            this.txtReplaceText.Size = new System.Drawing.Size(214, 21);
            this.txtReplaceText.TabIndex = 3;
            this.txtReplaceText.Text = "Automation";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(16, 56);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(81, 12);
            this.label3.TabIndex = 2;
            this.label3.Text = "변경할문자열:";
            // 
            // txtFindText
            // 
            this.txtFindText.Location = new System.Drawing.Point(18, 32);
            this.txtFindText.Name = "txtFindText";
            this.txtFindText.Size = new System.Drawing.Size(214, 21);
            this.txtFindText.TabIndex = 1;
            this.txtFindText.Text = "HWP";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(16, 17);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(69, 12);
            this.label2.TabIndex = 0;
            this.label2.Text = "찾을문자열:";
            // 
            // groupBox4
            // 
            this.groupBox4.Controls.Add(this.btnSetTableCell);
            this.groupBox4.Controls.Add(this.txtCellText);
            this.groupBox4.Controls.Add(this.label8);
            this.groupBox4.Controls.Add(this.txtCellCol);
            this.groupBox4.Controls.Add(this.label7);
            this.groupBox4.Controls.Add(this.txtCellRow);
            this.groupBox4.Controls.Add(this.label6);
            this.groupBox4.Controls.Add(this.txtTableIndex);
            this.groupBox4.Controls.Add(this.label5);
            this.groupBox4.Controls.Add(this.btnCreateTable);
            this.groupBox4.Controls.Add(this.txtTableCols);
            this.groupBox4.Controls.Add(this.label4);
            this.groupBox4.Controls.Add(this.txtTableRows);
            this.groupBox4.Controls.Add(this.lblRows);
            this.groupBox4.Location = new System.Drawing.Point(5, 142);
            this.groupBox4.Name = "groupBox4";
            this.groupBox4.Size = new System.Drawing.Size(350, 180);
            this.groupBox4.TabIndex = 8;
            this.groupBox4.TabStop = false;
            this.groupBox4.Text = "표기능";
            // 
            // btnSetTableCell
            // 
            this.btnSetTableCell.BackColor = System.Drawing.SystemColors.Highlight;
            this.btnSetTableCell.FlatAppearance.BorderColor = System.Drawing.Color.Silver;
            this.btnSetTableCell.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnSetTableCell.ForeColor = System.Drawing.Color.White;
            this.btnSetTableCell.Location = new System.Drawing.Point(232, 144);
            this.btnSetTableCell.Name = "btnSetTableCell";
            this.btnSetTableCell.Size = new System.Drawing.Size(100, 23);
            this.btnSetTableCell.TabIndex = 13;
            this.btnSetTableCell.Text = "셀에값입력";
            this.btnSetTableCell.UseVisualStyleBackColor = false;
            this.btnSetTableCell.Click += new System.EventHandler(this.btnSetTableCell_Click);
            // 
            // txtCellText
            // 
            this.txtCellText.Location = new System.Drawing.Point(18, 146);
            this.txtCellText.Name = "txtCellText";
            this.txtCellText.Size = new System.Drawing.Size(208, 21);
            this.txtCellText.TabIndex = 12;
            this.txtCellText.Text = "New Value";
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(16, 131);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(60, 12);
            this.label8.TabIndex = 11;
            this.label8.Text = "Cell Text:";
            // 
            // txtCellCol
            // 
            this.txtCellCol.Location = new System.Drawing.Point(282, 107);
            this.txtCellCol.Name = "txtCellCol";
            this.txtCellCol.Size = new System.Drawing.Size(50, 21);
            this.txtCellCol.TabIndex = 10;
            this.txtCellCol.Text = "1";
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(230, 110);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(28, 12);
            this.label7.TabIndex = 9;
            this.label7.Text = "Col:";
            // 
            // txtCellRow
            // 
            this.txtCellRow.Location = new System.Drawing.Point(176, 107);
            this.txtCellRow.Name = "txtCellRow";
            this.txtCellRow.Size = new System.Drawing.Size(50, 21);
            this.txtCellRow.TabIndex = 8;
            this.txtCellRow.Text = "0";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(124, 110);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(34, 12);
            this.label6.TabIndex = 7;
            this.label6.Text = "Row:";
            // 
            // txtTableIndex
            // 
            this.txtTableIndex.Location = new System.Drawing.Point(70, 107);
            this.txtTableIndex.Name = "txtTableIndex";
            this.txtTableIndex.Size = new System.Drawing.Size(50, 21);
            this.txtTableIndex.TabIndex = 6;
            this.txtTableIndex.Text = "0";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(16, 110);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(52, 12);
            this.label5.TabIndex = 5;
            this.label5.Text = "T.Index:";
            // 
            // btnCreateTable
            // 
            this.btnCreateTable.BackColor = System.Drawing.SystemColors.Highlight;
            this.btnCreateTable.FlatAppearance.BorderColor = System.Drawing.Color.Silver;
            this.btnCreateTable.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnCreateTable.ForeColor = System.Drawing.Color.White;
            this.btnCreateTable.Location = new System.Drawing.Point(232, 30);
            this.btnCreateTable.Name = "btnCreateTable";
            this.btnCreateTable.Size = new System.Drawing.Size(100, 23);
            this.btnCreateTable.TabIndex = 4;
            this.btnCreateTable.Text = "표 만들기";
            this.btnCreateTable.UseVisualStyleBackColor = false;
            this.btnCreateTable.Click += new System.EventHandler(this.btnCreateTable_Click);
            // 
            // txtTableCols
            // 
            this.txtTableCols.Location = new System.Drawing.Point(176, 32);
            this.txtTableCols.Name = "txtTableCols";
            this.txtTableCols.Size = new System.Drawing.Size(50, 21);
            this.txtTableCols.TabIndex = 3;
            this.txtTableCols.Text = "4";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(124, 35);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(35, 12);
            this.label4.TabIndex = 2;
            this.label4.Text = "Cols:";
            // 
            // txtTableRows
            // 
            this.txtTableRows.Location = new System.Drawing.Point(70, 32);
            this.txtTableRows.Name = "txtTableRows";
            this.txtTableRows.Size = new System.Drawing.Size(50, 21);
            this.txtTableRows.TabIndex = 1;
            this.txtTableRows.Text = "3";
            // 
            // lblRows
            // 
            this.lblRows.AutoSize = true;
            this.lblRows.Location = new System.Drawing.Point(16, 35);
            this.lblRows.Name = "lblRows";
            this.lblRows.Size = new System.Drawing.Size(41, 12);
            this.lblRows.TabIndex = 0;
            this.lblRows.Text = "Rows:";
            // 
            // groupBox5
            // 
            this.groupBox5.Controls.Add(this.btnInsertImage);
            this.groupBox5.Location = new System.Drawing.Point(361, 142);
            this.groupBox5.Name = "groupBox5";
            this.groupBox5.Size = new System.Drawing.Size(250, 60);
            this.groupBox5.TabIndex = 9;
            this.groupBox5.TabStop = false;
            this.groupBox5.Text = "그림기능";
            // 
            // btnInsertImage
            // 
            this.btnInsertImage.BackColor = System.Drawing.SystemColors.Highlight;
            this.btnInsertImage.FlatAppearance.BorderColor = System.Drawing.Color.Silver;
            this.btnInsertImage.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnInsertImage.ForeColor = System.Drawing.Color.White;
            this.btnInsertImage.Location = new System.Drawing.Point(9, 22);
            this.btnInsertImage.Name = "btnInsertImage";
            this.btnInsertImage.Size = new System.Drawing.Size(110, 23);
            this.btnInsertImage.TabIndex = 0;
            this.btnInsertImage.Text = "그림추가";
            this.btnInsertImage.UseVisualStyleBackColor = false;
            this.btnInsertImage.Click += new System.EventHandler(this.btnInsertImage_Click);
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // button1
            // 
            this.button1.BackColor = System.Drawing.SystemColors.Highlight;
            this.button1.FlatAppearance.BorderColor = System.Drawing.Color.Silver;
            this.button1.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button1.ForeColor = System.Drawing.Color.White;
            this.button1.Location = new System.Drawing.Point(640, 164);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(100, 23);
            this.button1.TabIndex = 14;
            this.button1.Text = "일일재정현황";
            this.button1.UseVisualStyleBackColor = false;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // ultraGrid1
            // 
            appearance1.BackColor = System.Drawing.SystemColors.Window;
            appearance1.BorderColor = System.Drawing.SystemColors.InactiveCaption;
            this.ultraGrid1.DisplayLayout.Appearance = appearance1;
            this.ultraGrid1.DisplayLayout.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
            this.ultraGrid1.DisplayLayout.CaptionVisible = Infragistics.Win.DefaultableBoolean.False;
            appearance2.BackColor = System.Drawing.SystemColors.ActiveBorder;
            appearance2.BackColor2 = System.Drawing.SystemColors.ControlDark;
            appearance2.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
            appearance2.BorderColor = System.Drawing.SystemColors.Window;
            this.ultraGrid1.DisplayLayout.GroupByBox.Appearance = appearance2;
            appearance3.ForeColor = System.Drawing.SystemColors.GrayText;
            this.ultraGrid1.DisplayLayout.GroupByBox.BandLabelAppearance = appearance3;
            this.ultraGrid1.DisplayLayout.GroupByBox.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
            appearance4.BackColor = System.Drawing.SystemColors.ControlLightLight;
            appearance4.BackColor2 = System.Drawing.SystemColors.Control;
            appearance4.BackGradientStyle = Infragistics.Win.GradientStyle.Horizontal;
            appearance4.ForeColor = System.Drawing.SystemColors.GrayText;
            this.ultraGrid1.DisplayLayout.GroupByBox.PromptAppearance = appearance4;
            this.ultraGrid1.DisplayLayout.MaxColScrollRegions = 1;
            this.ultraGrid1.DisplayLayout.MaxRowScrollRegions = 1;
            appearance5.BackColor = System.Drawing.SystemColors.Window;
            appearance5.ForeColor = System.Drawing.SystemColors.ControlText;
            this.ultraGrid1.DisplayLayout.Override.ActiveCellAppearance = appearance5;
            appearance6.BackColor = System.Drawing.SystemColors.Highlight;
            appearance6.ForeColor = System.Drawing.SystemColors.HighlightText;
            this.ultraGrid1.DisplayLayout.Override.ActiveRowAppearance = appearance6;
            this.ultraGrid1.DisplayLayout.Override.BorderStyleCell = Infragistics.Win.UIElementBorderStyle.Dotted;
            this.ultraGrid1.DisplayLayout.Override.BorderStyleRow = Infragistics.Win.UIElementBorderStyle.Dotted;
            appearance7.BackColor = System.Drawing.SystemColors.Window;
            this.ultraGrid1.DisplayLayout.Override.CardAreaAppearance = appearance7;
            appearance8.BorderColor = System.Drawing.Color.Silver;
            appearance8.TextTrimming = Infragistics.Win.TextTrimming.EllipsisCharacter;
            this.ultraGrid1.DisplayLayout.Override.CellAppearance = appearance8;
            this.ultraGrid1.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.EditAndSelectText;
            this.ultraGrid1.DisplayLayout.Override.CellPadding = 0;
            appearance9.BackColor = System.Drawing.SystemColors.Control;
            appearance9.BackColor2 = System.Drawing.SystemColors.ControlDark;
            appearance9.BackGradientAlignment = Infragistics.Win.GradientAlignment.Element;
            appearance9.BackGradientStyle = Infragistics.Win.GradientStyle.Horizontal;
            appearance9.BorderColor = System.Drawing.SystemColors.Window;
            this.ultraGrid1.DisplayLayout.Override.GroupByRowAppearance = appearance9;
            appearance10.TextHAlignAsString = "Left";
            this.ultraGrid1.DisplayLayout.Override.HeaderAppearance = appearance10;
            this.ultraGrid1.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
            this.ultraGrid1.DisplayLayout.Override.HeaderStyle = Infragistics.Win.HeaderStyle.WindowsXPCommand;
            appearance11.BackColor = System.Drawing.SystemColors.Window;
            appearance11.BorderColor = System.Drawing.Color.Silver;
            this.ultraGrid1.DisplayLayout.Override.RowAppearance = appearance11;
            this.ultraGrid1.DisplayLayout.Override.RowSelectors = Infragistics.Win.DefaultableBoolean.False;
            appearance12.BackColor = System.Drawing.SystemColors.ControlLight;
            this.ultraGrid1.DisplayLayout.Override.TemplateAddRowAppearance = appearance12;
            this.ultraGrid1.DisplayLayout.ScrollBounds = Infragistics.Win.UltraWinGrid.ScrollBounds.ScrollToFill;
            this.ultraGrid1.DisplayLayout.ScrollStyle = Infragistics.Win.UltraWinGrid.ScrollStyle.Immediate;
            this.ultraGrid1.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy;
            this.ultraGrid1.Location = new System.Drawing.Point(361, 218);
            this.ultraGrid1.Name = "ultraGrid1";
            this.ultraGrid1.Size = new System.Drawing.Size(406, 182);
            this.ultraGrid1.TabIndex = 15;
            this.ultraGrid1.Text = "ultraGrid1";
            // 
            // HangulTest
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(782, 450);
            this.Controls.Add(this.ultraGrid1);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.groupBox3);
            this.Controls.Add(this.groupBox4);
            this.Controls.Add(this.groupBox5);
            this.Controls.Add(this.button3);
            this.Controls.Add(this.btnChangeHwp1);
            this.Controls.Add(this.btnNewHwp);
            this.Name = "HangulTest";
            this.Text = "한글연동 테스트";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.groupBox1.ResumeLayout(false);
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.groupBox3.ResumeLayout(false);
            this.groupBox3.PerformLayout();
            this.groupBox4.ResumeLayout(false);
            this.groupBox4.PerformLayout();
            this.groupBox5.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button btnNewHwp;
        private System.Windows.Forms.Button btnChangeHwp1;
        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Button btnSaveAsDistribution;
        private System.Windows.Forms.Button btnSaveAsPdf;
        private System.Windows.Forms.Button btnOpenDocument;
        private System.Windows.Forms.Button btnNewDocument;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.Button btnInsertText;
        private System.Windows.Forms.TextBox txtInsertText;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.GroupBox groupBox3;
        private System.Windows.Forms.Button btnReplaceAll;
        private System.Windows.Forms.TextBox txtReplaceText;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox txtFindText;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.GroupBox groupBox4;
        private System.Windows.Forms.Button btnSetTableCell;
        private System.Windows.Forms.TextBox txtCellText;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.TextBox txtCellCol;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.TextBox txtCellRow;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.TextBox txtTableIndex;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Button btnCreateTable;
        private System.Windows.Forms.TextBox txtTableCols;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox txtTableRows;
        private System.Windows.Forms.Label lblRows;
        private System.Windows.Forms.GroupBox groupBox5;
        private System.Windows.Forms.Button btnInsertImage;
        private System.Windows.Forms.SaveFileDialog saveFileDialog1;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.Button button1;
        private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
    }
}

