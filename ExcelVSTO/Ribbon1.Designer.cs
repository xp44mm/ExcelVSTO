namespace ExcelVSTO
{
    partial class Ribbon1 : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Ribbon1()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// 清理所有正在使用的资源。
        /// </summary>
        /// <param name="disposing">如果应释放托管资源，为 true；否则为 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region 组件设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            this.tab1 = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.btnSelectInUsedRange = this.Factory.CreateRibbonButton();
            this.btnSelectArray = this.Factory.CreateRibbonButton();
            this.btnMergeCells = this.Factory.CreateRibbonButton();
            this.group2 = this.Factory.CreateRibbonGroup();
            this.btnUnroll = this.Factory.CreateRibbonButton();
            this.btnRollup = this.Factory.CreateRibbonButton();
            this.btnAlternateRows = this.Factory.CreateRibbonButton();
            this.btnInsertBlank = this.Factory.CreateRibbonButton();
            this.btnRemoveBlank = this.Factory.CreateRibbonButton();
            this.group3 = this.Factory.CreateRibbonGroup();
            this.btnIncrease = this.Factory.CreateRibbonButton();
            this.btnDecrease = this.Factory.CreateRibbonButton();
            this.group5 = this.Factory.CreateRibbonGroup();
            this.btnBisect = this.Factory.CreateRibbonButton();
            this.btnSuccessive = this.Factory.CreateRibbonButton();
            this.group4 = this.Factory.CreateRibbonGroup();
            this.deprecatedFormulae = this.Factory.CreateRibbonButton();
            this.clearName_button = this.Factory.CreateRibbonButton();
            this.group6 = this.Factory.CreateRibbonGroup();
            this.btn_referencesOfWorksheet = this.Factory.CreateRibbonButton();
            this.btn_dependentsOfWorksheet = this.Factory.CreateRibbonButton();
            this.group7 = this.Factory.CreateRibbonGroup();
            this.btn_RenderFSharp = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            this.group2.SuspendLayout();
            this.group3.SuspendLayout();
            this.group5.SuspendLayout();
            this.group4.SuspendLayout();
            this.group6.SuspendLayout();
            this.group7.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.group1);
            this.tab1.Groups.Add(this.group2);
            this.tab1.Groups.Add(this.group3);
            this.tab1.Groups.Add(this.group5);
            this.tab1.Groups.Add(this.group4);
            this.tab1.Groups.Add(this.group6);
            this.tab1.Groups.Add(this.group7);
            this.tab1.Label = "TabAddIns";
            this.tab1.Name = "tab1";
            // 
            // group1
            // 
            this.group1.Items.Add(this.btnSelectInUsedRange);
            this.group1.Items.Add(this.btnSelectArray);
            this.group1.Items.Add(this.btnMergeCells);
            this.group1.Label = "选择工具";
            this.group1.Name = "group1";
            // 
            // btnSelectInUsedRange
            // 
            this.btnSelectInUsedRange.Label = "缩小选择";
            this.btnSelectInUsedRange.Name = "btnSelectInUsedRange";
            this.btnSelectInUsedRange.ScreenTip = "将整行整列的选择限定在已使用的范围内";
            this.btnSelectInUsedRange.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnSelectInUsedRange_Click);
            // 
            // btnSelectArray
            // 
            this.btnSelectArray.Label = "选择数组";
            this.btnSelectArray.Name = "btnSelectArray";
            this.btnSelectArray.ScreenTip = "扩展当前单元格为数组公式的范围";
            this.btnSelectArray.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnSelectArray_Click);
            // 
            // btnMergeCells
            // 
            this.btnMergeCells.Label = "合并单元格";
            this.btnMergeCells.Name = "btnMergeCells";
            this.btnMergeCells.SuperTip = "合并单元格及其内容";
            this.btnMergeCells.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnMergeCells_Click);
            // 
            // group2
            // 
            this.group2.Items.Add(this.btnUnroll);
            this.group2.Items.Add(this.btnRollup);
            this.group2.Items.Add(this.btnAlternateRows);
            this.group2.Items.Add(this.btnInsertBlank);
            this.group2.Items.Add(this.btnRemoveBlank);
            this.group2.Label = "表格工具";
            this.group2.Name = "group2";
            // 
            // btnUnroll
            // 
            this.btnUnroll.Label = "展开列";
            this.btnUnroll.Name = "btnUnroll";
            this.btnUnroll.ScreenTip = "unroll";
            this.btnUnroll.SuperTip = "填充下方的空单元格";
            this.btnUnroll.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnUnroll_Click);
            // 
            // btnRollup
            // 
            this.btnRollup.Label = "卷起列";
            this.btnRollup.Name = "btnRollup";
            this.btnRollup.ScreenTip = "rollup";
            this.btnRollup.SuperTip = "合并相同值到最顶单元格";
            this.btnRollup.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnRollup_Click);
            // 
            // btnAlternateRows
            // 
            this.btnAlternateRows.Label = "交替行";
            this.btnAlternateRows.Name = "btnAlternateRows";
            this.btnAlternateRows.SuperTip = "交替行着色，取第一个单元格颜色";
            this.btnAlternateRows.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnAlternateRows_Click);
            // 
            // btnInsertBlank
            // 
            this.btnInsertBlank.Label = "插入空行";
            this.btnInsertBlank.Name = "btnInsertBlank";
            this.btnInsertBlank.SuperTip = "在不同类别的行之间插入空行";
            this.btnInsertBlank.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnInsertBlank_Click);
            // 
            // btnRemoveBlank
            // 
            this.btnRemoveBlank.Label = "删除空行";
            this.btnRemoveBlank.Name = "btnRemoveBlank";
            this.btnRemoveBlank.SuperTip = "删除空行";
            this.btnRemoveBlank.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnRemoveBlank_Click);
            // 
            // group3
            // 
            this.group3.Items.Add(this.btnIncrease);
            this.group3.Items.Add(this.btnDecrease);
            this.group3.Label = "数学工具";
            this.group3.Name = "group3";
            // 
            // btnIncrease
            // 
            this.btnIncrease.Label = "自加一";
            this.btnIncrease.Name = "btnIncrease";
            this.btnIncrease.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnIncrease_Click);
            // 
            // btnDecrease
            // 
            this.btnDecrease.Label = "自减一";
            this.btnDecrease.Name = "btnDecrease";
            this.btnDecrease.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnDecrease_Click);
            // 
            // group5
            // 
            this.group5.Items.Add(this.btnBisect);
            this.group5.Items.Add(this.btnSuccessive);
            this.group5.Label = "方程求根";
            this.group5.Name = "group5";
            // 
            // btnBisect
            // 
            this.btnBisect.Label = "对分法归零";
            this.btnBisect.Name = "btnBisect";
            this.btnBisect.ScreenTip = "选中单元格=(A+B)/2";
            this.btnBisect.SuperTip = "选中单元格的R[1]为目标单元格，其值小于零修改A，大于零修改B";
            this.btnBisect.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnBisect_Click);
            // 
            // btnSuccessive
            // 
            this.btnSuccessive.Label = "代入法归零";
            this.btnSuccessive.Name = "btnSuccessive";
            this.btnSuccessive.ScreenTip = "选中单元格=A-B";
            this.btnSuccessive.SuperTip = "do B <- A";
            this.btnSuccessive.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnSuccessive_Click);
            // 
            // group4
            // 
            this.group4.Items.Add(this.deprecatedFormulae);
            this.group4.Items.Add(this.clearName_button);
            this.group4.Label = "公式";
            this.group4.Name = "group4";
            // 
            // deprecatedFormulae
            // 
            this.deprecatedFormulae.Label = "不支持的公式";
            this.deprecatedFormulae.Name = "deprecatedFormulae";
            this.deprecatedFormulae.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.deprecatedFormulae_Click);
            // 
            // clearName_button
            // 
            this.clearName_button.Label = "清除名称";
            this.clearName_button.Name = "clearName_button";
            this.clearName_button.ScreenTip = "清除单元格对名称的引用";
            this.clearName_button.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.clearName_button_Click);
            // 
            // group6
            // 
            this.group6.Items.Add(this.btn_referencesOfWorksheet);
            this.group6.Items.Add(this.btn_dependentsOfWorksheet);
            this.group6.Label = "工作表";
            this.group6.Name = "group6";
            // 
            // btn_referencesOfWorksheet
            // 
            this.btn_referencesOfWorksheet.Label = "工作表引用";
            this.btn_referencesOfWorksheet.Name = "btn_referencesOfWorksheet";
            this.btn_referencesOfWorksheet.Tag = "工作表输入";
            this.btn_referencesOfWorksheet.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_referencesOfWorksheet_Click);
            // 
            // btn_dependentsOfWorksheet
            // 
            this.btn_dependentsOfWorksheet.Label = "工作表依赖";
            this.btn_dependentsOfWorksheet.Name = "btn_dependentsOfWorksheet";
            this.btn_dependentsOfWorksheet.Tag = "工作表输出";
            this.btn_dependentsOfWorksheet.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_dependentsOfWorksheet_Click);
            // 
            // group7
            // 
            this.group7.Items.Add(this.btn_RenderFSharp);
            this.group7.Label = "渲染";
            this.group7.Name = "group7";
            // 
            // btn_RenderFSharp
            // 
            this.btn_RenderFSharp.Label = "渲染FSharp";
            this.btn_RenderFSharp.Name = "btn_RenderFSharp";
            this.btn_RenderFSharp.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_RenderFSharp_Click);
            // 
            // Ribbon1
            // 
            this.Name = "Ribbon1";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.group2.ResumeLayout(false);
            this.group2.PerformLayout();
            this.group3.ResumeLayout(false);
            this.group3.PerformLayout();
            this.group5.ResumeLayout(false);
            this.group5.PerformLayout();
            this.group4.ResumeLayout(false);
            this.group4.PerformLayout();
            this.group6.ResumeLayout(false);
            this.group6.PerformLayout();
            this.group7.ResumeLayout(false);
            this.group7.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnSelectInUsedRange;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnSelectArray;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnMergeCells;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnUnroll;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnRollup;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnInsertBlank;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnRemoveBlank;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnAlternateRows;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnIncrease;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnDecrease;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnSuccessive;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnBisect;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group2;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group3;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group4;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton deprecatedFormulae;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group5;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton clearName_button;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group6;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn_referencesOfWorksheet;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn_dependentsOfWorksheet;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group7;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn_RenderFSharp;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon1 Ribbon1
        {
            get { return this.GetRibbon<Ribbon1>(); }
        }
    }
}
