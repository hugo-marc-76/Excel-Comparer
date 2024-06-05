using ClosedXML.Excel;

namespace FileContainSameData__For_Excel_File_
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            this.AllowDrop = true;
            this.DragEnter += new DragEventHandler(Form1_DragEnter);
            this.DragDrop += new DragEventHandler(Form1_DragDrop);
        }

        private void Form1_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effect = DragDropEffects.Copy;
            }
        }

        private void Form1_DragDrop(object sender, DragEventArgs e)
        {
            string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
            if (files.Length == 2)
            {
                CompareExcelFiles(files[0], files[1]);
            }
            else
            {
                MessageBox.Show("Please drop exactly 2 files.");
            }
        }

        private void CompareExcelFiles(string filePath1, string filePath2)
        {
            if (!File.Exists(filePath1) || !File.Exists(filePath2))
            {
                MessageBox.Show("One or both of the specified files do not exist.");
                return;
            }

            try
            {
                using (var workbook1 = new XLWorkbook(filePath1))
                using (var workbook2 = new XLWorkbook(filePath2))
                {
                    bool areIdentical = CompareWorkbooks(workbook1, workbook2);
                    if (areIdentical)
                    {
                        MessageBox.Show("The files are identical.");
                    }
                    else
                    {
                        MessageBox.Show("The files are different.");
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"An error occurred: {ex.Message}");
            }
        }

        private bool CompareWorkbooks(XLWorkbook workbook1, XLWorkbook workbook2)
        {
            if (workbook1.Worksheets.Count != workbook2.Worksheets.Count)
                return false;

            for (int i = 1; i <= workbook1.Worksheets.Count; i++)
            {
                var ws1 = workbook1.Worksheet(i);
                var ws2 = workbook2.Worksheet(i);

                if (!CompareWorksheets(ws1, ws2))
                    return false;
            }

            return true;
        }

        private bool CompareWorksheets(IXLWorksheet ws1, IXLWorksheet ws2)
        {
            var range1 = ws1.RangeUsed();
            var range2 = ws2.RangeUsed();

            if (range1.RowCount() != range2.RowCount() || range1.ColumnCount() != range2.ColumnCount())
                return false;

            foreach (var cell1 in range1.Cells())
            {
                var cell2 = range2.Cell(cell1.Address.RowNumber, cell1.Address.ColumnNumber);
                if (cell1.Value.ToString() != cell2.Value.ToString())
                    return false;
            }

            return true;
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }
    }
}
