using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.Windows.Forms;
using TableDefinitionCreator.dto;
using TableDefinitionCreator.enums;
using TableDefinitionCreator.utils;

namespace TableDefinitionCreator
{
    public partial class Form1 : Form
    {
        private readonly BindingList<DataTable> _tableList = new BindingList<DataTable>();

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            InitialBindingList();

            bool isConnectionStringInitialized = IsConnectionStringInitialized();
            if (isConnectionStringInitialized == false)
            {
                SetupConnectionString();
            }
           
            this.SetDoubleBuffered(true, true);
        }

        private void InitialBindingList()
        {
            lbTables.DataSource = _tableList;
            lbTables.DisplayMember = "TableName";
        }

        private void SetupConnectionString()
        {
            using (var setupDialog = new ConnectionSetupDialog())
            {
                setupDialog.ShowDialog();
                if (IsConnectionStringInitialized() == false)
                {
                    MessageBox.Show("데이터베이스 연결 설정이 필요합니다.\n프로그램을 종료합니다.", "설정 필요", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    this.Close();
                }
            }
        }

        private bool IsConnectionStringInitialized()
        {
            return ConfigManager.GetConnectionString() != null && DbAccess.IsConn(out _);
        }

        private void txtTableName_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                btnSearch.PerformClick();
            }
        }

        private void btnSearch_Click(object sender, EventArgs e)
        {
            string input = txtTableName.Text.Trim();
            if (string.IsNullOrEmpty(input))
                return;

            //첫글자 대문자로
            input = char.ToUpper(input[0]) + input.Substring(1);

            ClearSearchResult();

            // 이미 테이블리스트에 있으면 가져오기
            DataTable inListDt = _tableList.FirstOrDefault(it => it.TableName.Equals(input, StringComparison.OrdinalIgnoreCase));
            if (inListDt != null)
            {
                BindingDataTable(inListDt);
                return;
            }

            DataTable dt = QueryManager.GetTableDefinitionList(input).FirstOrDefault();
            if (dt == null || dt.Rows.Count == 0)
            {
                ShowError("쿼리의 결과가 없습니다. 테이블명을 다시 확인해주세요.");
                return;
            }

            BindingDataTable(dt);
        }

        /// <summary>
        /// (현재 사용안함) DataGridView를 TSV 혹은 CSV로 Clipboard에 복사
        /// </summary>
        private void btnCopy_Click(object sender, EventArgs e)
        {
            if((dgvResult.DataSource is DataTable dt) == false || dt.Rows.Count == 0)
            {
                MessageBox.Show("복사할 데이터가 없습니다.", "데이터 없음", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            dt.CopyToClipBoard(CopyType.TSV, false); // Or CopyType.CSV
        }

        private void 연결설정ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SetupConnectionString();
        }

        private void btnExportExcel_Click(object sender, EventArgs e)
        {
            if (_tableList.Count == 0)
            {
                MessageBox.Show("내보낼 테이블이 없습니다.", "테이블 없음", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            try
            {
                string filePath = ShowSaveFileDialog(SaveType.EXCEL);
                if (string.IsNullOrEmpty(filePath))
                    return;
                this.Cursor = Cursors.WaitCursor;


                ExcelManager.ExportToExcel(_tableList.ToList(), filePath, true);
                
                DialogResult dialogResult = MessageBox.Show("내보내기가 완료되었습니다.\r\n파일을 여시겠습니까?", "완료", 
                    MessageBoxButtons.YesNo, MessageBoxIcon.Information);
                
                if (dialogResult == DialogResult.Yes)
                {
                    Process.Start(filePath);
                }
            }
            catch (Exception ex)
            {
                ShowError(ex.Message);
            }
            finally
            {
                this.Cursor = Cursors.Default;
            }
        }

        private void btnAddToList_Click(object sender, EventArgs e)
        {
            DataTable currentDt = (dgvResult.DataSource as DataTable);
            if (currentDt == null || currentDt.Rows.Count == 0)
            {
                MessageBox.Show("추가할 데이터가 없습니다.", "데이터 없음", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            // 기존에 있으면 제거 후 다시 추가
            if (_tableList.Contains(currentDt))
            {
                _tableList.Remove(currentDt); 
            }

            currentDt.ExtendedProperties["Remark"] = richTxtRemark.Text;
            _tableList.Insert(0, currentDt);
            lbTables.SelectedItem = currentDt;
            ClearSearchResult();
        }

        private void ClearSearchResult()
        {
            dgvResult.DataSource = null;
            lblTableName.Text = string.Empty;
            lblTableDesc.Text = string.Empty;
            richTxtRemark.Clear();
            txtTableName.Clear();

            txtTableName.Focus();
        }

        private void ClearBindingList()
        {
            _tableList.Clear();
        }

        private void ReplaceBindingList(List<DataTable> dts)
        {
            _tableList.Clear();
            dts.ForEach(dt => _tableList.Add(dt));
        }

        private void ShowError(string msg)
        {
            MessageBox.Show(msg, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        private string ShowSaveFileDialog(SaveType saveType)
        {
            using (var dialog = new SaveFileDialog())
            {
                switch (saveType)
                {
                    case SaveType.EXCEL:
                        dialog.Filter = "엑셀 파일 (*.xlsx)|*.xlsx";
                        dialog.Title = "EXCEL";
                        dialog.FileName = $"SF-TD. 테이블 정의서_{DateTime.Now:yyyyMMdd}.xlsx";
                        break;
                    case SaveType.JSON:
                        dialog.Filter = "JSON 파일 (*.json)|*.json";
                        dialog.Title = "JSON";
                        dialog.FileName = "tables.json";
                        break;

                    case SaveType.HTML:
                        dialog.Filter = "HTML 파일 (*.html)|*.html";
                        dialog.Title = "HTML";
                        dialog.FileName = $"SF-TD. 테이블 정의서_{DateTime.Now:yyyyMMdd}.html";
                        break;
                    default:
                        throw new ArgumentOutOfRangeException(nameof(saveType), saveType, null);
                }
               

                var dialogResult = dialog.ShowDialog();
                return dialogResult == DialogResult.OK ? dialog.FileName : string.Empty;
            }
        }

        private string ShowOpenFileDialog()
        {
            using (var dialog = new OpenFileDialog())
            {
                dialog.Filter = "JSON File (*.json)|*.json";
                dialog.Title = "JSON";
                var dialogResult = dialog.ShowDialog();

                return dialogResult == DialogResult.OK ? dialog.FileName : string.Empty;
            }
        }

        private void btnExcludeTable_Click(object sender, EventArgs e)
        {
            if (lbTables.SelectedItems.Count == 0)
            {
                return;
            }

            var selectedItems = lbTables.SelectedItems.Cast<DataTable>().ToList();

            string msg = selectedItems.Count == 1 
                ? $"{selectedItems[0].TableName} 테이블을 목록에서 제외하시겠습니까?" 
                : $"{selectedItems.Count}개의 테이블을 목록에서 제외하시겠습니까?";

            var dialogResult = MessageBox.Show(
                msg, 
                "제외",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question);

            if (dialogResult == DialogResult.Yes)
            {
                selectedItems.ForEach(item => _tableList.Remove(item));
            }
        }

        private void btnLoadFile_Click(object sender, EventArgs e)
        {
            try
            {
                var jsonFileName = ShowOpenFileDialog();
                if (string.IsNullOrEmpty(jsonFileName))
                    return;

                this.Cursor = Cursors.WaitCursor;

                List<TableInformation> parsedList = JsonManager.FromJson(jsonFileName);
                List<DataTable> dtList = QueryManager.GetTableDefinitionList(parsedList);

                ReplaceBindingList(dtList);
                ClearSearchResult();

                MessageBox.Show("불러오기가 완료되었습니다.", "완료", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                ShowError(ex.Message);
            }
            finally
            {
                this.Cursor = Cursors.Default;
            }
        }

        private void btnSaveFile_Click(object sender, EventArgs e)
        {
            if (_tableList.Count == 0)
            {
                MessageBox.Show("저장할 테이블이 없습니다.", "테이블 없음", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            try
            {
                string filepath = ShowSaveFileDialog(SaveType.JSON);
                if (string.IsNullOrEmpty(filepath))
                    return;

                this.Cursor = Cursors.WaitCursor;
                JsonManager.SaveToJson(_tableList.ToList(), filepath);
                MessageBox.Show("저장이 완료되었습니다.", "완료", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                ShowError(ex.Message);
            }
            finally
            {
                this.Cursor = Cursors.Default;
            }
        }

        private void lbTables_DoubleClick(object sender, EventArgs e)
        {
            DataTable selectedDt = (lbTables.SelectedItem as DataTable);
            try
            {
                if (selectedDt != null)
                {
                    ClearSearchResult();
                    // 이럴일은 없을것같은데 혹시 몰라서..
                    if (_tableList.Contains(selectedDt) == false)
                    {
                        throw new Exception("ListView와 BindingList가 동기화되지 않았습니다. 리스트를 초기화합니다.");
                        ClearBindingList();
                    }

                    BindingDataTable(selectedDt);
                }
            }
            catch(Exception ex)
            {
                ShowError(ex.Message);
            }
        }
        /// <summary>
        /// dt.TableName -> 테이블명 <para />
        /// dt.ExtendedProperties["Remark"] -> 워크시트 하단 REMARK <para />
        /// dt.ExtendedProperties["Description"] -> 테이블 Description
        /// </summary>
        /// <param name="dt"></param>
        private void BindingDataTable(DataTable dt)
        {
            dgvResult.DataSource = dt;
            dgvResult.SwitchColumnSortable(DataGridViewColumnSortMode.NotSortable);
            try
            {
                dgvResult.Columns["ColumnName"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
                dgvResult.Columns["Length"].Width = 60;
                dgvResult.Columns["Prec"].Width = 60;
                dgvResult.Columns["Scale"].Width = 60;
                dgvResult.Columns["Nullable"].Width = 60;
                dgvResult.Columns["Pk"].Width = 60;
                dgvResult.Columns["Description"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
            }
            catch
            { 
                // ignored
            }
            
            string tableDesc = dt.ExtendedProperties["Description"]?.ToString() ?? "테이블 주석 없음";
            lblTableName.Text = dt.TableName;
            lblTableDesc.Text = tableDesc;
            richTxtRemark.Text = dt.ExtendedProperties["Remark"]?.ToString() ?? string.Empty;
        }

        private void btnExportHtml_Click(object sender, EventArgs e)
        {
            if (_tableList.Count == 0)
            {
                MessageBox.Show("내보낼 테이블이 없습니다.", "테이블 없음", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            try
            {
                string filePath = ShowSaveFileDialog(SaveType.HTML);
                if (string.IsNullOrEmpty(filePath))
                    return;
                this.Cursor = Cursors.WaitCursor;

                HtmlManager.ExportToHtml(_tableList.ToList(), filePath, true);

                DialogResult dialogResult = MessageBox.Show("내보내기가 완료되었습니다.\r\n파일을 여시겠습니까?", "완료",
                    MessageBoxButtons.YesNo, MessageBoxIcon.Information);

                if (dialogResult == DialogResult.Yes)
                {
                    Process.Start(filePath);
                }
            }
            catch (Exception ex)
            {
                ShowError(ex.Message);
            }
            finally
            {
                this.Cursor = Cursors.Default;
            }
        }
    }
}