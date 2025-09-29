using System;
using System.Collections.Generic;
using System.Linq;
using ClosedXML.Excel;
using DataTable = System.Data.DataTable;

namespace TableDefinitionCreator.utils
{
    internal static class ExcelManager
    { 
        private const string COVER_SHEET_NAME = "목차";
        private static readonly XLColor _primaryBackgroundColor = XLColor.FromArgb(45, 65, 85); // 어두운 슬레이트 색
        private static readonly XLColor _colBackgroundColor = XLColor.FromArgb(204, 255, 204); // 연한 연두색

        /// <summary>
        /// 테이블 정의서를 엑셀로 내보냅니다.
        /// </summary>
        /// <param name="tables">테이블 목록</param>
        /// <param name="fullFilePath">내보낼 경로 ex C:\Test.xlsx</param>
        /// <param name="createCoverPage">목차 페이지를 생성할지 여부</param>
        public static void ExportToExcel(List<DataTable> tables, string fullFilePath, bool createCoverPage)
        {
            using (var workbook = new XLWorkbook())
            {
                // 전역 설정
                SetGlobalSettings(workbook);

                // 각 테이블 정렬 후 WorkSheet 생성
                foreach (DataTable dt in tables.OrderBy(table => table.TableName))
                {
                    CreateDescSheet(workbook, dt);
                }

                // 커버페이지
                if(createCoverPage)
                    CreateCoverSheet(workbook, DbAccess.InitialCatalog);

                // 파일 생성
                workbook.SaveAs(fullFilePath);
            }
        }

        /// <summary>
        /// DataTable을 이용하여 Table Description WorkSheet 생성
        /// </summary>
        private static void CreateDescSheet(XLWorkbook workbook, DataTable dt)
        {
            var worksheet = workbook.Worksheets.Add(dt.TableName);

            // 첫번째 행
            worksheet.Cell(1, 1).Value = "TableName";
            worksheet.Cell(1, 2).Value = dt.TableName;
            worksheet.Cell(1, 5).Value = "Desc.";
            worksheet.Cell(1, 6).Value = dt.ExtendedProperties["Description"]?.ToString();
            worksheet.Cell(1, 10).Value = $"{COVER_SHEET_NAME} 돌아가기";
            worksheet.Cell(1, 10).SetHyperlink(new XLHyperlink($"'{COVER_SHEET_NAME}'!A1"));
            // 첫번째행의 2~4컬럼 병합 ( TableName )
            worksheet.Range(1, 2, 1, 4).Merge();
            // 첫번째행의 6~9컬럼 병합 ( TableDesc )
            worksheet.Range(1, 6, 1, 9).Merge();

            // 두번째 행
            for (int i = 0; i < dt.Columns.Count; i++)
            {
                worksheet.Cell(2, i + 1).Value = dt.Columns[i].ColumnName;
            }

            // 컬럼 너비 조정
            worksheet.Columns(1, 1).Width = 20; // ColumnName
            worksheet.Columns(2, 2).Width = 10; // Type
            worksheet.Columns(3, 4).Width = 15; // Computed
            worksheet.Columns(4, 4).Width = 10; // Length
            worksheet.Columns(5, 5).Width = 7; // Prec
            worksheet.Columns(6, 6).Width = 7; // Scale
            worksheet.Columns(7, 7).Width = 7; // Nullable
            worksheet.Columns(8, 8).Width = 7; // Pk
            worksheet.Columns(9, 9).Width = 35; // Description
            worksheet.Columns(10, 10).Width = 20; // 목차로 돌아가기

            // Body 
            // createTable true설정시 Excel의 Table표형태로 삽입됨
            IXLTable table = worksheet.Cell(2, 1).InsertTable(dt, false);
            // TableName, Description은 좌측정렬
            table.Range(2, 1, table.RowCount(), 1).Style.Alignment.Horizontal =
                XLAlignmentHorizontalValues.Left;
            table.Range(2, 9, table.RowCount(), 9).Style.Alignment.Horizontal =
                XLAlignmentHorizontalValues.Left;

            // Note
            if (string.IsNullOrEmpty(dt.ExtendedProperties["Remark"]?.ToString()) == false)
            {
                int noteRow = dt.Rows.Count + 2 + 2;
                var noteRowRange = worksheet.Range(noteRow, 2, noteRow, 9);

                noteRowRange.Merge();

                noteRowRange.Style.Alignment.WrapText = true;
                noteRowRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
                noteRowRange.Style.Fill.BackgroundColor = XLColor.LightYellow;

                worksheet.Cell(noteRow, 1).Value = "REMARK";
                worksheet.Cell(noteRow, 1).Style.Font.Bold = true;
                worksheet.Cell(noteRow, 1).Style.Font.FontSize = 14;
                worksheet.Cell(noteRow, 2).Value = dt.ExtendedProperties["Remark"]?.ToString();


                // 병합된 CELL은 RowHeight 자동조절이 안돼서 임시컬럼 생성하여 맞춘 후 Hide
                int tempColumn = 11;
                worksheet.Column(tempColumn).Width = 150;
                var helperCell = worksheet.Cell(noteRow, tempColumn);
                helperCell.Value = dt.ExtendedProperties["Remark"]?.ToString();
                helperCell.Style = noteRowRange.Style;

                worksheet.Column(tempColumn).AdjustToContents(noteRow, noteRow);
                worksheet.Column(tempColumn).Hide();
            }


            // Background Color
            worksheet.Cell(1, 1).Style.Fill.SetBackgroundColor(_colBackgroundColor); // TableName
            worksheet.Cell(1, 5).Style.Fill.SetBackgroundColor(_colBackgroundColor); // Desc.
            worksheet.Row(2).Cells(1, 9).Style.Fill.SetBackgroundColor(_colBackgroundColor); // Columns

            // Border
            var range = worksheet.Range(1, 1, dt.Rows.Count + 2, 9);
            range?.Style.Border.SetOutsideBorder(XLBorderStyleValues.Thin);
            range?.Style.Border.SetInsideBorder(XLBorderStyleValues.Thin);

            // Height
            worksheet.Rows(1, dt.Rows.Count + 2).Height = 20;
        }

        /// <summary>
        /// EXCEL 전역설정
        /// </summary>
        private static void SetGlobalSettings(IXLWorkbook workbook)
        {
            workbook.Style.Font.FontName = "맑은 고딕";
            workbook.Style.Font.FontSize = 9;
            workbook.ShowGridLines = false; // 배경 바둑판제거
            workbook.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center; // 가로정렬
            workbook.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center; // 세로정렬
        }

        /// <summary>
        /// 커버페이지 생성 ( WorkSheet가 먼저 생성되어 있어야 함. 없으면 생성안함 )
        /// </summary>
        private static void CreateCoverSheet(IXLWorkbook workbook, string dbName)
        {
            
            var tableSheets = workbook.Worksheets.Where(ws => ws.Name != COVER_SHEET_NAME).ToList();
            if (!tableSheets.Any())
                return;

            // 시트가 이미 있다면 삭제하고 새로 만듦
            if (workbook.Worksheets.Contains(COVER_SHEET_NAME))
            {
                workbook.Worksheet(COVER_SHEET_NAME).Delete();
            }

            // 시트를 가장 첫 번째 위치에 추가
            var coverSheet = workbook.Worksheets.Add(COVER_SHEET_NAME, 0);
            coverSheet.ShowGridLines = false;
            coverSheet.TabColor = _primaryBackgroundColor;
            // --- 2. 상단 헤더 ---
            coverSheet.Range("B2:J3").Merge().Value = $"[ {dbName} ] Wise-MES 테이블 정의서";
            coverSheet.Rows(2, 3).Height = 30;

            var header = coverSheet.Cell("B2");
            header.Style.Font.Bold = true;
            header.Style.Font.FontSize = 22;
            header.Style.Font.FontColor = XLColor.White;
            header.Style.Fill.BackgroundColor = _primaryBackgroundColor;

            // --- 3. 보고서 정보 추가 ---
            coverSheet.Cell("B5").Value = "기준 일자";
            coverSheet.Cell("C5").Value = DateTime.Now.ToString("yyyy-MM-dd");
            coverSheet.Range("B5:C5").Style.Font.FontSize = 12;
            coverSheet.Range("B5:C5").Style.Font.FontColor = XLColor.Gray;
            coverSheet.Range("B5:C5").Style.Font.Bold = true;


            // --- 4. 목차 및 하이퍼링크 생성 ---
            var groupedTables = tableSheets
                .Select(sheet => sheet.Name)
                .OrderBy(name => name)
                .GroupBy(name => char.ToUpper(name[0]))
                .OrderBy(group => group.Key)
                .ToList();

            int linkRow = 7; // 목차 시작 행

            foreach (var group in groupedTables)
            {
                // 4-1. 그룹 헤더 [A], [B], [C]...
                var headerCell = coverSheet.Cell(linkRow, 2); // B열에 그룹 헤더
                headerCell.Value = $"   [ {group.Key} ]";
                headerCell.Style.Font.Bold = true;
                headerCell.Style.Font.FontSize = 14;
                headerCell.Style.Font.FontColor = XLColor.FromArgb(45, 65, 85);
                headerCell.Style.Fill.BackgroundColor = XLColor.WhiteSmoke;
                headerCell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;

                // 그룹 헤더 아래에 얇은 선 추가
                coverSheet.Range(linkRow, 2, linkRow, 10)
                    .Merge()
                    .Style.Border.SetBottomBorder(XLBorderStyleValues.Thin);
                linkRow += 2; // 그룹 헤더 후 두 줄 띄우기

                // 4-2. 해당 그룹에 속한 테이블 링크들
                int currentColumn = 2; // B열부터 링크 시작
                foreach (var tableName in group)
                {
                    var linkCell = coverSheet.Cell(linkRow, currentColumn);
                    linkCell.Value = tableName;
                    linkCell.SetHyperlink(new XLHyperlink($"'{tableName}'!A1"));

                    // 링크 스타일
                    linkCell.Style.Font.FontColor = XLColor.Blue;
                    linkCell.Style.Font.Underline = XLFontUnderlineValues.None; // 밑줄 제거

                    currentColumn++;
                    // 한 줄에 9개씩
                    if (currentColumn > 10)
                    {
                        currentColumn = 2;
                        linkRow++;
                    }
                }
                linkRow += 2; // 그룹 간 간격
            }

            // 열 너비 조정
            coverSheet.Columns(2, 10).Width = 25;
            coverSheet.Column(1).Width = 3; // 왼쪽 여백
            // 외곽선
            coverSheet.RangeUsed(XLCellsUsedOptions.All)?.Style.Border.SetOutsideBorder(XLBorderStyleValues.Thin);
        }
    }
}
