using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using DataTable = System.Data.DataTable;

namespace TableDefinitionCreator.utils
{
    internal static class HtmlManager
    {
        private const string COVER_PAGE_ID = "cover";
        private const string PRIMARY_COLOR = "#2d4155";
        private const string COL_BACKGROUND_COLOR = "#ccffcc";
        private const string LIGHT_YELLOW = "#ffffe0";

        /// <summary>
        /// 테이블 정의서를 HTML로 내보냅니다.
        /// </summary>
        /// <param name="tables">테이블 목록</param>
        /// <param name="fullFilePath">내보낼 경로 ex C:\Test.html</param>
        /// <param name="createCoverPage">목차 페이지를 생성할지 여부</param>
        public static void ExportToHtml(List<DataTable> tables, string fullFilePath, bool createCoverPage)
        {
            var html = new StringBuilder();

            // HTML 시작
            html.AppendLine("<!DOCTYPE html>");
            html.AppendLine("<html lang='ko'>");
            html.AppendLine("<head>");
            html.AppendLine("    <meta charset='UTF-8'>");
            html.AppendLine("    <meta name='viewport' content='width=device-width, initial-scale=1.0'>");
            html.AppendLine($"    <title>{DbAccess.InitialCatalog} - 테이블 정의서</title>");
            html.AppendLine(GetStyles());
            html.AppendLine("</head>");
            html.AppendLine("<body>");

            // 커버페이지
            if (createCoverPage)
            {
                html.AppendLine(CreateCoverPage(tables, DbAccess.InitialCatalog));
            }

            // 각 테이블 정렬 후 생성
            foreach (DataTable dt in tables.OrderBy(table => table.TableName))
            {
                html.AppendLine(CreateTableSection(dt));
            }

            // HTML 종료
            html.AppendLine("</body>");
            html.AppendLine("</html>");

            // 파일 생성
            File.WriteAllText(fullFilePath, html.ToString(), Encoding.UTF8);
        }

        /// <summary>
        /// CSS 스타일 생성
        /// </summary>
        private static string GetStyles()
        {
            return @"
    <style>
        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }
        
        body {
            font-family: '맑은 고딕', 'Malgun Gothic', sans-serif;
            font-size: 12pt;
            background-color: #f5f5f5;
            padding: 20px;
        }
        
        .cover-page {
            background: white;
            padding: 40px;
            margin-bottom: 30px;
            border-radius: 8px;
            box-shadow: 0 2px 8px rgba(0,0,0,0.1);
        }
        
        .cover-header {
            background-color: " + PRIMARY_COLOR + @";
            color: white;
            padding: 20px;
            font-size: 22pt;
            font-weight: bold;
            text-align: center;
            border-radius: 4px;
            margin-bottom: 20px;
        }
        
        .cover-info {
            font-size: 12pt;
            color: #666;
            font-weight: bold;
            margin-bottom: 30px;
        }
        
        .cover-group {
            margin-bottom: 30px;
        }
        
        .cover-group-header {
            font-size: 14pt;
            font-weight: bold;
            color: " + PRIMARY_COLOR + @";
            background-color: #f5f5f5;
            padding: 10px;
            border-bottom: 1px solid #ddd;
            margin-bottom: 15px;
        }
        
        .cover-links {
            display: grid;
            grid-template-columns: repeat(auto-fill, minmax(200px, 1fr));
            gap: 10px;
            padding: 0 10px;
        }
        
        .cover-links a {
            color: #0066cc;
            text-decoration: none;
            padding: 8px 12px;
            border-radius: 4px;
            transition: background-color 0.2s;
        }
        
        .cover-links a:hover {
            background-color: #e8f4ff;
        }
        
        .table-section {
            background: white;
            padding: 30px;
            margin-bottom: 30px;
            border-radius: 8px;
            box-shadow: 0 2px 8px rgba(0,0,0,0.1);
        }
        
        .table-header {
            display: grid;
            grid-template-columns: 200px 1fr 150px 1fr auto;
            gap: 10px;
            margin-bottom: 20px;
            padding: 15px;
            border: 1px solid #ddd;
        }
        
        .table-header-label {
            background-color: " + COL_BACKGROUND_COLOR + @";
            padding: 10px;
            font-weight: bold;
            text-align: center;
        }
        
        .table-header-value {
            padding: 10px;
            text-align: left;
        }
        
        .back-link {
            grid-column: 5;
            text-align: center;
        }
        
        .back-link a {
            color: #0066cc;
            text-decoration: none;
            padding: 8px 16px;
            border-radius: 4px;
            display: inline-block;
        }
        
        .back-link a:hover {
            background-color: #e8f4ff;
        }
        
        table {
            width: 100%;
            border-collapse: collapse;
            margin-bottom: 20px;
            table-layout: fixed;
        }
        
        th, td {
            padding: 10px;
            border: 1px solid #ddd;
            word-wrap: break-word;
            word-break: break-word;
            overflow-wrap: break-word;
        }
        
        th {
            background-color: " + COL_BACKGROUND_COLOR + @";
            text-align: center;
            font-weight: bold;
        }
        
        td {
            text-align: center;
        }
        
        /* 컬럼별 고정 너비 */
        th:nth-child(1), td:nth-child(1) { width: 20%; text-align: left; }   /* ColumnName */
        th:nth-child(2), td:nth-child(2) { width: 10%; }                      /* Type */
        th:nth-child(3), td:nth-child(3) { width: 10%; }                      /* Computed */
        th:nth-child(4), td:nth-child(4) { width: 8%; }                       /* Length */
        th:nth-child(5), td:nth-child(5) { width: 7%; }                       /* Prec */
        th:nth-child(6), td:nth-child(6) { width: 7%; }                       /* Scale */
        th:nth-child(7), td:nth-child(7) { width: 7%; }                       /* Nullable */
        th:nth-child(8), td:nth-child(8) { width: 7%; }                       /* Pk */
        th:nth-child(9), td:nth-child(9) { width: 24%; text-align: left; }   /* Description */
        
        .remark-section {
            margin-top: 20px;
            padding: 15px;
            border: 1px solid #ddd;
        }
        
        .remark-label {
            font-weight: bold;
            font-size: 14pt;
            margin-bottom: 10px;
        }
        
        .remark-content {
            background-color: " + LIGHT_YELLOW + @";
            padding: 15px;
            white-space: pre-wrap;
            text-align: left;
        }
    </style>";
        }

        /// <summary>
        /// 커버페이지 생성
        /// </summary>
        private static string CreateCoverPage(List<DataTable> tables, string dbName)
        {
            var sb = new StringBuilder();

            sb.AppendLine($"<div id='{COVER_PAGE_ID}' class='cover-page'>");
            sb.AppendLine($"    <div class='cover-header'>[ {dbName} ] Wise-MES 테이블 정의서</div>");
            sb.AppendLine($"    <div class='cover-info'>기준 일자: {DateTime.Now:yyyy-MM-dd}</div>");

            // 테이블을 첫 글자로 그룹화
            var groupedTables = tables
                .OrderBy(t => t.TableName)
                .GroupBy(t => char.ToUpper(t.TableName[0]))
                .OrderBy(g => g.Key);

            foreach (var group in groupedTables)
            {
                sb.AppendLine("    <div class='cover-group'>");
                sb.AppendLine($"        <div class='cover-group-header'>[ {group.Key} ]</div>");
                sb.AppendLine("        <div class='cover-links'>");

                foreach (var table in group)
                {
                    sb.AppendLine($"            <a href='#{table.TableName}'>{table.TableName}</a>");
                }

                sb.AppendLine("        </div>");
                sb.AppendLine("    </div>");
            }

            sb.AppendLine("</div>");
            return sb.ToString();
        }

        /// <summary>
        /// 테이블 섹션 생성
        /// </summary>
        private static string CreateTableSection(DataTable dt)
        {
            var sb = new StringBuilder();

            sb.AppendLine($"<div id='{dt.TableName}' class='table-section'>");

            // 헤더
            sb.AppendLine("    <div class='table-header'>");
            sb.AppendLine("        <div class='table-header-label'>TableName</div>");
            sb.AppendLine($"        <div class='table-header-value'>{dt.TableName}</div>");
            sb.AppendLine("        <div class='table-header-label'>Desc.</div>");
            sb.AppendLine($"        <div class='table-header-value'>{dt.ExtendedProperties["Description"]?.ToString() ?? ""}</div>");
            sb.AppendLine("        <div class='back-link'><a href='#cover'>목차 돌아가기</a></div>");
            sb.AppendLine("    </div>");

            // 테이블
            sb.AppendLine("    <table>");
            sb.AppendLine("        <thead>");
            sb.AppendLine("            <tr>");
            foreach (System.Data.DataColumn col in dt.Columns)
            {
                sb.AppendLine($"                <th>{col.ColumnName}</th>");
            }
            sb.AppendLine("            </tr>");
            sb.AppendLine("        </thead>");
            sb.AppendLine("        <tbody>");

            foreach (System.Data.DataRow row in dt.Rows)
            {
                sb.AppendLine("            <tr>");
                foreach (var item in row.ItemArray)
                {
                    sb.AppendLine($"                <td>{item?.ToString() ?? ""}</td>");
                }
                sb.AppendLine("            </tr>");
            }

            sb.AppendLine("        </tbody>");
            sb.AppendLine("    </table>");

            // Remark
            if (!string.IsNullOrEmpty(dt.ExtendedProperties["Remark"]?.ToString()))
            {
                sb.AppendLine("    <div class='remark-section'>");
                sb.AppendLine("        <div class='remark-label'>REMARK</div>");
                sb.AppendLine($"        <div class='remark-content'>{dt.ExtendedProperties["Remark"]}</div>");
                sb.AppendLine("    </div>");
            }

            sb.AppendLine("</div>");
            return sb.ToString();
        }
    }
}