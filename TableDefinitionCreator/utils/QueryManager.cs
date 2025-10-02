using System.Collections.Generic;
using System.Data;
using System.Linq;
using TableDefinitionCreator.dto;

namespace TableDefinitionCreator.utils
{
    internal static class QueryManager
    {
        public static List<DataTable> GetTableDefinitionList(string tableName)
        {
            return GetTableDefinitionList(
                new List<TableInformation>
                {
                    new TableInformation
                    {
                        TableName = tableName , Remark = string.Empty
                    }
                });
        }
        public static List<DataTable> GetTableDefinitionList(TableInformation tableInfo)
        {
            return GetTableDefinitionList(
                new List<TableInformation>
                {
                   tableInfo
                });
        }


        /// <remarks>
        /// 문 하나로 처리하기위해 TABLE_NAME과 TABLE_DESC 컬럼을 중복 조회함..
        /// 최종 결과물에는 해당 컬럼 삭제하여 RETURN
        /// </remarks>
        public static List<DataTable> GetTableDefinitionList(List<TableInformation> tableInfos)
        {
            var dtList = new List<DataTable>();
            if (tableInfos == null || tableInfos.Count == 0)
                return dtList;

            string tableCsv = tableInfos.StringJoin(",", ti => ti.TableName);
            string query = $@"
DECLARE @TableName NVARCHAR(MAX) = '{tableCsv}'

SELECT
    C.TABLE_NAME
  , EP_T.value                                                                                                                                                                  AS TABLE_DESC
  , C.COLUMN_NAME                                                                                                                                                               AS ColumnName
  , LOWER( C.DATA_TYPE )                                                                                                                                                        AS Type
  , CASE
        -- 진짜 계산된열이라면
        WHEN CC.definition IS NOT NULL
            THEN 'ComputedColumn'
        -- Identity
        WHEN COLUMNPROPERTY( OBJECT_ID( C.TABLE_SCHEMA + '.' + C.TABLE_NAME ), C.COLUMN_NAME, 'IsIdentity' ) = 1
            THEN 'Identity(' + CAST( IDENT_SEED( C.TABLE_SCHEMA + '.' + C.TABLE_NAME ) AS VARCHAR ) + ',' + CAST( IDENT_INCR( C.TABLE_SCHEMA + '.' + C.TABLE_NAME ) AS VARCHAR ) + ')'
        -- Default
        ELSE CASE
                 WHEN C.COLUMN_DEFAULT LIKE '((%))' THEN SUBSTRING( C.COLUMN_DEFAULT, 3, LEN( C.COLUMN_DEFAULT ) - 4 ) -- ((0))       to 0
                 WHEN C.COLUMN_DEFAULT LIKE '(%)'   THEN SUBSTRING( C.COLUMN_DEFAULT, 2, LEN( C.COLUMN_DEFAULT ) - 2 ) -- (getdate()) to getdate()
                 ELSE C.COLUMN_DEFAULT
             END
    END                                                                                                                                                                         AS Computed
  , CASE
        WHEN CC.definition IS NOT NULL        THEN 'auto'
        WHEN C.CHARACTER_MAXIMUM_LENGTH = -1  THEN 'max'
        ELSE CAST( C.CHARACTER_MAXIMUM_LENGTH AS VARCHAR(10) )
    END AS Length
  , C.NUMERIC_PRECISION                                                                                                                                                         AS Prec
  , C.NUMERIC_SCALE                                                                                                                                                             AS Scale
  , IIF( C.IS_NULLABLE = 'YES', 'Y', NULL )                                                                                                                                     AS Nullable
  , PK.PK_ORDER                                                                                                                                                                 AS Pk
  , CAST( EP.value AS NVARCHAR(500))                                                                                                                                            AS Description
FROM
    INFORMATION_SCHEMA.COLUMNS C
    LEFT JOIN sys.computed_columns CC
              ON C.TABLE_NAME  = OBJECT_NAME(CC.object_id)
             AND C.COLUMN_NAME = CC.name
    LEFT JOIN (
                  SELECT
                      KU.TABLE_CATALOG
                    , KU.TABLE_SCHEMA
                    , KU.TABLE_NAME
                    , KU.COLUMN_NAME
                    , TC.CONSTRAINT_TYPE
                    , KU.ORDINAL_POSITION AS PK_ORDER
                  FROM
                      INFORMATION_SCHEMA.TABLE_CONSTRAINTS           AS TC
                      INNER JOIN INFORMATION_SCHEMA.KEY_COLUMN_USAGE AS KU
                                 ON TC.CONSTRAINT_TYPE = 'PRIMARY KEY'
                                AND TC.CONSTRAINT_NAME = KU.CONSTRAINT_NAME
              ) PK
              ON PK.TABLE_CATALOG = C.TABLE_CATALOG
             AND PK.TABLE_SCHEMA  = C.TABLE_SCHEMA
             AND PK.TABLE_NAME    = C.TABLE_NAME
             AND PK.COLUMN_NAME   = C.COLUMN_NAME
    LEFT JOIN sys.columns SC
              ON SC.object_id     = OBJECT_ID( C.TABLE_SCHEMA + '.' + C.TABLE_NAME )
             AND SC.name          = C.COLUMN_NAME
    LEFT JOIN sys.extended_properties EP
              ON EP.major_id      = SC.object_id
             AND EP.minor_id      = SC.column_id
             AND EP.name          = 'MS_Description'
    LEFT JOIN sys.extended_properties EP_T
              ON EP_T.major_id    = SC.object_id
             AND EP_T.minor_id    = 0
             AND EP_T.name        = 'MS_Description'
WHERE
    C.TABLE_NAME IN ( SELECT TRIM(value) FROM STRING_SPLIT(@TableName, ',') )
ORDER BY
    C.TABLE_NAME
  , C.ORDINAL_POSITION
            ";

            DataTable resultDt = DbAccess.GetDataTable(query);

            // TABLE_NAME별로 그룹화
            var resultDtGroup = resultDt
                .Rows
                .Cast<DataRow>()
                .GroupBy(row => row["TABLE_NAME"].ToString());
            foreach (var group in resultDtGroup)
            {
                DataTable dt = group.CopyToDataTable();
                dt.ExtendedProperties["Remark"] = tableInfos.FirstOrDefault(ti => ti.TableName == group.Key)?.Remark ?? string.Empty;
                dt.ExtendedProperties["Description"] = dt.Rows[0]["TABLE_DESC"]?.ToString() ?? string.Empty;
                dt.TableName = group.Key;

                dt.Columns.Remove("TABLE_NAME");    // 테이블명
                dt.Columns.Remove("TABLE_DESC");    // 테이블설명
                dtList.Add(dt);
            }

            return dtList;
        }
    }
}
