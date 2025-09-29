using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using Newtonsoft.Json.Linq;
using TableDefinitionCreator.dto;

namespace TableDefinitionCreator.utils
{
    internal static class JsonManager
    {
        private const string PROPERTY_NAME = "Tables";
        private const string FILE_NAME = "tables.json";
        private static string ToJsonString(DataTable dt)
        {
            return ToJsonString(new List<DataTable> { dt });
        }
        private static string ToJsonString(List<DataTable> dtList)
        {
            JObject configData = new JObject();

            var result = dtList.Select(dt => new
            {
                TableName = dt.TableName,
                Remark = dt.ExtendedProperties["Remark"]?.ToString() ?? string.Empty
            }).ToArray();

            configData.Add(PROPERTY_NAME, JArray.FromObject(result));
            return configData.ToString();
        }

        /// <summary>
        /// 테이블 정보 리스트를 JSON 파일로 저장합니다.
        /// </summary>
        /// <param name="dtList">Datatable List</param>
        /// <param name="filepath">저장 경로. 기본 저장경로는 "상대경로\tables.json" 입니다.</param>
        /// <exception cref="ArgumentOutOfRangeException">리스트가 NULL 이거나 비어있을 경우</exception>
        public static void SaveToJson(List<DataTable> dtList, string filepath = FILE_NAME)
        {
            if (dtList == null || dtList.Count == 0)
            {
                throw new ArgumentOutOfRangeException(nameof(dtList), "Empty Collection.");
            }

            string jsonString = ToJsonString(dtList);
            File.WriteAllText(filepath, jsonString);
        }

        /// <summary>
        /// JSON 파일을 파싱하여 테이블 정보 리스트를 반환합니다.
        /// </summary>
        /// <param name="filepath">json 파일</param>
        /// <returns></returns>
        /// <exception cref="ArgumentOutOfRangeException">json 형식이 올바르지 않거나 테이블명이 비어있을 경우</exception>
        public static List<TableInformation> FromJson(string filepath)
        {
            var json = File.ReadAllText(filepath);
            var jObject = JObject.Parse(json);
            var tables = jObject[PROPERTY_NAME] as JArray;
            List<TableInformation> tableInfos = new List<TableInformation>();
            if (tables == null)
                throw new ArgumentOutOfRangeException(nameof(tables), "json 형식이 올바르지 않습니다.");

            foreach (var table in tables)
            {
                string tableName = table["TableName"]?.ToString();
                if (string.IsNullOrEmpty(tableName))
                {
                    throw new ArgumentOutOfRangeException(nameof(tableName), "json 형식이 올바르지 않거나 테이블명이 비어있습니다.");
                }

                string remark = table["Remark"]?.ToString() ?? string.Empty;
                tableInfos.Add(new TableInformation
                {
                    TableName = tableName,
                    Remark = remark
                });
            }
            return tableInfos;
        }
    }
}
