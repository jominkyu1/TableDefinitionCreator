using System.Configuration;
using System.Data.Common;

namespace TableDefinitionCreator.utils
{
    internal static class ConfigManager
    {
        private const string CON_STR_NAME = "WiseM";
        private const string PROVIDER_NAME = "Microsoft.Data.SqlClient";

        private static readonly Configuration configModifier = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
        private static readonly ConnectionStringsSection conSection = (ConnectionStringsSection)configModifier.GetSection("connectionStrings");

        /// <summary>
        /// Nullable
        /// </summary>
        public static string GetConnectionString()
        {
            return conSection.ConnectionStrings[CON_STR_NAME]?.ConnectionString;
        }

        public static void SetConnectionString(string ip, string catalog, string userId, string pwd)
        {
            string connectionString = CreateConnectionStrings(ip, catalog, userId, pwd);
            ConnectionStringSettings conField = conSection.ConnectionStrings[CON_STR_NAME];

            //app.config내에 <connectionStrings><WiseM></></> 없을경우 생성
            if (conField == null)
            {
                conSection.ConnectionStrings.Add(new ConnectionStringSettings(CON_STR_NAME, connectionString, PROVIDER_NAME));
            }
            else
            {
                conField.ConnectionString = connectionString;
                conField.ProviderName = PROVIDER_NAME;
            }

            EncryptConStr();
            configModifier.Save(ConfigurationSaveMode.Modified);
            ConfigurationManager.RefreshSection("connectionStrings");
        }

        private static string CreateConnectionStrings(string ip, string catalog, string userId, string pwd)
        {
            var builder = new DbConnectionStringBuilder();

            builder["Data Source"] = ip;
            builder["Initial Catalog"] = catalog;
            builder["User ID"] = userId;
            builder["Password"] = pwd;
            builder["Connect Timeout"] = 3;

            return builder.ConnectionString;
        }

        private static void EncryptConStr()
        {
            if (!conSection.SectionInformation.IsProtected)
            {
                conSection.SectionInformation.ProtectSection("DataProtectionConfigurationProvider");
            }
        }
    }
}
