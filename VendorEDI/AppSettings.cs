using System;
using System.ComponentModel;
using System.Configuration;

namespace VendorEDI
{
    public static class AppSettings
    {
        public static T Get<T>(string key)
        {
            var appSetting = ConfigurationManager.AppSettings[key];
            if (string.IsNullOrWhiteSpace(appSetting))
                throw new ArgumentOutOfRangeException("key", key, "Unable to retrieve AppSetting value.");

            var converter = TypeDescriptor.GetConverter(typeof(T));
            return (T)(converter.ConvertFromInvariantString(appSetting));
        }
    }
}