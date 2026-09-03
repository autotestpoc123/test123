using Azure;
using ICSharpCode.SharpZipLib.Tar;
using ICSharpCode.SharpZipLib.Zip;
using MorganStanley.COD.FirmwideDirectory.API.Faults;
using MorganStanley.COD.FirmwideDirectory.API.Models.Options;
using System.Collections.Generic;
using System.Globalization;
using System.Net;
using System.Numerics;
using System.Reflection;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using static Microsoft.Extensions.Logging.EventSource.LoggingEventSource;

namespace MorganStanley.COD.FirmwideDirectory.API.Common
{
    public static class Utility
    {
        // ShortPhoneNumberMapping will be filled when load the user configuration
        public static Dictionary<string, string> ShortPhoneNumberMapping = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        public static WebProxy CreateWebProxy(HttpOptions config)
            => CreateWebProxy(config.ProxyUri, config.UseDefaultProxyCredentials, config.BypassProxyOnLocal, config.BypassList);

        public static WebProxy CreateWebProxy(string uri, bool useDefaultProxyCredentials = true, bool bypassProxyOnLocal = true, string[] bypassList = null)
            => new WebProxy
            {
                Address = new Uri(uri),
                UseDefaultCredentials = useDefaultProxyCredentials,
                BypassProxyOnLocal = bypassProxyOnLocal,
                BypassList = bypassList
            };

        /// <summary>
        /// Is the string in the format of mail address
        /// </summary>
        /// <param name="mail"></param>
        public static bool IsMailAddress(string mail)
        {
            if (string.IsNullOrWhiteSpace(mail))
            {
                return false;
            }

            return mail.Contains('@');
        }

        /// <summary>
        /// Is the string in the format of phone number
        /// </summary>
        /// <remarks>
        /// If length is 5 and pure numbers, then we asuing it is MSID, not a PhoneNumber
        /// </remarks>
        /// <param name="phoneNumber"></param>
        public static bool IsPhoneNumber(string phoneNumber)
        {
            if (string.IsNullOrWhiteSpace(phoneNumber) || phoneNumber.Length < 3)
            {
                return false;
            }

            // we assuming it is MSID with only numbers
            if (phoneNumber.Length == 5 && Regex.IsMatch(phoneNumber, @"^[0-9]+$"))
            {
                return false;
            }

            // basic phone number characters check
            return Regex.IsMatch(phoneNumber, @"^[0-9()+\-\s]+$");
        }

        public static bool IsReadyToLoad(DateTime _LastLoadTime, GlobalFileLoadOption _options, ILogger _logger)
        {

            bool isTimetoLoad = false;
            // check whether it is the time to load User Data
            if (string.IsNullOrWhiteSpace(_options.UpdateTimeWindow) is false)
            {
                string[] loadTimes = _options.UpdateTimeWindow.Split('-');
                if (loadTimes.Length == 2)
                {
                    TimeSpan.TryParse(loadTimes[0], CultureInfo.InvariantCulture, out TimeSpan start);
                    TimeSpan.TryParse(loadTimes[1], CultureInfo.InvariantCulture, out TimeSpan end);
                    TimeSpan now = DateTime.Now.TimeOfDay;
                    // Make sure we only load once per day within the Window
                    if (now >= start && now <= end && (DateTime.Now.Date != _LastLoadTime.Date))
                    {
                        isTimetoLoad = true;
                    }
                }
            }
            if (isTimetoLoad == false && _options.SkipValidation == false && _LastLoadTime > (DateTime.UnixEpoch))
            {
                return false;
            }
            if (!File.Exists(_options.ZipFilePath))
            {
                throw Faults.Faults.GlobalXmlFileLoadFailed($"{_options.ZipFilePath} could not be accessed");
            }

            var lastWriteTime = File.GetLastWriteTime(_options.ZipFilePath);
            if (lastWriteTime < _LastLoadTime && _options.SkipValidation == false)
            {
                _logger.LogWarning(string.Format("Skip loading data as {0} already loaded, LastWriteTime {1}, LastLoadTime {2}", _options.ZipFilePath, lastWriteTime, _LastLoadTime));
                return false;
            }

            return true;
        }

        /// <summary>
        /// Init the ShortPhoneNumberMapping: mapping longPrefix to shortPrefix
        /// </summary>
        /// <param name="phoneNumberMappingConfigPath"></param>
        public static void InitShortPhoneNumberMapping(string phoneNumberMappingConfigPath)
        {
            var mapping = ParseShortPhoneNumberMapping(phoneNumberMappingConfigPath);
            if (mapping.Count > 0)
            {
                ShortPhoneNumberMapping = mapping;
            }
        }

        /// <summary>
        /// Parse the phoneNumberMappingConfigPath
        /// </summary>
        /// <param name="phoneNumberMadineConfigPath"></param>
        private static Dictionary<string, string> ParseShortPhoneNumberMapping(string phoneNumberMappingConfigPath)
        {
            Dictionary<string, string> shortPhoneNumberMapping = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            if (string.IsNullOrEmpty(phoneNumberMappingConfigPath))
            {
                return shortPhoneNumberMapping;
            }

            var mappingContent = File.ReadLines(phoneNumberMappingConfigPath);
            foreach (var line in mappingContent)
            {
                if (string.IsNullOrWhiteSpace(line))
                { continue; } // Skip empty lines

                // Split by comma (basic CSV, no quotes handling)
                string[] values = line.Split(',');
                if (values.Length >= 2)
                {
                    string shortNumber = values[0].Trim().Replace("X", string.Empty);
                    string longNumber = values[1].Trim().Replace("X", string.Empty);
                    if (string.IsNullOrWhiteSpace(shortNumber) == false && string.IsNullOrWhiteSpace(longNumber) == false)
                    {
                        shortPhoneNumberMapping.TryAdd(longNumber, shortNumber);
                    }
                }
            }

            return shortPhoneNumberMapping;
        }

        public static string GenerateCODPhoneShort(string phone)
        {
            if (string.IsNullOrWhiteSpace(phone))
            {
                return string.Empty;
            }

            // if the ShortPhoneNumberMapping rules not empty, use the mapping rule
            if (ShortPhoneNumberMapping.Count != 0)
            {
                return GenerateShortPhoneNumber(phone);
            }
            else
            {
                // Otherwise, use the hardcode rules below
                string trimedPhone = phone.Trim();
                if (trimedPhone.StartsWith("+86 756"))
                {
                    // For Zhuhai, the phone is 7 digits. Remove the prefix
                    return trimedPhone.Substring("+86 756".Length).Trim();
                }
                else if (trimedPhone.StartsWith("+86 21"))
                {
                    // For Shanghai, the phone is 8 digits. Remove the prefix and remove the second digit of phone number
                    return trimedPhone.Substring("+86 21".Length).Trim().Remove(1, 1);
                }
                else if (phone.StartsWith("+86 10 83"))
                {
                    // For Beijing, if MSMS and MSBIC, then remove the prefix and remove the first digit of phone number
                    return trimedPhone.Substring("+86 10".Length).Trim().Remove(0, 1);
                }
                else if (trimedPhone.StartsWith("+86 10 89"))
                {
                    // For Beijing, if MSFC, then remove the prefix and remove the second digit of phone number
                    return trimedPhone.Substring("+86 10".Length).Trim().Remove(1, 1);
                }
                else
                {
                    return string.Empty;
                }
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="phoneNumber"></param>
        /// <returns>shortNumber if any rule matched; otherwise string.Empty</returns>
        public static string GenerateShortPhoneNumber(string phoneNumber)
        {
            if (string.IsNullOrWhiteSpace(phoneNumber))
            {
                return string.Empty;
            }

            if (ShortPhoneNumberMapping.Count == 0)
            {
                return string.Empty;
            }

            string workPhoneFormatted = phoneNumber.Replace(" ", string.Empty).Replace("-", string.Empty).Trim();
            var matchedLongNumerPrefixes = ShortPhoneNumberMapping.Keys.Where(x => workPhoneFormatted.StartsWith(x)).ToList();

            // only mapping if only 1 prefix matched
            if (matchedLongNumerPrefixes.Count == 1)
            {
                string shortPhone = workPhoneFormatted.Replace(matchedLongNumerPrefixes[0], ShortPhoneNumberMapping[matchedLongNumerPrefixes[0]]);
                // shortNumber length always 7 as confirmed with ENS
                if (shortPhone.Length == 7)
                {
                    return shortPhone.Substring(0, 3) + "-" + shortPhone.Substring(3, 4);
                }
                else
                {
                    return string.Empty;
                }
            }
            else
            {
                return string.Empty;
            }
        }
        public static string FormatCODPhone(string phone)
        {
            if (string.IsNullOrWhiteSpace(phone))
            {
                return string.Empty;
            }

            if (phone.Length == 7 && !phone.Contains("-"))
            {
                // assuming shortnumber and does not contains '-'
                // pick the last 6 number and add dash (xx-xxxx)
                // using the last 6 number as the shortNumber rule is not consistent
                return phone.Substring(1, 2) + "-" + phone.Substring(phone.Length - 4, 4);
            }
            else if (phone.Length == 8 && phone.Contains("-"))
            {
                // assuming shortnumber and contains '-'
                // pick the last 7 (xx-xxxx)
                // using the last 6 number as the shortNumber rule is not consistent
                return phone.Substring(phone.Length - 7, 7);
            }

            return phone;
        }

        public static bool ZipFile(string filePath, string zipFilePath)
        {
            if (File.Exists(filePath))
            {
                using (ZipOutputStream s = new ZipOutputStream(File.Create(zipFilePath)))
                {
                    s.SetLevel(5);
                    byte[] buffer = new byte[4096000];

                    var entry = new ZipEntry(Path.GetFileName(filePath));
                    s.PutNextEntry(entry);

                    using (FileStream fs = File.OpenRead(filePath))
                    {
                        int sourceBytes;
                        do
                        {
                            sourceBytes = fs.Read(buffer, 0, buffer.Length);
                            s.Write(buffer, 0, sourceBytes);
                        } while (sourceBytes > 0);
                    }
                    s.Finish();
                    s.Close();
                }
                return true;
            }
            else
            {
                return false;
            }
        }
        /// <summary>
        /// Generate ETag based on the rawData
        /// </summary>
        /// <param name="rawData"></param>
        /// <returns></returns>

        public static string GenerateETag(string rawData)
        {
            using var sha256 = SHA256.Create();
            {
                byte[] hashBytes = sha256.ComputeHash(Encoding.UTF8.GetBytes(rawData));
                return Convert.ToBase64String(hashBytes);
            }
        }

        public static bool IsValidMSIDForPhoto(string msid)
        {
            if (string.IsNullOrWhiteSpace(msid))
            {
                return false;
            }

            return Regex.IsMatch(msid, @"^[A-Za-z0-9]+$");
        }

        /// <summary>
        /// Get user's photo fullPath based on MSID
        /// </summary>
        /// <Remarks>
        /// There will be large numbers of photos. To avoid too many items in one folder, will use MSID first and second char as subfolders.
        /// The folder will be created in advance.
        /// </Remarks>
        /// <param name="msid"></param>
        /// <param name="photoOptions"></param>
        /// <returns></returns>
        public static string GetUserPhotoFullPath(string msid, PhotoOptions photoOptions)
        {
            if (string.IsNullOrEmpty(photoOptions.PhotoFolder) || Directory.Exists(photoOptions.PhotoFolder) == false)
            {
                throw Faults.Faults.CreatePhotoCouldNotAccess($"PhotoFolder {photoOptions.PhotoFolder} could not be accessed");
            }

            if (IsValidMSIDForPhoto(msid) == false)
            {
                throw Faults.Faults.CreateInvalidArgument(nameof(msid));
            }

            var photoSubPath = Char.ToUpper(msid[0]) + @"\" + Char.ToUpper(msid[1]) + @"\" + $"{msid}{photoOptions.PhotoType}";
            var photoRootPath = Path.GetFullPath(photoOptions.PhotoFolder);
            var photoFullPath = Path.GetFullPath(Path.Combine(photoRootPath, photoSubPath));

            if (!photoRootPath.EndsWith(Path.DirectorySeparatorChar.ToString()))
            {
                photoRootPath += Path.DirectorySeparatorChar;
            }

            if (photoFullPath.StartsWith(photoRootPath, StringComparison.OrdinalIgnoreCase) == false)
            {
                throw Faults.Faults.CreateInvalidArgument(nameof(msid));
            }

            return photoFullPath;
        }
    }
}
