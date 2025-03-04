using System.Text.RegularExpressions;
using BarcodeStandard;
using Newtonsoft.Json.Linq;
using SkiaSharp;
using SkiaSharp.QrCode;

namespace WordTemplater
{
    /// <summary>
    /// Provides utility methods for the WordTemplater library.
    /// </summary>
    internal class Utils
    {
        /// <summary>
        /// Parses a string of parameters into a list of objects.
        /// </summary>
        /// <param name="parametters">The string of parameters to parse.</param>
        /// <returns>A list of parsed parameter values.</returns>
        internal static List<object> PaserParametters(string parametters)
        {
            var returnList = new List<object>();
            if (!string.IsNullOrEmpty(parametters))
            {
                var paramList = ParseParameters(
                    parametters.Replace("\\\"", "\"").Replace("\\\\", "\\")
                );

                foreach (var p in paramList)
                {
                    returnList.Add(ConvertStringValue(p));
                }
            }
            return returnList;
        }

        /// <summary>
        /// Parses a string of parameters into a list of strings.
        /// </summary>
        /// <param name="input">The string of parameters to parse.</param>
        /// <returns>A list of parsed parameter strings.</returns>
        public static List<string> ParseParameters(string input)
        {
            var parameters = new List<string>();

            var regex = new Regex(Constant.PARSER_PARAM_REGEX);

            var matches = regex.Matches(input);
            foreach (Match match in matches)
            {
                if (match.Groups[1].Success)
                {
                    parameters.Add("\"" + match.Groups[1].Value + "\"");
                }
                else if (match.Groups[2].Success)
                {
                    parameters.Add("'" + match.Groups[2].Value + "'");
                }
                else if (match.Groups[3].Success)
                {
                    parameters.Add(match.Groups[3].Value);
                }
            }

            return parameters;
        }

        /// <summary>
        /// Converts a string value to the appropriate object type.
        /// </summary>
        /// <param name="value">The string value to convert.</param>
        /// <returns>The converted value as an object.</returns>
        internal static object ConvertStringValue(string value)
        {
            if (string.IsNullOrEmpty(value))
                return string.Empty;

            if (Int64.TryParse(value, out var intVal))
                return intVal;

            if (double.TryParse(value, out var dbVal))
                return dbVal;

            if (decimal.TryParse(value, out var dcVal))
                return dcVal;

            if (DateTime.TryParse(value, out var dtVal))
                return dtVal;

            if (bool.TryParse(value, out var blVal))
                return blVal;

            if ("null".Equals(value, StringComparison.OrdinalIgnoreCase))
                return null;

            if (value.StartsWith('\'') && value.EndsWith('\''))
                value = value.Substring(1, value.Length - 2);
            else if (value.StartsWith('"') && value.EndsWith('"'))
                value = value.Substring(1, value.Length - 2);

            return value;
        }

        /// <summary>
        /// Compares two objects and returns their relationship.
        /// </summary>
        /// <param name="obj1">The first object to compare.</param>
        /// <param name="obj2">The second object to compare.</param>
        /// <returns>A CompareValue indicating the relationship between the objects.</returns>
        public static CompareValue CompareObjects(object obj1, object obj2)
        {
            if (obj1 != null && obj2 != null)
            {
                System.Type type1 = obj1.GetType();
                System.Type type2 = obj2.GetType();

                if (type1 != type2)
                {
                    return CompareValue.Lt | CompareValue.Gt;
                }

                if (obj1 is IComparable comparable1 && obj2 is IComparable comparable2)
                {
                    var i = comparable1.CompareTo(comparable2);
                    if (i < 0)
                        return CompareValue.Lt;
                    else if (i > 0)
                        return CompareValue.Gt;
                    else
                        return CompareValue.Eq;
                }
                else
                {
                    return CompareValue.Lt | CompareValue.Gt;
                }
            }
            else if (obj1 == null)
            {
                if (obj2 == null)
                    return CompareValue.Eq;
                if (
                    string.Compare(obj2.ToString(), "null", StringComparison.OrdinalIgnoreCase) == 0
                )
                    return CompareValue.Eq;
            }

            return CompareValue.Lt | CompareValue.Gt;
        }

        /// <summary>
        /// Generates a QR code image as a stream.
        /// </summary>
        /// <param name="data">The data to encode in the QR code.</param>
        /// <returns>A stream containing the QR code image.</returns>
        internal static Stream GetQRCodeImage(string data)
        {
            var generator = new QRCodeGenerator();
            var qr = generator.CreateQrCode(
                data,
                ECCLevel.L,
                quietZoneSize: Constant.DEFAULT_QRCODE_BORDER
            );
            var info = new SKImageInfo(Constant.DEFAULT_QRCODE_SIZE, Constant.DEFAULT_QRCODE_SIZE);
            using var surface = SKSurface.Create(info);
            var canvas = surface.Canvas;
            canvas.Render(
                qr,
                info.Width,
                info.Height,
                Constant.DEFAULT_DARK_COLOR,
                Constant.DEFAULT_LIGHT_COLOR
            );
            using (var image = surface.Snapshot())
            {
                using (var imgData = image.Encode(SKEncodedImageFormat.Png, 100))
                {
                    Stream stream = new MemoryStream();
                    imgData.SaveTo(stream);
                    stream.Position = 0;
                    return stream;
                }
            }
        }

        /// <summary>
        /// Generates a QR code image as a byte array.
        /// </summary>
        /// <param name="data">The data to encode in the QR code.</param>
        /// <returns>A byte array containing the QR code image.</returns>
        internal static byte[] GetQRCodeImageBytes(string data)
        {
            var generator = new QRCodeGenerator();
            var qr = generator.CreateQrCode(
                data,
                ECCLevel.L,
                quietZoneSize: Constant.DEFAULT_QRCODE_BORDER
            );
            var info = new SKImageInfo(Constant.DEFAULT_QRCODE_SIZE, Constant.DEFAULT_QRCODE_SIZE);
            using var surface = SKSurface.Create(info);
            var canvas = surface.Canvas;
            canvas.Render(
                qr,
                info.Width,
                info.Height,
                Constant.DEFAULT_DARK_COLOR,
                Constant.DEFAULT_LIGHT_COLOR
            );
            using (var image = surface.Snapshot())
            {
                using (var imgData = image.Encode(SKEncodedImageFormat.Png, 100))
                {
                    return imgData.ToArray();
                }
            }
        }

        /// <summary>
        /// Generates a barcode image as a stream.
        /// </summary>
        /// <param name="data">The data to encode in the barcode.</param>
        /// <returns>A stream containing the barcode image.</returns>
        internal static Stream? GetBarCodeImage(string data)
        {
            var barCode = new Barcode();
            barCode.Height = Constant.DEFAULT_BARCODE_HIGHT;
            barCode.BarWidth = Constant.DEFAULT_BARCODE_BARWIDTH;
            var img = barCode.Encode(
                BarcodeStandard.Type.Code128,
                data,
                Constant.DEFAULT_DARK_COLOR,
                Constant.DEFAULT_LIGHT_COLOR
            );
            SKData encoded = img.Encode(SKEncodedImageFormat.Png, 100);
            var stream = encoded.AsStream();
            stream.Position = 0;
            return stream;
        }

        /// <summary>
        /// Generates a barcode image as a byte array.
        /// </summary>
        /// <param name="data">The data to encode in the barcode.</param>
        /// <returns>A byte array containing the barcode image.</returns>
        internal static byte[]? GetBarCodeImageBytes(string data)
        {
            var barCode = new Barcode();
            barCode.Height = Constant.DEFAULT_BARCODE_HIGHT;
            barCode.BarWidth = Constant.DEFAULT_BARCODE_BARWIDTH;
            var img = barCode.Encode(
                BarcodeStandard.Type.Code128,
                data,
                Constant.DEFAULT_DARK_COLOR,
                Constant.DEFAULT_LIGHT_COLOR
            );
            SKData encoded = img.Encode(SKEncodedImageFormat.Png, 100);
            return encoded.ToArray();
        }

        /// <summary>
        /// Gets the size of an image in Word document units.
        /// </summary>
        /// <param name="stream">The stream containing the image data.</param>
        /// <returns>The size of the image in Word document units.</returns>
        internal static Size GetImageSize(Stream stream)
        {
            stream.Position = 0;
            var image = SKImage.FromEncodedData(stream);
            stream.Position = 0;
            if (image != null)
            {
                var width = (long)(image.Width / Constant.DEFAULT_DPI * Constant.PIXEL_PER_INCH);
                var height = (long)(image.Height / Constant.DEFAULT_DPI * Constant.PIXEL_PER_INCH);
                return new Size(width, height);
            }
            return null;
        }

        /// <summary>
        /// Gets all properties that can be repeated from a JSON array.
        /// </summary>
        /// <param name="arr">The JSON array to analyze.</param>
        /// <returns>A list of property names that can be repeated.</returns>
        internal static List<string> GetAllRepeatProperties(JArray arr)
        {
            if (arr == null)
                return new List<string>();
            var repeatProperties = new List<string>();
            foreach (var item in arr)
            {
                if (item is JObject)
                    repeatProperties.AddRange(GetAllRepeatProperties((JObject)item));
            }
            return repeatProperties.Distinct().ToList();
        }

        /// <summary>
        /// Gets all properties that can be repeated from a JSON object.
        /// </summary>
        /// <param name="item">The JSON object to analyze.</param>
        /// <returns>A list of property names that can be repeated.</returns>
        internal static List<string> GetAllRepeatProperties(JObject item)
        {
            if (item == null)
                return new List<string>();
            var repeatProperties = new List<string>();
            repeatProperties.AddRange(item.Properties().Select(p => p.Name).ToList());
            foreach (var subItem in item)
            {
                if (subItem.Value is JArray)
                {
                    repeatProperties.AddRange(GetAllRepeatProperties((JArray)subItem.Value));
                }
                else if (subItem.Value is JObject)
                {
                    repeatProperties.AddRange(GetAllRepeatProperties((JObject)subItem.Value));
                }
            }
            return repeatProperties.Distinct().ToList();
        }

        /// <summary>
        /// Generates a random hexadecimal number with the specified number of digits.
        /// </summary>
        /// <param name="digits">The number of digits in the hexadecimal number.</param>
        /// <returns>A random hexadecimal number as a string.</returns>
        internal static string GetRandomHexNumber(int digits)
        {
            Random random = new Random();
            byte[] buffer = new byte[digits / 2];
            random.NextBytes(buffer);
            string result = string.Concat(buffer.Select(x => x.ToString("X2")).ToArray());
            if (digits % 2 == 0)
                return result;
            return result + random.Next(16).ToString("X");
        }

        private static uint _initUint = 10000U;

        /// <summary>
        /// Generates a unique unsigned integer ID.
        /// </summary>
        /// <returns>A unique unsigned integer ID.</returns>
        internal static uint GetUintId()
        {
            return _initUint++;
        }

        /// <summary>
        /// Generates a unique string ID.
        /// </summary>
        /// <returns>A unique string ID.</returns>
        internal static string GetUniqueStringID()
        {
            return "r" + Guid.NewGuid().ToString().Replace("-", "");
        }

        /// <summary>
        /// Clones a stream to a new memory stream.
        /// </summary>
        /// <param name="stream">The stream to clone.</param>
        /// <returns>A new memory stream containing the same data.</returns>
        internal static Stream CloneStream(Stream stream)
        {
            if (stream != null && stream.CanSeek)
            {
                stream.Position = 0;
                MemoryStream newStream = new MemoryStream();
                stream.CopyTo(newStream);
                stream.Position = 0;
                newStream.Position = 0;
                return newStream;
            }
            return null;
        }
    }

    /// <summary>
    /// Represents a size in Word document units.
    /// </summary>
    internal class Size
    {
        /// <summary>
        /// Gets or sets the width in Word document units.
        /// </summary>
        internal long Width { get; set; }

        /// <summary>
        /// Gets or sets the height in Word document units.
        /// </summary>
        internal long Height { get; set; }

        /// <summary>
        /// Initializes a new instance of the Size class.
        /// </summary>
        /// <param name="w">The width in Word document units.</param>
        /// <param name="h">The height in Word document units.</param>
        internal Size(long w, long h)
        {
            Width = w;
            Height = h;
        }
    }
}
