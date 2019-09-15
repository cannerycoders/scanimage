using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WIA;
using System.IO;


namespace scanimage
{
    class Program
    {
        static int Main(string[] args)
        {
            string usage = "scanimage usage\n\n" +
               "\t-listscanners  (if provided we list, then bail)\n" +
               "\t-scanner id (default is first in the list)\n" +
               "\t-x x (0) offset \n" +
               "\t-y y (0) offst \n" +
               "\t-width x (1000)\n" +
               "\t-height y (1518)\n" +
               "\t-resolution r (150)\n" +
               "\t-outfile fn (default is scan.jpg)\n" +
               "\t-help  (this message)";

            // Configure scanner
            //  comic is 10.125 x 6.625
            // at 150dpi, 1003x1518
            int resolution = 150;
            int startLeft = 0;
            int startTop = 0;
            int widthPixels = 1000;
            int heightPixels = 1518;
            int color_mode = 0; /* 0: rgb, 1: grayscale 2: monochrome, 3: autocolor */
            int brightnessPct = 0;
            int contrastPct = 0;
            int deviceId = -1;
            bool err;
            string path = "scan.jpg";

            for (int i=0;i<args.Length;i++)
            {
                switch(args[i])
                {
                    case "-help":
                        bail(usage);
                        break;
                    case "-l": // fall thru
                    case "-listscanners":
                        listScanningDevices();
                        return 0;
                    case "-scanner":
                        err = int.TryParse(args[i + 1], out deviceId);
                        if (!err)
                            bail(usage);
                        i++;
                        break;
                    case "-o":
                    case "-outfile":
                        path = args[i + 1];
                        i++;
                        break;
                    case "-x":
                        err = int.TryParse(args[i + 1], out startLeft);
                        if (!err) bail(usage);
                        i++;
                        break;
                    case "-y":
                        err = int.TryParse(args[i + 1], out startTop);
                        if (!err) bail(usage);
                        i++;
                        break;
                    case "-w": // fall through
                    case "-width":
                        err = int.TryParse(args[i + 1], out widthPixels);
                        if (!err) bail(usage);
                        i++;
                        break;
                    case "-h": // fall through
                    case "-height":
                        err = int.TryParse(args[i + 1], out heightPixels);
                        if (!err) bail(usage);
                        i++;
                        break;
                    case "-r": // fall through
                    case "-resolution":
                        err = int.TryParse(args[i + 1], out resolution);
                        if (!err) bail(usage);
                        i++;
                        break;
                    default:
                        bail(usage);
                        break;
                }
            }

            // Create a DeviceManager instance
            var deviceManager = new DeviceManager();

            // Create an empty variable to store the scanner instance
            DeviceInfo targetScanner = null;

            if (deviceId != -1)
                targetScanner = deviceManager.DeviceInfos[deviceId];
            else
            // Loop through the list of devices to choose the first available
            for (int i = 1; i <= deviceManager.DeviceInfos.Count; i++)
            {
                // Skip the device if it's not a scanner
                if (deviceManager.DeviceInfos[i].Type != WiaDeviceType.ScannerDeviceType)
                    continue;
                targetScanner = deviceManager.DeviceInfos[i];
                break;
            }

            if(targetScanner == null)
            {
                Console.WriteLine("Scanner not found");
                return -1;
            }

            // Connect to the first available scanner
            var device = targetScanner.Connect();

            // Select the scanner
            var scannerItem = device.Items[1];
            adjustScannerSettings(scannerItem, resolution, startLeft, startTop,
                            widthPixels, heightPixels, brightnessPct, contrastPct,
                            color_mode);
            Console.WriteLine("Scanning with " + targetScanner.Properties["Name"].get_Value() + " to " + path);
            Console.WriteLine($"  resolution: {resolution}  image size: {widthPixels},{heightPixels}");

            // Retrieve a image in JPEG format and store it into a variable
            var imageFile = (ImageFile)scannerItem.Transfer(FormatID.wiaFormatJPEG);

            // Save the image in some path with filename
            if (File.Exists(path))
            {
                File.Delete(path);
            }

            // Save image !
            imageFile.SaveFile(path);

            return 0;
        }

        private static void
        bail(string usage)
        {
            Console.WriteLine(usage);
            Environment.Exit(1);
        }

        /// <summary>
        /// Adjusts the settings of the scanner with the providen parameters.
        /// </summary>
        /// <param name="scannnerItem">Scanner Item</param>
        /// <param name="scanResolutionDPI">Provide the DPI resolution that should be used e.g 150</param>
        /// <param name="scanStartLeftPixel"></param>
        /// <param name="scanStartTopPixel"></param>
        /// <param name="scanWidthPixels"></param>
        /// <param name="scanHeightPixels"></param>
        /// <param name="brightnessPercents"></param>
        /// <param name="contrastPercents">Modify the contrast percent</param>
        /// <param name="colorMode">Set the color mode</param>
        private static void
        adjustScannerSettings(IItem scannnerItem, int scanResolutionDPI, int scanStartLeftPixel,
                        int scanStartTopPixel, int scanWidthPixels, int scanHeightPixels,
                        int brightnessPercents, int contrastPercents, int colorMode)
        {
            const string WIA_SCAN_COLOR_MODE = "6146";
            const string WIA_HORIZONTAL_SCAN_RESOLUTION_DPI = "6147";
            const string WIA_VERTICAL_SCAN_RESOLUTION_DPI = "6148";
            const string WIA_HORIZONTAL_SCAN_START_PIXEL = "6149";
            const string WIA_VERTICAL_SCAN_START_PIXEL = "6150";
            const string WIA_HORIZONTAL_SCAN_SIZE_PIXELS = "6151";
            const string WIA_VERTICAL_SCAN_SIZE_PIXELS = "6152";
            const string WIA_SCAN_BRIGHTNESS_PERCENTS = "6154";
            const string WIA_SCAN_CONTRAST_PERCENTS = "6155";
            setWIAProperty(scannnerItem.Properties, WIA_HORIZONTAL_SCAN_RESOLUTION_DPI, scanResolutionDPI);
            setWIAProperty(scannnerItem.Properties, WIA_VERTICAL_SCAN_RESOLUTION_DPI, scanResolutionDPI);
            setWIAProperty(scannnerItem.Properties, WIA_HORIZONTAL_SCAN_START_PIXEL, scanStartLeftPixel);
            setWIAProperty(scannnerItem.Properties, WIA_VERTICAL_SCAN_START_PIXEL, scanStartTopPixel);
            setWIAProperty(scannnerItem.Properties, WIA_HORIZONTAL_SCAN_SIZE_PIXELS, scanWidthPixels);
            setWIAProperty(scannnerItem.Properties, WIA_VERTICAL_SCAN_SIZE_PIXELS, scanHeightPixels);
            setWIAProperty(scannnerItem.Properties, WIA_SCAN_BRIGHTNESS_PERCENTS, brightnessPercents);
            setWIAProperty(scannnerItem.Properties, WIA_SCAN_CONTRAST_PERCENTS, contrastPercents);
            setWIAProperty(scannnerItem.Properties, WIA_SCAN_COLOR_MODE, colorMode);
        }

        /// <summary>
        /// Modify a WIA property
        /// </summary>
        /// <param name="properties"></param>
        /// <param name="propName"></param>
        /// <param name="propValue"></param>
        private static void
        setWIAProperty(IProperties properties, object propName, object propValue)
        {
            Property prop = properties.get_Item(ref propName);
            prop.set_Value(ref propValue);
        }

        private static void
        listScanningDevices()
        {
            var deviceManager = new DeviceManager();
            // Loop through the list of devices
            Console.WriteLine("Scanning for scanners");
            for (int i = 1; i <= deviceManager.DeviceInfos.Count; i++)
            {
                // Skip the device if it's not a scanner
                if (deviceManager.DeviceInfos[i].Type != WiaDeviceType.ScannerDeviceType)
                {
                    continue;
                }

                // Print something like e.g "WIA Canoscan 4400F"
                string info = (string) deviceManager.DeviceInfos[i].Properties["Name"].get_Value();
                Console.WriteLine($" {i} {info}");
                // e.g Canoscan 4400F
                //Console.WriteLine(deviceManager.DeviceInfos[i].Properties["Description"].get_Value());
                // e.g \\.\Usbscan0
                //Console.WriteLine(deviceManager.DeviceInfos[i].Properties["Port"].get_Value());
            } // Create a DeviceManager instance
        }
    }
}
