using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDrawIo
{
    static class ImagePropertiesHelper
    {
        private const int CommentPropertyItemId = 37510;

        public static void RemoveComment(Image img)
        {
            if (img == null || img.PropertyItems.All(t => t.Id != CommentPropertyItemId))
                return;

            img.RemovePropertyItem(CommentPropertyItemId);
        }

        public static void SetComment(Image img, byte[] bytes)
        {
            if (img == null)
                return;

            Image tmpImg;
            using (var stream = Helpers.GetResourceStream("Resources.new.png"))
                tmpImg = Image.FromStream(stream);
            var propItem = tmpImg.GetPropertyItem(CommentPropertyItemId);

            //var bytes = Encoding.UTF8.GetBytes(data);
            var base64 = Convert.ToBase64String(bytes);
            var asciiBytes = Encoding.ASCII.GetBytes(base64 + char.MinValue);

            propItem.Value = asciiBytes;
            propItem.Len = bytes.Length;

            img.SetPropertyItem(propItem);
        }

        public static byte[] GetComment(Image img)
        {
            if (img == null || img.PropertyItems.All(t => t.Id != CommentPropertyItemId))
                return null;

            var propItem = img.GetPropertyItem(CommentPropertyItemId);
            var value = propItem.Value;

            var base64 = Encoding.ASCII.GetString(value, 0, value.Length - 1);
            var bytes = Convert.FromBase64String(base64);

            //var data = Encoding.UTF8.GetString(bytes);

            return bytes;
        }
    }
}
