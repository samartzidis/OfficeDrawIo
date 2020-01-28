using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Word = Microsoft.Office.Interop.Word;

namespace OfficeDrawIoWord
{
    public class ShapeHolder
    {
        private readonly dynamic _shape;

        public ShapeHolder(Word.InlineShape shape)
        {
            _shape = shape;
        }

        public ShapeHolder(Word.Shape shape)
        {
            _shape = shape;
        }

        public void Delete()
        {
            _shape.Delete();
        }

        public Word.InlineShape InlineShape => _shape as Word.InlineShape;
        public Word.Shape Shape => _shape as Word.Shape;
        public bool IsInlineShape => _shape is Word.InlineShape;
        public string Title => _shape.Title;
        public int AnchorID => _shape.AnchorID;
        public bool Visible
        {
            get => _shape.Visible;
            set => _shape.Visible = value;
        }
    }
}
