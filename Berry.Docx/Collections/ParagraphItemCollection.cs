﻿using System;
using System.Linq;
using System.Collections.Generic;
using O = DocumentFormat.OpenXml;
using Berry.Docx.Field;

namespace Berry.Docx.Collections
{
    /// <summary>
    /// Represent a ParagraphItem collection.
    /// </summary>
    public class ParagraphItemCollection : DocumentItemCollection, IEnumerable<ParagraphItem>
    {
        #region Private Members
        private readonly O.OpenXmlElement _owner;
        private readonly IEnumerable<ParagraphItem> _items;
        #endregion

        #region Constructors
#if NET35
        internal ParagraphItemCollection(O.OpenXmlElement owner, IEnumerable<ParagraphItem> items) : base(owner, items.Convert())
#else
        internal ParagraphItemCollection(O.OpenXmlElement owner, IEnumerable<ParagraphItem> items) : base(owner, items)
#endif
        {
            _owner = owner;
            _items = items;
        }
        #endregion

        /// <summary>
        /// Gets the paragraph child item at the specified index.
        /// </summary>
        /// <param name="index">The zero-based index.</param>
        /// <returns>The paragraph child item at the specified index.</returns>
        public new ParagraphItem this[int index] { get { return (ParagraphItem)base[index]; } }

        public override void Add(DocumentObject obj)
        {
            ParagraphItem item = obj as ParagraphItem;
            if (item == null)
            {
                throw new InvalidCastException("{obj.DocumentObjectType} is not a ParagraphItem!");
            }
            Add(item);
        }

        public void Add(ParagraphItem item)
        {
            if (item.InsideRun)
                _owner.Append(item.OwnerRun);
            else
                _owner.Append(item.XElement);
        }

        public override void InsertAt(DocumentObject obj, int index)
        {
            ParagraphItem item = obj as ParagraphItem;
            if (item == null)
            {
                throw new InvalidCastException("{obj.DocumentObjectType} is not a ParagraphItem!");
            }
            InsertAt(item, index);
        }

        public void InsertAt(ParagraphItem item, int index)
        {
            if (index == _items.Count())
            {
                Add(item);
            }
            else
            {
                _items.ElementAt(index).InsertBeforeSelf(item);
            }
        }

        public IEnumerator<ParagraphItem> GetEnumerator()
        {
            return _items.GetEnumerator();
        }
    }
}
