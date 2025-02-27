﻿using Berry.Docx.Collections;
using System;
using System.Collections.Generic;
using System.Text;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace Berry.Docx.Documents
{
    public class SdtBlockContent : DocumentItem
    {
        internal SdtBlockContent(Document doc, W.SdtContentBlock sdt) : base(doc, sdt)
        {
        }

        public override DocumentObjectType DocumentObjectType { get { return DocumentObjectType.SdtBlockContent; } }
    }
}
