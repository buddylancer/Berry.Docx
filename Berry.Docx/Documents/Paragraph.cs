﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

using O = DocumentFormat.OpenXml;
using W = DocumentFormat.OpenXml.Wordprocessing;
using P = DocumentFormat.OpenXml.Packaging;
using A = DocumentFormat.OpenXml.Drawing;
using Pic = DocumentFormat.OpenXml.Drawing.Pictures;
using Wps = DocumentFormat.OpenXml.Office2010.Word.DrawingShape;
using Wpg = DocumentFormat.OpenXml.Office2010.Word.DrawingGroup;
using Wpc = DocumentFormat.OpenXml.Office2010.Word.DrawingCanvas;
using Dgm = DocumentFormat.OpenXml.Drawing.Diagrams;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using M = DocumentFormat.OpenXml.Math;
using V = DocumentFormat.OpenXml.Vml;
using Office = DocumentFormat.OpenXml.Vml.Office;

using Berry.Docx.Formatting;
using Berry.Docx.Field;
using Berry.Docx.Collections;
using Berry.Docx.Utils;

namespace Berry.Docx.Documents
{
    /// <summary>
    /// Represent the paragraph.
    /// </summary>
    public class Paragraph : DocumentItem
    {
        #region Private Members
        private readonly Document _doc;
        private readonly W.Paragraph _paragraph;
        private readonly ParagraphFormat _pFormat;
        private readonly CharacterFormat _cFormat;
        #endregion

        #region Constructors
        /// <summary>
        /// The paragraph constructor.
        /// </summary>
        /// <param name="doc">The owner document.</param>
        public Paragraph(Document doc) : this(doc, new W.Paragraph())
        {
        }

        internal Paragraph(Document doc, W.Paragraph paragraph) : base(doc, paragraph)
        {
            _doc = doc;
            _paragraph = paragraph;
            _pFormat = new ParagraphFormat(_doc, paragraph);
            _cFormat = new CharacterFormat();
            if(paragraph?.ParagraphProperties?.ParagraphMarkRunProperties != null)
                _cFormat = new CharacterFormat(_doc, paragraph.ParagraphProperties.ParagraphMarkRunProperties);
        }
        #endregion

        #region Public Properties
        /// <summary>
        /// The DocumentObject type.
        /// </summary>
        public override DocumentObjectType DocumentObjectType { get => DocumentObjectType.Paragraph; }

        /// <summary>
        ///Gets all child <see cref="ParagraphItem"/> of this paragraph.
        /// </summary>
        public ParagraphItemCollection ChildItems => new ParagraphItemCollection(_paragraph, ChildObjectsPrivate());

        /// <summary>
        /// Gets all child <see cref="DocumentObject"/> of this paragraph. 
        /// </summary>
        public override DocumentObjectCollection ChildObjects => new ParagraphItemCollection(_paragraph, ChildObjectsPrivate());

        

        /// <summary>
        /// The paragraph text.
        /// </summary>
        public string Text
        {
            get
            {
                StringBuilder text = new StringBuilder();
                foreach(DocumentObject item in ChildObjects)
                {
                    if(item is TextRange)
                    {
                        text.Append(((TextRange)item).Text);
                    }
                }
                return text.ToString();
            }
            set
            {
                _paragraph.RemoveAllChildren<W.Run>();
                W.Run run = RunGenerator.Generate(value);
                _paragraph.AddChild(run);
            }
        }

        /// <summary>
        /// The paragraph format.
        /// </summary>
        public ParagraphFormat Format => _pFormat;

        /// <summary>
        /// The character format of paragraph mark for this paragraph.
        /// </summary>
        public CharacterFormat CharacterFormat => _cFormat;

        /// <summary>
        /// The paragraph style. 
        /// </summary>
        public ParagraphStyle Style
        {
            get
            {
                if (_paragraph == null || _paragraph.GetStyle(_doc) == null) return null;
                return new ParagraphStyle(_doc, _paragraph.GetStyle(_doc));
            }
        }

        /// <summary>
        /// Gets the owener section of the current paragraph.
        /// </summary>
        public Section Section
        {
            get
            {
                foreach(Section section in _doc.Sections)
                {
                    if (section.Paragraphs.Contains(this))
                        return section;
                }
                return null;
            }
        }
        #endregion

        #region Public Methods
        /// <summary>
        /// Insert a section break with the specified type to the current paragraph. 
        /// <para>The current paragraph must have an owner section, otherwise an exception will be thrown.</para>
        /// </summary>
        /// <param name="type">Type of section break.</param>
        /// <exception cref="NullReferenceException"/>
        /// <returns>The section.</returns>
        public Section InsertSectionBreak(SectionBreakType type)
        {
            if (Section != null)
            {
                W.SectionProperties curSectPr = Section.XElement;
                // Clone a new SectionProperties from current section.
                W.SectionProperties newSectPr = (W.SectionProperties)curSectPr.CloneNode(true);
                // Set the current section type
                W.SectionType curSectType = curSectPr.Elements<W.SectionType>().FirstOrDefault();
                if (curSectType == null)
                {
                    curSectType = new W.SectionType();
                    curSectPr.AddChild(curSectType);
                }
                switch (type)
                {
                    case SectionBreakType.Continuous:
                        curSectType.Val = W.SectionMarkValues.Continuous;
                        break;
                    case SectionBreakType.OddPage:
                        curSectType.Val = W.SectionMarkValues.OddPage;
                        break;
                    case SectionBreakType.EvenPage:
                        curSectType.Val = W.SectionMarkValues.EvenPage;
                        break;
                    default:
                        curSectType.Remove();
                        break;
                }
                // Move current section to the next new paragraph if SectionProperties is present in current paragraph.
                if (_paragraph.Descendants<W.SectionProperties>().Any())
                {
                    W.Paragraph paragraph = new W.Paragraph() { ParagraphProperties = new W.ParagraphProperties() };
                    curSectPr.Remove();
                    paragraph.ParagraphProperties.AddChild(curSectPr);
                    _paragraph.InsertAfterSelf(paragraph);
                }
                // Add the new SectionProperties to the ParagraphProperties of the current paragraph
                if (_paragraph.ParagraphProperties == null)
                    _paragraph.ParagraphProperties = new W.ParagraphProperties();
                _paragraph.ParagraphProperties.AddChild(newSectPr);

                return new Section(_doc, newSectPr);
            }
            else
            {
                throw new NullReferenceException("The owner section of the current paragraph is null.");
            }
        }

        /// <summary>
        ///  Searches the paragraph for the first occurrence of the specified regular expression.
        /// </summary>
        /// <param name="pattern">The regular expression to search for a match</param>
        /// <returns>An object that contains information about the match.</returns>
        public TextMatch Find(Regex pattern)
        {
            Match match = pattern.Match(Text);
            if (match.Success)
            {
                return new TextMatch(this, match.Index, match.Index + match.Length - 1);
            }
            return null;
        }

        /// <summary>
        /// Searches the paragraph for all occurrences of a regular expression.
        /// </summary>
        /// <param name="pattern">The regular expression to search for a match</param>
        /// <returns>
        /// A list of the <see cref="TextMatch"/> objects found by the search.
        /// </returns>
        public List<TextMatch> FindAll(Regex pattern)
        {
            List<TextMatch> matches = new List<TextMatch>();
            foreach (Match match in pattern.Matches(Text))
            {
                if (match.Success)
                {
                    matches.Add(new TextMatch(this, match.Index, match.Index + match.Length - 1));
                }
            }
            return matches;
        }

        /// <summary>
        /// Appends a comment to the current paragraph.
        /// </summary>
        /// <param name="author">The author of the comment.</param>
        /// <param name="contents">The paragraphs content of the comment.</param>
        public void AppendComment(string author, params string[] contents)
        {
            int id = 0; // comment id
            P.WordprocessingCommentsPart part = _doc.Package.MainDocumentPart.WordprocessingCommentsPart;
            if (part == null)
            {
                part = _doc.Package.MainDocumentPart.AddNewPart<P.WordprocessingCommentsPart>();
                part.Comments = new W.Comments();
            }
            W.Comments comments = part.Comments;
            // max id + 1
            List<int> ids = new List<int>();
            foreach (W.Comment c in comments)
                ids.Add(c.Id.Value.ToInt());
            if (ids.Count > 0)
            {
                ids.Sort();
                id = ids.Last() + 1;
            }
            // comments content
            
            W.Comment comment = new W.Comment() { Id = id.ToString(), Author = author };
            foreach(string content in contents)
            {
                W.Paragraph paragraph = new W.Paragraph(new W.Run(new W.Text(content)));
                comment.Append(paragraph);
            }
            comments.Append(comment);
            // comment mark
            W.CommentRangeStart startMark = new W.CommentRangeStart() { Id = id.ToString() };
            W.CommentRangeEnd endMark = new W.CommentRangeEnd() { Id = id.ToString() };
            W.Run referenceRun = new W.Run(new W.CommentReference() { Id = id.ToString() });
            // Insert comment mark
            O.OpenXmlElement ele = _paragraph.FirstChild;
            if(ele is W.ParagraphProperties || ele is W.CommentRangeStart)
            {
                while(ele.NextSibling() != null && ele.NextSibling() is W.CommentRangeStart)
                {
                    ele = ele.NextSibling();
                }
                ele.InsertAfterSelf(startMark);
            }
            else
            {
                _paragraph.InsertAt(startMark, 0);
            }
            _paragraph.Append(endMark);
            _paragraph.Append(referenceRun);
        }

        /// <summary>
        /// Appends a comment to the current paragraph.
        /// </summary>
        /// <param name="author">The author of the comment.</param>
        /// <param name="contents">The paragraphs content of the comment.</param>
        public void AppendComment(string author, IEnumerable<string> contents)
        {
            int id = 0; // comment id
            P.WordprocessingCommentsPart part = _doc.Package.MainDocumentPart.WordprocessingCommentsPart;
            if (part == null)
            {
                part = _doc.Package.MainDocumentPart.AddNewPart<P.WordprocessingCommentsPart>();
                part.Comments = new W.Comments();
            }
            W.Comments comments = part.Comments;
            // max id + 1
            List<int> ids = new List<int>();
            foreach (W.Comment c in comments)
                ids.Add(c.Id.Value.ToInt());
            if (ids.Count > 0)
            {
                ids.Sort();
                id = ids.Last() + 1;
            }
            // comments content
            W.Comment comment = new W.Comment() { Id = id.ToString(), Author = author };
            foreach (string content in contents)
            {
                W.Paragraph paragraph = new W.Paragraph(new W.Run(new W.Text(content)));
                comment.Append(paragraph);
            }
            comments.Append(comment);
            // comment mark
            W.CommentRangeStart startMark = new W.CommentRangeStart() { Id = id.ToString() };
            W.CommentRangeEnd endMark = new W.CommentRangeEnd() { Id = id.ToString() };
            W.Run referenceRun = new W.Run(new W.CommentReference() { Id = id.ToString() });
            // Insert comment mark
            O.OpenXmlElement ele = _paragraph.FirstChild;
            if (ele is W.ParagraphProperties || ele is W.CommentRangeStart)
            {
                while (ele.NextSibling() != null && ele.NextSibling() is W.CommentRangeStart)
                {
                    ele = ele.NextSibling();
                }
                ele.InsertAfterSelf(startMark);
            }
            else
            {
                _paragraph.InsertAt(startMark, 0);
            }
            _paragraph.Append(endMark);
            _paragraph.Append(referenceRun);
        }

        #endregion

        #region Private Methods
        private IEnumerable<ParagraphItem> ChildObjectsPrivate()
        {
            foreach (O.OpenXmlElement ele in _paragraph.ChildElements)
            {
                if (ele is W.Run)
                {
                    W.Run run = (W.Run)ele;
                    // text range
                    if(run.Elements<W.Text>().Any())
                        yield return new TextRange(_doc, run);
                    
                    // footnote reference
                    if(run.Elements<W.FootnoteReference>().Any())
                    {
                        yield return new FootnoteReference(_doc, run, run.Elements<W.FootnoteReference>().First());
                    }
                    // endnote reference
                    if (run.Elements<W.EndnoteReference>().Any())
                    {
                        yield return new EndnoteReference(_doc, run, run.Elements<W.EndnoteReference>().First());
                    }
                    // picture
                    foreach (W.Drawing drawing in run.Descendants<W.Drawing>())
                    {
                        A.GraphicData graphicData = drawing.Descendants<A.GraphicData>().FirstOrDefault();
                        if(graphicData != null)
                        {
                            if(graphicData.FirstChild is Pic.Picture)
                                yield return new Picture(_doc, run, drawing);
                            else if(graphicData.FirstChild is Wps.WordprocessingShape)
                                yield return new Shape(_doc, run, drawing);
                            else if (graphicData.FirstChild is Wpg.WordprocessingGroup)
                                yield return new GroupShape(_doc, run, drawing);
                            else if (graphicData.FirstChild is Wpc.WordprocessingCanvas)
                                yield return new Canvas(_doc, run, drawing);
                            else if (graphicData.FirstChild is Dgm.RelationshipIds)
                                yield return new Diagram(_doc, run, drawing);
                            else if (graphicData.FirstChild is C.ChartReference)
                                yield return new Chart(_doc, run, drawing);
                        }
                    }
                    // embedded object
                    foreach(W.EmbeddedObject obj in run.Elements<W.EmbeddedObject>())
                    {
                        yield return new EmbeddedObject(_doc, run, obj);
                    }
                }
                if(ele is W.Hyperlink)
                {
                    foreach (O.OpenXmlElement e in ele.ChildElements)
                    {
                        if (e is W.Run)
                        {
                            W.Run run = (W.Run)e;
                            if (run.Elements<W.Text>().Any())
                                yield return new TextRange(_doc, run);
                        }
                    }
                }
                    
            }
        }
        #endregion

        #region TODO

        /// <summary>
        /// 段落编号(默认为1)
        /// </summary>
        public string ListText
        {
            get
            {
                if (_pFormat.NumberingFormat == null) return string.Empty;
                string lvlText = _pFormat.NumberingFormat.Format;
                //Console.WriteLine($"{lvlText},{_pFormat.NumberingFormat.Style}");
                if (_pFormat.NumberingFormat.Style == W.NumberFormatValues.Decimal)
                    lvlText = lvlText.RxReplace(@"%[0-9]", "1");
                else if (_pFormat.NumberingFormat.Style == W.NumberFormatValues.ChineseCounting
                    || _pFormat.NumberingFormat.Style == W.NumberFormatValues.ChineseCountingThousand
                    || _pFormat.NumberingFormat.Style == W.NumberFormatValues.JapaneseCounting)
                    lvlText = lvlText.RxReplace(@"%[0-9]", "一");

                return lvlText;
            }
        }

        private FieldCodeCollection FieldCodes
        {
            get
            {
                List<FieldCode> fieldcodes = new List<FieldCode>();
                List<O.OpenXmlElement> childElements = new List<O.OpenXmlElement>();

                int begin_times = 0;
                int end_times = 0;

                foreach (O.OpenXmlElement ele in _paragraph.Descendants())
                {
                    if (ele.GetType().FullName.Equals("DocumentFormat.OpenXml.Wordprocessing.SimpleField"))
                    {
                        fieldcodes.Add(new FieldCode((W.SimpleField)ele));
                    }
                    else if (ele.GetType().FullName.Equals("DocumentFormat.OpenXml.Wordprocessing.Run"))
                    {
                        W.Run run = (W.Run)ele;
                        if (run.Elements<W.FieldChar>().Any() && run.Elements<W.FieldChar>().First().FieldCharType != null)
                        {
                            string field_type = run.Elements<W.FieldChar>().First().FieldCharType.ToString();
                            if (field_type == "begin") begin_times++;
                            else if (field_type == "end") end_times++;
                        }
                        if (begin_times > 0)
                        {
                            childElements.Add(ele);
                            if (end_times == begin_times)
                            {
                                fieldcodes.Add(new FieldCode(childElements));
                                begin_times = 0;
                                end_times = 0;
                                childElements.Clear();
                            }
                        }
                    }
                }
                return new FieldCodeCollection(fieldcodes);
            }
        }
        #endregion
    }
}
