﻿using System;
using System.Collections.Generic;
using System.Text;
using System.Linq;
using System.Windows;
using Berry.Docx.Visual.Field;
using BF = Berry.Docx.Field;

namespace Berry.Docx.Visual.Documents
{
    public class Paragraph : DocumentItem
    {
        #region Private Members
        private Berry.Docx.Documents.Paragraph _paragraph;
        private readonly double _width = 0;
        private double _charSpace = 0;
        private double _lineSpace = 0;
        private DocGridType _gridType;
        
        private double _leftIndent = 0;
        private double _rightIndent = 0;
        private double _specialIndent = 0;
        private double _beforeSpace = 0;
        private double _afterSpace = 0;

        private Margin _margin = new Margin(0, 0, 0, 0);
        private Margin _padding = new Margin(0, 0, 0, 0);

        private double _normalFontSize = 10.5;

        private List<ParagraphLine> _lines;
        #endregion

        #region Constructor
        internal Paragraph(Berry.Docx.Documents.Paragraph paragraph, double width, double charSpace, double lineSpace, Berry.Docx.DocGridType gridType)
        {
            _paragraph = paragraph;
            _width = width;
            _charSpace = charSpace;
            _lineSpace = lineSpace;
            _gridType = gridType;

            _lines = new List<ParagraphLine>();

            float firstCharSize = 0;
            if (paragraph.ChildItems.Count > 0 && 
                (paragraph.ChildItems[0].DocumentObjectType == DocumentObjectType.TextRange
                || paragraph.ChildItems[0].DocumentObjectType == DocumentObjectType.Picture))
            {
                firstCharSize = paragraph.ChildItems[0].CharacterFormat.FontSize;
            }
            else
            {
                // 空行
                firstCharSize = paragraph.MarkFormat.FontSize;
                paragraph.ChildItems.InsertAt(new Berry.Docx.Field.TextRange(_paragraph.Document, " "), 0);
            }

            _normalFontSize = Berry.Docx.Formatting.ParagraphStyle.Default(paragraph.Document).CharacterFormat.FontSize;

            #region 缩进
            // 缩进
            var leftInd = paragraph.Format.GetLeftIndent();
            var rightInd = paragraph.Format.GetRightIndent();
            var specialInd = paragraph.Format.GetSpecialIndentation();
            // 左侧缩进
            if (leftInd.Unit == IndentationUnit.Character)
            {
                if (gridType == DocGridType.LinesAndChars)
                {
                    _leftIndent = charSpace * leftInd.Val;
                }
                else if (gridType == DocGridType.SnapToChars)
                {
                    _leftIndent = charSpace * Math.Ceiling(leftInd.Val);
                }
                else
                {
                    _leftIndent = (_normalFontSize * leftInd.Val) / 72 * 96;
                }
            }
            else
            {
                _leftIndent = leftInd.Val / 72 * 96;
                if (gridType == DocGridType.SnapToChars)
                {
                    int cnt = 1;
                    while (_leftIndent > charSpace * cnt) cnt++;
                    _leftIndent = charSpace * cnt;
                }
            }
            // 右侧缩进
            if (rightInd.Unit == IndentationUnit.Character)
            {
                if (gridType == DocGridType.LinesAndChars
                    || gridType == DocGridType.SnapToChars)
                {
                    _rightIndent = charSpace * rightInd.Val;
                }
                else if (gridType == DocGridType.SnapToChars)
                {
                    _rightIndent = charSpace * Math.Ceiling(rightInd.Val);
                }
                else
                {
                    _rightIndent = (_normalFontSize * rightInd.Val) / 72 * 96;
                }
            }
            else
            {
                _rightIndent = rightInd.Val / 72 * 96;
                if (gridType == DocGridType.SnapToChars)
                {
                    int cnt = 1;
                    while (_rightIndent > charSpace * cnt) cnt++;
                    _rightIndent = charSpace * cnt;
                }
            }

            // 特殊缩进
            if (gridType == DocGridType.LinesAndChars)
            {
                if (specialInd.Type == SpecialIndentationType.FirstLine)
                {
                    if (specialInd.Unit == IndentationUnit.Character)
                        _specialIndent = (charSpace + (firstCharSize - _normalFontSize) / 72 * 96) * specialInd.Val;
                    else
                        _specialIndent = specialInd.Val / 72 * 96;
                }
                else if (specialInd.Type == SpecialIndentationType.Hanging)
                {
                    if (specialInd.Unit == IndentationUnit.Character)
                        _specialIndent = -(charSpace + (firstCharSize - _normalFontSize) / 72 * 96) * specialInd.Val;
                    else
                        _specialIndent = -specialInd.Val / 72 * 96;
                }
            }
            else if (gridType == DocGridType.SnapToChars)
            {
                if (specialInd.Type == SpecialIndentationType.FirstLine)
                {
                    if (specialInd.Unit == IndentationUnit.Character)
                        _specialIndent = Math.Ceiling((charSpace + (firstCharSize - _normalFontSize) / 72 * 96) * specialInd.Val / charSpace) * charSpace;
                    else
                        _specialIndent = Math.Ceiling(specialInd.Val / 72 * 96 / charSpace) * charSpace;
                }
                else if (specialInd.Type == SpecialIndentationType.Hanging)
                {
                    if (specialInd.Unit == IndentationUnit.Character)
                        _specialIndent = -Math.Ceiling((charSpace + (firstCharSize - _normalFontSize) / 72 * 96) * specialInd.Val / charSpace) * charSpace;
                    else
                        _specialIndent = -Math.Ceiling(specialInd.Val / 72 * 96 / charSpace) * charSpace;
                }
            }
            else
            {
                if (specialInd.Type == SpecialIndentationType.FirstLine)
                {
                    if (specialInd.Unit == IndentationUnit.Character)
                        _specialIndent = firstCharSize / 72 * 96 * specialInd.Val;
                    else
                        _specialIndent = specialInd.Val / 72 * 96;
                }
                else if (specialInd.Type == SpecialIndentationType.Hanging)
                {
                    if (specialInd.Unit == IndentationUnit.Character)
                        _specialIndent = -firstCharSize / 72 * 96 * specialInd.Val;
                    else
                        _specialIndent = -specialInd.Val / 72 * 96;
                }
            }
            #endregion

            #region 间距
            // 段前段后间距
            var beforeSpacing = paragraph.Format.GetBeforeSpacing();
            var afterSpacing = paragraph.Format.GetAfterSpacing();
            if (gridType == DocGridType.None)
            {
                if (beforeSpacing.Unit == SpacingUnit.Line)
                    _beforeSpace = beforeSpacing.Val * 12f.ToPixel();
                else
                    _beforeSpace = beforeSpacing.Val.ToPixel();

                if (afterSpacing.Unit == SpacingUnit.Line)
                    _afterSpace = afterSpacing.Val * 12f.ToPixel();
                else
                    _afterSpace = afterSpacing.Val.ToPixel();
            }
            else
            {
                if (beforeSpacing.Unit == SpacingUnit.Line)
                    _beforeSpace = beforeSpacing.Val * lineSpace;
                else
                    _beforeSpace = beforeSpacing.Val.ToPixel();

                if (afterSpacing.Unit == SpacingUnit.Line)
                    _afterSpace = afterSpacing.Val * lineSpace;
                else
                    _afterSpace = afterSpacing.Val.ToPixel();
            }
            #endregion

            _margin = new Margin(0, _beforeSpace, 0, _afterSpace);
            _padding = new Margin(_leftIndent, 0, _rightIndent, 0);
        }
        #endregion

        #region Public Properties
        public double Width { get { return _width; } }

        public double Height
        {
            get
            {
                double height = 0;
                foreach(var line in _lines)
                {
                    height += line.Margin.Top + line.Height + line.Margin.Bottom;
                }
                return height;
            }
        }

        public Margin Margin { get { return _margin; } }

        public Margin Padding { get { return _padding; } }

        public List<ParagraphLine> Lines { get { return _lines; } }
        #endregion

        #region Internal Methods
        internal List<ParagraphLine> GenerateLines()
        {
            List<ParagraphLine> lines = new List<ParagraphLine>();
            int index = 0;
            double width = _width - _margin.Left - _margin.Right - _padding.Left - _padding.Right;
            lines.Add(new ParagraphLine(_paragraph, width, _charSpace, _lineSpace, _gridType));
            if (_specialIndent > 0) lines[index].Padding.Left = _specialIndent;
            // 段落编号
            if (_paragraph.ListFormat.CurrentLevel != null
                &&!string.IsNullOrEmpty(_paragraph.ListText))
            {
                string list = _paragraph.ListText;
                if (_paragraph.ListFormat.CurrentLevel.SuffixCharacter == LevelSuffixCharacter.Tab)
                    list += "    ";
                else if (_paragraph.ListFormat.CurrentLevel.SuffixCharacter == LevelSuffixCharacter.Space)
                    list += " ";
                foreach(var c in list)
                {
                    var c1 = new Docx.Field.Character(c, _paragraph.MarkFormat);
                    Character character = new Character(c1, _charSpace, _normalFontSize, _gridType);
                    while (!lines[index].TryAppend(character))
                    {
                        var line = new ParagraphLine(_paragraph, width, _charSpace, _lineSpace, _gridType);
                        if (_specialIndent < 0) line.Padding.Left = Math.Abs(_specialIndent);
                        lines.Add(line);
                        index++;
                    }
                }
            }

            foreach (var item in _paragraph.ChildItems)
            {
                // 文本
                if (item is Berry.Docx.Field.TextRange)
                {
                    var tr = (Berry.Docx.Field.TextRange)item;
                    foreach (var c in tr.Characters)
                    {
                        Character character = new Character(c, _charSpace, _normalFontSize, _gridType);
                        while (!lines[index].TryAppend(character))
                        {
                            var line = new ParagraphLine(_paragraph, width, _charSpace, _lineSpace, _gridType);
                            if (_specialIndent < 0) line.Padding.Left = Math.Abs(_specialIndent);
                            lines.Add(line);
                            index++;
                        }
                    }
                }
                // 图片
                else if(item is Berry.Docx.Field.Picture)
                {
                    Picture picture = new Picture((Berry.Docx.Field.Picture)item);
                    while (!lines[index].TryAppend(picture))
                    {
                        var line = new ParagraphLine(_paragraph, width, _charSpace, _lineSpace, _gridType);
                        if (_specialIndent < 0) line.Padding.Left = Math.Abs(_specialIndent);
                        lines.Add(line);
                        index++;
                    }
                }
                // 分页符
                else if (item is Berry.Docx.Field.Break)
                {
                    var br = (Berry.Docx.Field.Break)item;
                    if (br.Type == BreakType.Page)
                    {
                        lines[index].EndsWithPageBreak = true;
                        var line = new ParagraphLine(_paragraph, width, _charSpace, _lineSpace, _gridType);
                        if (item == _paragraph.ChildItems.First())
                        {
                            if (_specialIndent > 0) line.Padding.Left = _specialIndent;
                        }
                        else if (item == _paragraph.ChildItems.Last())
                        {
                            continue;
                        }
                        else
                        {
                            if (_specialIndent < 0) line.Padding.Left = Math.Abs(_specialIndent);
                        }
                        lines.Add(line);
                        index++;
                    }
                }

            }
            return lines;
        }
        #endregion
    }
}
