﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using O = DocumentFormat.OpenXml;
using W = DocumentFormat.OpenXml.Wordprocessing;

using Berry.Docx.Collections;
using Berry.Docx.Formatting;

namespace Berry.Docx.Documents
{
    /// <summary>
    /// Represent the table.
    /// </summary>
    public class Table : DocumentItem
    {
        #region Private Members
        private readonly Document _doc;
        private readonly W.Table _table;
        #endregion

        #region Constructors
        /// <summary>
        /// The table constructor.
        /// </summary>
        /// <param name="doc">The owner document.</param>
        /// <param name="rowCnt">Table row count.</param>
        /// <param name="columnCnt">Table column count.</param>
        public Table(Document doc, int rowCnt, int columnCnt)
            : this(doc, TableGenerator.GenerateTable(rowCnt, columnCnt))
        {

        }

        internal Table(Document doc, W.Table table) : base(doc, table)
        {
            _doc = doc;
            _table = table;
        }
        #endregion

        #region Public Properties
        /// <summary>
        /// Gets the row count.
        /// </summary>
        public int RowCount { get { return Rows.Count; } }

        /// <summary>
        /// Gets the column count.
        /// </summary>
        public int ColumnCount
        {
            get
            {
                int maxCnt = 0;
                foreach(var row in Rows)
                {
                    maxCnt = Math.Max(maxCnt, row.Cells.Count);
                }
                return maxCnt;
            }
        }

        /// <summary>
        /// Gets the table format.
        /// </summary>
        public TableFormat Format { get { return new TableFormat(_doc, this); } }

        /// <summary>
        /// The DocumentObject type.
        /// </summary>
        public override DocumentObjectType DocumentObjectType { get { return DocumentObjectType.Table; } }

        /// <summary>
        /// The child DocumentObjects of this table.
        /// </summary>
        public override DocumentObjectCollection ChildObjects { get { return Rows; } }

        /// <summary>
        /// Gets the table row at the specified row.
        /// </summary>
        /// <param name="row">The zero-based index.</param>
        /// <returns>The table row at the specified index.</returns>
        public TableRow this[int row] { get { return Rows[row]; } }

        /// <summary>
        /// The table rows collection.
        /// </summary>
        public TableRowCollection Rows { get { return new TableRowCollection(_table, TableRowsPrivate()); } }

        /// <summary>
        /// Gets the width of columns.
        /// </summary>
        public ColumnWidthCollection ColumnWidths { get { return new ColumnWidthCollection(_doc, this); } }
        #endregion

        #region Public Methods
        /// <summary>
        /// Adds a new row to the end of table.
        /// </summary>
        /// <returns>The table row.</returns>
        public TableRow AddRow()
        {
            TableRow row = (TableRow)Rows.Last().Clone();
            row.ClearContent();
            Rows.Add(row);
            return row;
        }

        /// <summary>
        /// Gets the table style.
        /// </summary>
        /// <returns>The table style.</returns>
        public TableStyle GetStyle()
        {
            W.Styles styles = _doc.Package.MainDocumentPart.StyleDefinitionsPart.Styles;
            W.TableProperties tblPr = _table.GetFirstChild<W.TableProperties>();
            if (tblPr.TableStyle.Val != null)
            {
                string styleId = tblPr.TableStyle.Val.ToString();
                W.Style style = styles.Elements<W.Style>().Where(s => s.StyleId == styleId).FirstOrDefault();
                if(style != null)
                {
                    return new TableStyle(_doc, style);
                }
            }
            return TableStyle.Default(_doc);
        }

        /// <summary>
        /// Applies the table style.
        /// </summary>
        /// <param name="styleName">The table style name.</param>
        public void ApplyStyle(string styleName)
        {
            if (string.IsNullOrEmpty(styleName)) return;
            var style = _doc.Styles.FindByName(styleName, StyleType.Table);
            if (style == null)
            {
                style = new TableStyle(_doc, styleName);
                _doc.Styles.Add(style);
            }
            ApplyStyle(style as TableStyle);
        }

        /// <summary>
        /// Applies the table style.
        /// </summary>
        /// <param name="style"></param>
        public void ApplyStyle(TableStyle style)
        {
            if (!_table.Elements<W.TableProperties>().Any())
            {
                _table.AddChild(new W.TableProperties());
            }
            W.TableProperties tblPr = _table.GetFirstChild<W.TableProperties>();
            tblPr.TableStyle = new W.TableStyle() { Val = style.StyleId };
        }

        /// <summary>
        /// Automatically resize the table columns.
        /// </summary>
        /// <param name="method">The resize method.</param>
        public void AutoFit(AutoFitMethod method)
        {
            if (_table.GetFirstChild<W.TableProperties>() == null)
            {
                _table.AddChild(new W.TableProperties());
            }
            W.TableProperties tblPr = _table.GetFirstChild<W.TableProperties>();
            if (method == AutoFitMethod.AutoFitContents)
            {
                tblPr.TableWidth = new W.TableWidth() { Width = "0", Type = W.TableWidthUnitValues.Auto };
                tblPr.TableLayout = null;
            }
            else if (method == AutoFitMethod.AutoFitWindow)
            {
                tblPr.TableWidth = new W.TableWidth() { Width = "5000", Type = W.TableWidthUnitValues.Pct };
                tblPr.TableLayout = null;
            }
            else
            {
                tblPr.TableWidth = new W.TableWidth() { Width = "0", Type = W.TableWidthUnitValues.Auto };
                tblPr.TableLayout = new W.TableLayout() { Type = W.TableLayoutValues.Fixed };
            }
        }

        /// <summary>
        /// Sets the width of the specified column.
        /// </summary>
        /// <param name="colIndex">The zero-based column index.</param>
        /// <param name="width">The column width.</param>
        /// <param name="cellWidthType">The measurement type of the width.  </param>
        /// <exception cref="ArgumentOutOfRangeException"></exception>
        public void SetColumnWidth(int colIndex, float width, CellWidthType cellWidthType)
        {
            if(colIndex < 0 || colIndex >= ColumnCount)
            {
                throw new ArgumentOutOfRangeException("Invalid table column index.");
            }
            if (_table.GetFirstChild<W.TableGrid>() == null)
            {
                W.TableGrid grid = new W.TableGrid();
                for(int i = 0; i < ColumnCount; i++)
                {
                    grid.Append(new W.GridColumn() { Width = "222" });
                }
                _table.AddChild(grid);
            }
            W.TableGrid tblGrid = _table.GetFirstChild<W.TableGrid>();
            float totalWidth = _doc.LastSection.PageSetup.PageWidth - _doc.LastSection.PageSetup.LeftMargin - _doc.LastSection.PageSetup.RightMargin;
            int w = 0;
            if(cellWidthType == CellWidthType.Percent)
            {
                w = (int)Math.Round(totalWidth * width * 20 / 100.0F);
            }
            else
            {
                w = (int)Math.Round(width * 20);
            }
            tblGrid.Elements<W.GridColumn>().ElementAt(colIndex).Width = w.ToString();
            foreach (TableRow row in Rows)
            {
                row.Cells[colIndex].SetCellWidth(width, cellWidthType);
            }
        }
        #endregion

        #region Internal
        internal new W.Table XElement { get { return _table; } }
        #endregion

        #region Private Methods
        private IEnumerable<TableRow> TableRowsPrivate()
        {
            foreach (W.TableRow row in _table.Elements<W.TableRow>())
            {
                yield return new TableRow(_doc, row);
            }
        }
        #endregion
    }
}
